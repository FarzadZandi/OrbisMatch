#!/usr/bin/env python3
"""
Resume-safe Orbis Batch Search candidate scraper.

This script intentionally does not automate authentication. It launches a headed
persistent Playwright browser, navigates to the configured TUM/Orbis entry URL,
and waits until the human types "ready" in the terminal.
"""

from __future__ import annotations

import argparse
import json
import logging
import random
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlparse

import openpyxl
from matching_logic import COUNTRY_TO_ISO2, MIN_WITNESSES, has_convincing_candidate, select_best_match
from playwright.sync_api import (
    ElementHandle,
    Locator,
    Page,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)


DEFAULT_START_URL = "https://eaccess.tum.edu/login?qurl=https%3A%2F%2Forbis4.bvdinfo.com%2Fip"
DEFAULT_BATCH_SEARCH_URL = (
    "https://orbis-r1-bvdinfo-com.tum-eaccess.de/version-20260316-7-3/"
    "Orbis/1/Companies/BatchSearch/Start"
)

RESULT_TABLE_SELECTORS = [
    "select#pageSize",
    "div[title][data-id]",
    "table:has(select#pageSize)",
    "[data-id='Name']",
]

CANDIDATE_NAME_SELECTOR = "div[data-id='Name'][data-snapshot-bvdid]"
NO_RESULT_SELECTOR = "td.noResult"
COMPANY_DROPDOWN_SELECTOR = "img[role='button'][aria-expanded][src*='Icons/Orbis/dropdown']"
RUNTIME_HARD_BUFFER_MINUTES = 15
RATE_BLOCK_TEXT_PATTERNS = [
    "too many requests",
    "temporarily blocked",
    "verify you are human",
    "verify that you are human",
    "we have detected unusual traffic",
    "your computer or network may be sending automated queries",
]


class StopRun(RuntimeError):
    def __init__(self, message: str, stop_reason: str = "systemic safety stop") -> None:
        super().__init__(message)
        self.stop_reason = stop_reason


@dataclass(frozen=True)
class UploadRow:
    own_id: str
    company: str
    values: list[Any]
    source_file: str
    source_row: int
    reference_country: str = ""
    reference_city: str = ""


@dataclass(frozen=True)
class Batch:
    key: str
    source_file: Path
    upload_file: Path
    rows: list[UploadRow]


class OrbisScraper:
    def __init__(self, args: argparse.Namespace) -> None:
        self.args = args
        self.root = Path(args.run_dir).resolve()
        self.temp_upload_dir = self.root / "temp_uploads"
        self.screenshot_dir = self.root / "screenshots"
        self.log_dir = self.root / "logs"
        self.result_path = Path(args.output).resolve()
        self.checkpoint_path = Path(args.checkpoint).resolve()
        self.completed_ids = load_completed_ids(self.checkpoint_path)
        self.seen_result_keys = load_seen_result_keys(self.result_path)
        self.selected_bvdid_by_own_id = load_selected_bvdids(self.result_path)
        self.reference_by_id = load_reference_rows(Path(args.reference))
        self.companies_attempted_this_run = 0
        self.orbis_auto_match_count = 0
        self.consecutive_company_full_failures = 0
        self.last_company_counts_as_collection_failure = False
        self.last_block_response_reason = ""
        self.result_page_search_limit = 5
        self.runtime_started_at: float | None = None
        self.max_companies_reached_at: float | None = None
        self.last_finished_own_id = ""
        self.stop_reason = ""

        for directory in [
            self.root,
            self.temp_upload_dir,
            self.screenshot_dir,
            self.log_dir,
            self.result_path.parent,
            self.checkpoint_path.parent,
        ]:
            directory.mkdir(parents=True, exist_ok=True)

        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(message)s",
            handlers=[
                logging.FileHandler(self.log_dir / "orbis_scraper.log", encoding="utf-8"),
                logging.StreamHandler(sys.stdout),
            ],
        )
        logging.info("Loaded %s reference rows from %s.", len(self.reference_by_id), Path(args.reference))
        logging.info(
            "Human interaction pause range: %.2f-%.2f seconds.",
            self.args.interaction_min_wait,
            self.args.interaction_max_wait,
        )
        logging.info(
            "Between-company pause range: 12.00-30.00 seconds; distraction pause: 5%% chance of an extra 20-60 seconds; periodic company-count rest prompt is disabled.",
        )
        if self.args.max_runtime_minutes is not None:
            logging.info(
                "Runtime cap: %s minutes soft limit; %s minutes hard fallback (%s-minute fixed buffer).",
                self.args.max_runtime_minutes,
                self.args.max_runtime_minutes + RUNTIME_HARD_BUFFER_MINUTES,
                RUNTIME_HARD_BUFFER_MINUTES,
            )

    def run(self) -> None:
        batches = self.build_batches()
        if self.args.dry_run:
            logging.info("Dry run: %s pending batches.", len(batches))
            display_limit = self.args.max_batches or 20
            for batch in batches[:display_limit]:
                logging.info("%s -> %s rows -> %s", batch.key, len(batch.rows), batch.upload_file)
            if len(batches) > display_limit:
                logging.info("... %s more batches omitted from dry-run listing.", len(batches) - display_limit)
            return

        with sync_playwright() as playwright:
            context = playwright.chromium.launch_persistent_context(
                user_data_dir=str(Path(self.args.user_data_dir).resolve()),
                headless=False,
                accept_downloads=True,
                slow_mo=self.args.slow_mo,
                viewport={"width": 1440, "height": 1000},
            )
            page = context.pages[0] if context.pages else context.new_page()
            page.set_default_timeout(self.args.default_timeout_ms)
            page.set_default_navigation_timeout(self.args.navigation_timeout_ms)
            page.on("response", self.note_blocking_response)

            close_context = True
            try:
                if self.args.start_from_current_page:
                    if self.args.open_start_url_in_manual_mode:
                        logging.info("Manual handoff mode: opening start URL for human navigation.")
                        page.goto(self.args.start_url, wait_until="domcontentloaded", timeout=self.args.navigation_timeout_ms)
                    if not self.wait_for_manual_results_page(page):
                        self.stop_reason = "setup ended before results processing"
                        return
                    self.start_runtime_clock()
                    self.process_current_results_page(page, batches)
                    return

                logging.info("Opening login/start URL. Credentials are entered only by the human.")
                page.goto(self.args.start_url, wait_until="domcontentloaded", timeout=self.args.navigation_timeout_ms)
                wait_for_ready_signal()
                self.start_runtime_clock()
                if not self.check_plain_orbis_or_exit(page, context, "post-login"):
                    self.stop_reason = "setup validation failed"
                    return

                for index, batch in enumerate(batches, start=1):
                    if self.args.max_batches and index > self.args.max_batches:
                        logging.info("Reached --max-batches=%s; stopping.", self.args.max_batches)
                        self.stop_reason = "max-batches"
                        break
                    self.process_batch(page, batch, index, len(batches))
                    if self.stop_reason:
                        break
            except KeyboardInterrupt:
                close_context = False
                self.stop_reason = "user interruption"
                # Caveat: Playwright owns the browser subprocess, so skipping context.close()
                # may still not keep the window alive if driver shutdown hooks terminate it.
                # Verify this manually with one Ctrl+C test rather than assuming it works.
                logging.info(
                    "Interrupted by user (Ctrl+C). Leaving the browser open — you can inspect it manually, or restart the script later to resume."
                )
            except StopRun as exc:
                if not self.stop_reason:
                    self.stop_reason = exc.stop_reason
                logging.critical("%s", exc)
            finally:
                logging.info("Orbis National ID auto-match shortcut count: %s.", self.orbis_auto_match_count)
                if close_context:
                    context.close()
                self.print_run_summary()

    def wait_for_manual_results_page(self, page: Page) -> bool:
        print(
            "\nThis script has opened its own Playwright Chromium window. Use THAT browser window, "
            "even if it starts blank. In that window, manually open plain Orbis, log in through TUM, "
            "go to Batch Search, upload the Excel file, apply mapping if needed, and wait until the "
            'matching results are visible. Then return here and type "ready".\n'
        )
        while True:
            answer = input("> ").strip().lower()
            if answer not in {"ready", "r"}:
                print('Waiting. Type "ready" when the loaded Batch Search results are visible, or Ctrl+C to stop.')
                continue

            current_url = page.url
            logging.info("Current URL before results-page DOM validation: %s", current_url)
            print(f"Current URL before results-page DOM validation: {current_url}")
            ok, reason = validate_results_page_dom_with_retries(page)
            if ok:
                logging.info("Manual results page validation passed: %s", reason)
                return True

            logging.warning("Manual results page validation failed for %s: %s", current_url, reason)
            print(f"This does not look like a loaded Batch Search results page. Actual URL: {current_url}")
            print(f"Reason: {reason}")
            print("Please manually upload the Excel file, wait until matching results are visible, then type ready again.")

    def process_current_results_page(self, page: Page, batches: list[Batch]) -> None:
        self.dismiss_known_modal(page)
        if not batches:
            if self.args.reselect_checkpointed:
                logging.info("No checkpointed upload rows with stored selected BvD IDs were found for re-selection.")
            else:
                logging.info("No pending upload rows found after applying checkpoints.")
            return

        batch = batches[0]
        while True:
            confirmation = self.confirm_manual_source_batch(page, batch)
            if confirmation == "confirmed":
                break
            if confirmation == "aborted":
                logging.info("Manual results-page mode aborted before processing by user confirmation.")
                return

            # A rejected stale handoff returns control to the human, so that
            # manual re-upload/wait time must not consume the runtime budget.
            self.runtime_started_at = None
            if not self.wait_for_manual_results_page(page):
                self.stop_reason = "setup ended before results processing"
                return
            self.start_runtime_clock()

        actual_page_size = self.set_page_size(page, batch)
        if self.args.reselect_checkpointed:
            self.iterate_reselect_result_pages(page, batch, actual_page_size)
        else:
            self.iterate_result_pages(page, batch, actual_page_size)

    def confirm_manual_source_batch(self, page: Page, batch: Batch) -> str:
        if not self.args.current_upload_file:
            print(
                "\nWARNING: --current-upload-file was not provided. "
                "The script will assume the first pending upload batch."
            )
            logging.warning(
                "Manual results-page mode has no --current-upload-file; assuming first pending batch %s.",
                batch.key,
            )

        first_own_id = batch.rows[0].own_id if batch.rows else ""
        last_own_id = batch.rows[-1].own_id if batch.rows else ""
        print(f"\nManual handoff source file assumed: {batch.source_file.name}")
        print(f"Rows expected: {first_own_id} - {last_own_id}")
        print(f"Company count: {len(batch.rows)}")
        print("\nType `yes` to confirm this is the file you manually uploaded.")
        answer = input("> ").strip().casefold()
        if answer != "yes":
            print("Confirmation not received. Exiting safely without processing.")
            return "aborted"

        if self.args.reselect_checkpointed:
            return "confirmed"

        return "confirmed"

    def build_batches(self) -> list[Batch]:
        input_dir = Path(self.args.input_dir)
        files = sorted(input_dir.glob("orbis_upload_*.xlsx"), key=natural_key)
        selected_files: list[str] = []
        if self.args.only_file:
            selected_files.extend(self.args.only_file)
        if self.args.current_upload_file:
            selected_files.append(self.args.current_upload_file)
        if selected_files:
            wanted = {Path(item).name for item in selected_files}
            files = [path for path in files if path.name in wanted]
            missing = wanted - {path.name for path in files}
            if missing:
                logging.warning("Requested upload file(s) not found under %s: %s", input_dir, ", ".join(sorted(missing)))

        batches: list[Batch] = []
        for source_file in files:
            header, rows = read_upload_file(source_file)
            rows = [
                UploadRow(
                    own_id=row.own_id,
                    company=row.company,
                    values=row.values,
                    source_file=row.source_file,
                    source_row=row.source_row,
                    reference_country=normalize_cell(
                        (self.reference_by_id.get(row.own_id) or {}).get("ISO2")
                        or (self.reference_by_id.get(row.own_id) or {}).get("Country")
                    ),
                    reference_city=normalize_cell(
                        (self.reference_by_id.get(row.own_id) or {}).get("City")
                        or (self.reference_by_id.get(row.own_id) or {}).get("HQ city")
                    ),
                )
                for row in rows
            ]
            if self.args.reselect_checkpointed:
                eligible_ids = {
                    row.own_id
                    for row in rows
                    if row.own_id in self.completed_ids and self.selected_bvdid_by_own_id.get(row.own_id)
                }
                if not eligible_ids:
                    logging.info("%s: no checkpointed rows with stored selected BvD IDs; skipping re-select batch.", source_file.name)
                    continue
                processing_rows = rows
            else:
                processing_rows = [row for row in rows if row.own_id not in self.completed_ids]

            if self.args.reselect_checkpointed and self.args.start_from_current_page:
                row_chunks = [processing_rows]
            else:
                row_chunks = chunks(processing_rows, self.args.batch_size)
            for chunk_index, chunk in enumerate(row_chunks, start=1):
                first = chunk[0].source_row
                last = chunk[-1].source_row
                key = f"{source_file.stem}_part{chunk_index:03d}_rows{first:04d}-{last:04d}"
                upload_file = self.temp_upload_dir / f"{key}.xlsx"
                write_upload_slice(header, chunk, upload_file)
                batches.append(Batch(key, source_file, upload_file, chunk))
        return batches

    def process_batch(self, page: Page, batch: Batch, index: int, total: int) -> None:
        self.dismiss_known_modal(page)
        logging.info("Batch %s/%s: %s (%s companies)", index, total, batch.key, len(batch.rows))
        try:
            self.open_batch_search(page, batch)
            self.upload_batch(page, batch)
            self.wait_for_results(page, batch)
            actual_page_size = self.set_page_size(page, batch)
            if self.args.reselect_checkpointed:
                self.iterate_reselect_result_pages(page, batch, actual_page_size)
            else:
                self.iterate_result_pages(page, batch, actual_page_size)
        except StopRun:
            raise
        except Exception as exc:
            self.capture_failure(page, "batch", batch.key, exc)
            logging.exception("Batch failed and will be skipped for now: %s", batch.key)

    def iterate_reselect_result_pages(self, page: Page, batch: Batch, actual_page_size: int) -> None:
        max_pages = (len(batch.rows) + actual_page_size - 1) // actual_page_size + 2
        self.result_page_search_limit = max_pages
        page_number = 1
        processed_pages = 0
        reselected_count = 0

        logging.info(
            "Re-select mode: processing only checkpointed companies with stored selected BvD IDs; no candidate data will be collected or written."
        )
        while page_number <= max_pages:
            self.ensure_no_open_trays_before_scan(page, page_number)
            visible_names = collect_visible_company_names(page)
            visible_name_set = set(visible_names)
            rows = match_visible_rows_to_upload_rows(visible_names, batch.rows) if visible_names else []
            eligible_rows = [
                row
                for row in rows
                if row.own_id in self.completed_ids and self.selected_bvdid_by_own_id.get(row.own_id)
            ]

            companies_on_page = 0
            for row in eligible_rows:
                self.ensure_no_open_trays_before_scan(page, page_number)
                if self.reached_run_limit():
                    return
                if not company_name_is_visible_on_current_page(page, row):
                    restored_row = restore_queued_company_results_page(
                        page,
                        row,
                        max_pages=max_pages,
                        timeout_ms=self.args.results_timeout_ms,
                        before_click=self.interaction_pause,
                        after_click=self.interaction_pause,
                    )
                    if restored_row is None:
                        logging.warning(
                            "own_id=%s: could not locate checkpointed company in current Orbis results; stored selection was not changed.",
                            row.own_id,
                        )
                        continue

                self.note_company_attempt_started()
                companies_on_page += 1
                bvdid = self.selected_bvdid_by_own_id[row.own_id]
                if self.reselect_checkpointed_company(page, row, bvdid):
                    reselected_count += 1
                self.last_finished_own_id = row.own_id
                if self.stop_reason:
                    return
                self.after_company_pause()

            processed_pages += 1
            logging.info(
                "Re-select page %s complete (%s companies); advancing to page %s",
                page_number,
                companies_on_page,
                page_number + 1,
            )
            advanced = advance_to_next_results_page(
                page,
                visible_name_set,
                self.args.results_timeout_ms,
                before_click=self.interaction_pause,
                after_click=self.interaction_pause,
            )
            if not advanced:
                logging.info(
                    "No further result pages; %s pages processed total. Re-selected %s checkpointed companies.",
                    processed_pages,
                    reselected_count,
                )
                return
            self.clear_blocking_response_after_verified_success("next re-select results page loaded")
            self.raise_if_systemic_problem(page, "after next re-select results page loaded")
            page_number += 1

        logging.warning("Re-select pagination guard reached after %s pages.", max_pages)

    def reselect_checkpointed_company(self, page: Page, row: UploadRow, bvdid: str) -> bool:
        candidate = {"bvdid": bvdid, "candidate_name": ""}
        try:
            close_open_candidate_tray(page)
            company_row = resolve_company_row(page, row)
            if company_row is None:
                logging.warning(
                    "own_id=%s: could not resolve checkpointed company row; stored match %s was not re-selected.",
                    row.own_id,
                    bvdid,
                )
                return False

            self.interaction_pause()
            company_row.click(timeout=self.args.default_timeout_ms)
            self.interaction_pause()
            tray = wait_for_candidate_tray(page, self.args.results_timeout_ms, company_row, row)
            self.clear_blocking_response_after_verified_success("checkpointed company candidate tray loaded")
            self.raise_if_systemic_problem(page, "after opening checkpointed company candidate list")
            expected_candidate = tray.locator(
                f"{CANDIDATE_NAME_SELECTOR}[data-snapshot-bvdid={json.dumps(bvdid)}]"
            ).first
            if expected_candidate.count() == 0:
                logging.warning(
                    "own_id=%s: stored selected candidate %s is no longer present in the Orbis candidate tray; no alternative was selected.",
                    row.own_id,
                    bvdid,
                )
                return False

            candidate["candidate_name"] = str(expected_candidate.get_attribute("title") or "")
            if not self.select_winning_candidate_in_orbis(page, row, candidate):
                return False
            logging.info(
                "own_id=%s: re-selected existing match (%s) in Orbis UI without re-collecting candidates.",
                row.own_id,
                bvdid,
            )
            return True
        except StopRun:
            raise
        except Exception as exc:
            self.capture_failure(page, "reselect_checkpointed", row.own_id, exc)
            logging.warning(
                "own_id=%s: could not re-select stored candidate %s; no alternative was selected: %s",
                row.own_id,
                bvdid,
                exc,
            )
            return False
        finally:
            self.collapse_company_tray(page, row)

    def iterate_result_pages(self, page: Page, batch: Batch, actual_page_size: int) -> None:
        max_pages = (len(batch.rows) + actual_page_size - 1) // actual_page_size + 2
        self.result_page_search_limit = max_pages
        page_number = 1
        processed_pages = 0

        while page_number <= max_pages:
            self.ensure_no_open_trays_before_scan(page, page_number)
            visible_names = collect_visible_company_names(page)
            visible_name_set = set(visible_names)
            rows = match_visible_rows_to_upload_rows(visible_names, batch.rows) if visible_names else []
            if not rows:
                logging.warning(
                    "Could not map visible result rows on page %s to upload rows; no companies processed on this page.",
                    page_number,
                )

            companies_on_page = 0
            processed_this_page: set[str] = set()
            # Re-scan the page each pass instead of trusting the up-front queue: a successful
            # selection makes Orbis re-render/reorder the results list, so the remaining queued
            # rows must be re-resolved against the current DOM rather than assumed still in place.
            inner_pass = 0
            inner_pass_cap = len(batch.rows) + 5  # generous; guarantees termination
            while True:
                inner_pass += 1
                if inner_pass > inner_pass_cap:
                    logging.warning(
                        "Row-selection loop on page %s exceeded its safety cap (%s passes); breaking to avoid a stuck loop.",
                        page_number,
                        inner_pass_cap,
                    )
                    break
                if self.reached_run_limit():
                    return
                self.ensure_no_open_trays_before_scan(page, page_number)
                current_names = collect_visible_company_names(page)
                visible_name_set = set(current_names)
                current_rows = match_visible_rows_to_upload_rows(current_names, batch.rows) if current_names else []
                row = next(
                    (
                        r
                        for r in current_rows
                        if r.own_id not in processed_this_page and r.own_id not in self.completed_ids
                    ),
                    None,
                )
                if row is None:
                    break
                processed_this_page.add(row.own_id)

                if not company_name_is_visible_on_current_page(page, row):
                    logging.warning(
                        "own_id=%s: queued company is no longer on outer results page %s; recording a bounded "
                        "row-resolution failure and skipping (no page navigation attempted).",
                        row.own_id,
                        page_number,
                    )
                    incomplete_candidates = [
                        {
                            "candidate_index": "",
                            "bvdid": "",
                            "candidate_name": "",
                            "error": (
                                f"Row was no longer visible on results page {page_number} when its "
                                "turn came to be processed; not navigated or retried this run."
                            ),
                            "reason_type": "row_not_on_page",
                        }
                    ]
                    decision = build_decision_record(
                        row,
                        batch,
                        select_best_match(
                            self.reference_by_id.get(row.own_id),
                            [],
                            len(incomplete_candidates),
                        ),
                        [],
                        incomplete_candidates,
                    )
                    self.append_result(decision)
                    self.last_finished_own_id = row.own_id
                    continue
                self.note_company_attempt_started()
                companies_on_page += 1
                successful_candidates = self.process_company(page, batch, row)
                self.last_finished_own_id = row.own_id
                if self.stop_reason:
                    return
                if successful_candidates == 0 and self.last_company_counts_as_collection_failure:
                    self.consecutive_company_full_failures += 1
                    if self.consecutive_company_full_failures >= 3:
                        message = (
                            "Possible rate-block or session loss detected: 3 consecutive companies produced no "
                            "successfully collected candidates. Stopping the run to avoid escalation."
                        )
                        logging.critical("%s", message)
                        raise StopRun(message)
                elif successful_candidates > 0:
                    self.consecutive_company_full_failures = 0
                else:
                    self.consecutive_company_full_failures = 0
                    logging.debug(
                        "own_id=%s: zero candidates were returned for a non-collection reason; excluded from the consecutive collection-failure guard.",
                        row.own_id,
                    )

                self.after_company_pause()

            processed_pages += 1
            logging.info(
                "Page %s complete (%s companies); advancing to page %s",
                page_number,
                companies_on_page,
                page_number + 1,
            )

            advanced = advance_to_next_results_page(
                page,
                visible_name_set,
                self.args.results_timeout_ms,
                before_click=self.interaction_pause,
                after_click=self.interaction_pause,
            )
            if not advanced:
                logging.info("No further result pages; %s pages processed total.", processed_pages)
                self.reconcile_unresolved_batch_rows(batch)
                return
            self.clear_blocking_response_after_verified_success("next results page loaded")
            self.raise_if_systemic_problem(page, "after next results page loaded")
            page_number += 1

        self.reconcile_unresolved_batch_rows(batch)
        logging.warning("Pagination guard hit after %s pages; stopping to avoid an infinite loop.", max_pages)

    def reconcile_unresolved_batch_rows(self, batch: Batch) -> None:
        decided_own_ids = {
            key[1] for key in self.seen_result_keys if key[0] == "decision"
        }
        for row in batch.rows:
            if row.own_id in self.completed_ids or row.own_id in decided_own_ids:
                continue
            logging.warning(
                "own_id=%s: no decision record exists after this batch was fully "
                "traversed (likely a page-mapping failure); recording a bounded "
                "row-resolution failure.",
                row.own_id,
            )
            incomplete_candidates = [
                {
                    "candidate_index": "",
                    "bvdid": "",
                    "candidate_name": "",
                    "error": (
                        "Row never matched to a visible results-page entry in any page "
                        "of this batch; not scraped this run."
                    ),
                    "reason_type": "row_never_resolved",
                }
            ]
            decision = build_decision_record(
                row,
                batch,
                select_best_match(
                    self.reference_by_id.get(row.own_id),
                    [],
                    len(incomplete_candidates),
                ),
                [],
                incomplete_candidates,
            )
            self.append_result(decision)

    def start_runtime_clock(self) -> None:
        if self.runtime_started_at is None:
            self.runtime_started_at = time.monotonic()
            logging.info("Per-company runtime clock started after manual handoff.")

    def elapsed_runtime_minutes(self) -> float:
        if self.runtime_started_at is None:
            return 0.0
        return max(0.0, (time.monotonic() - self.runtime_started_at) / 60.0)

    def note_company_attempt_started(self) -> None:
        self.companies_attempted_this_run += 1
        if (
            self.args.max_companies
            and self.companies_attempted_this_run >= self.args.max_companies
            and self.max_companies_reached_at is None
        ):
            self.max_companies_reached_at = time.monotonic()

    def reached_run_limit(self) -> bool:
        if self.stop_reason:
            return True

        now = time.monotonic()
        max_companies_reached = bool(
            self.args.max_companies
            and self.companies_attempted_this_run >= self.args.max_companies
        )
        runtime_deadline = None
        runtime_reached = False
        if self.args.max_runtime_minutes is not None and self.runtime_started_at is not None:
            runtime_deadline = self.runtime_started_at + self.args.max_runtime_minutes * 60
            runtime_reached = now >= runtime_deadline

        runtime_wins = runtime_reached and (
            not max_companies_reached
            or self.max_companies_reached_at is None
            or (runtime_deadline is not None and runtime_deadline <= self.max_companies_reached_at)
        )
        if runtime_wins:
            self.stop_reason = "runtime cap"
            finished = self.last_finished_own_id or "<none>"
            logging.info(
                "Runtime cap of %s min reached after finishing %s - stopping run.",
                self.args.max_runtime_minutes,
                finished,
            )
            return True
        if max_companies_reached:
            self.stop_reason = "max-companies"
            logging.info("Reached --max-companies=%s; stopping.", self.args.max_companies)
            return True
        return False

    def runtime_hard_limit_reached(self) -> bool:
        if self.args.max_runtime_minutes is None or self.runtime_started_at is None:
            return False
        hard_limit_minutes = self.args.max_runtime_minutes + RUNTIME_HARD_BUFFER_MINUTES
        return self.elapsed_runtime_minutes() >= hard_limit_minutes

    def print_run_summary(self) -> None:
        reason = self.stop_reason or "batch exhausted"
        summary = (
            f"Run summary: stop reason={reason}; elapsed={self.elapsed_runtime_minutes():.1f} min; "
            f"companies processed this run={self.companies_attempted_this_run}."
        )
        logging.info(summary)
        print(summary)

    def note_blocking_response(self, response: Any) -> None:
        try:
            if response.status in {403, 429}:
                self.last_block_response_reason = f"HTTP {response.status} from {response.url}"
        except Exception:
            return

    def clear_blocking_response_after_verified_success(self, verified_state: str) -> None:
        if not self.last_block_response_reason:
            return
        logging.info(
            "Clearing transient blocking-response marker after verified UI success (%s): %s",
            verified_state,
            self.last_block_response_reason,
        )
        self.last_block_response_reason = ""

    def raise_if_systemic_problem(self, page: Page, context: str) -> None:
        reason = detect_systemic_problem(page, self.last_block_response_reason)
        if not reason:
            return
        message = f"Possible rate-block or session loss detected: {reason}. Stopping the run to avoid escalation."
        logging.critical("%s Context: %s", message, context)
        self.capture_systemic_problem(page, context, reason)
        raise StopRun(message)

    def capture_systemic_problem(self, page: Page, context: str, reason: str) -> None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_context = sanitize_filename(context)
        screenshot = self.screenshot_dir / f"{stamp}_{safe_context}_systemic_problem.png"
        diagnostic = self.log_dir / f"{stamp}_{safe_context}_systemic_problem.txt"
        screenshot_note = ""
        try:
            page.screenshot(path=str(screenshot), full_page=True)
            screenshot_note = f" Screenshot: {screenshot}."
        except Exception as exc:
            screenshot_note = f" Screenshot failed: {exc}."

        try:
            body_text = page.locator("body").inner_text(timeout=2000)
        except Exception as exc:
            body_text = f"<body text unavailable: {exc}>"
        excerpt = diagnostic_excerpt(body_text, reason)
        try:
            diagnostic.write_text(
                f"URL: {page.url}\nContext: {context}\nReason: {reason}\n\nRelevant page text:\n{excerpt}\n",
                encoding="utf-8",
            )
            logging.critical(
                "Systemic-problem diagnostics saved.%s Diagnostic: %s.",
                screenshot_note,
                diagnostic,
            )
        except Exception as exc:
            logging.critical("Could not save systemic-problem text diagnostic: %s.%s", exc, screenshot_note)

    def interaction_pause(self) -> None:
        human_pause(self.args.interaction_min_wait, self.args.interaction_max_wait)

    def after_company_pause(self) -> None:
        # Keep the normal randomized pause, but do not interrupt the run with
        # the former longer rest prompt after every 8-12 companies.
        human_pause(12.0, 30.0)

    def open_batch_search(self, page: Page, batch: Batch) -> None:
        try:
            batch_search_url = DEFAULT_BATCH_SEARCH_URL
            logging.info("Batch %s: navigating directly to Batch Search page: %s", batch.key, batch_search_url)
            page.goto(batch_search_url, wait_until="domcontentloaded", timeout=self.args.navigation_timeout_ms)
            page.wait_for_load_state("domcontentloaded", timeout=self.args.navigation_timeout_ms)
            self.raise_if_systemic_problem(page, "after opening batch search")
            ok, reason = explain_plain_orbis_url(page.url)
            logging.info("Current URL before plain-Orbis validation (batch-search): %s", page.url)
            print(f"Current URL before plain-Orbis validation: {page.url}")
            if not ok:
                raise RuntimeError(f"Landed on the wrong Orbis product: {page.url}. Rejected because: {reason}")
            human_pause(self.args.min_wait, self.args.max_wait)
        except Exception as exc:
            self.capture_failure(page, "open_batch_search", batch.key, exc)
            raise

    def check_plain_orbis_or_exit(self, page: Page, context: Any, label: str) -> bool:
        current_url = page.url
        logging.info("Current URL before plain-Orbis validation (%s): %s", label, current_url)
        print(f"Current URL before plain-Orbis validation: {current_url}")
        ok, reason = explain_plain_orbis_url(current_url)
        if ok:
            return True
        print(
            "Landed on the wrong Orbis product — please navigate to plain Orbis (not Orbis M&A) and re-run. "
            f"Actual URL: {current_url}. Rejected because: {reason}"
        )
        logging.error("Rejected Orbis URL (%s): %s; reason=%s", label, current_url, reason)
        return False

    def dismiss_known_modal(self, page: Page) -> None:
        try:
            found = page.evaluate(
                """() => {
                    function visible(el) {
                        const style = window.getComputedStyle(el);
                        const rect = el.getBoundingClientRect();
                        return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                    }
                    function text(el) {
                        return String(el.innerText || el.textContent || "").replace(/\\s+/g, " ").trim();
                    }
                    function okControl(container) {
                        const controls = [...container.querySelectorAll("button, a, input[type='button'], input[type='submit']")];
                        return controls.find((control) => {
                            const label = String(control.value || control.innerText || control.textContent || "").trim();
                            return visible(control) && /^OK$/i.test(label);
                        });
                    }

                    const modalTextPattern = /we have updated|we've updated|release notice|what's new|whats new|new features/i;
                    const containers = [...document.querySelectorAll("[role='dialog'], .modal, .ui-dialog, .ui-widget.ui-dialog, .owPopup, .popup")]
                        .filter(visible);
                    for (const container of containers) {
                        const body = text(container);
                        if (!modalTextPattern.test(body)) continue;
                        const ok = okControl(container);
                        if (!ok) continue;
                        ok.setAttribute("data-codex-known-modal-ok", "true");
                        return {matched: true, text: body.slice(0, 160)};
                    }
                    return {matched: false, text: ""};
                }"""
            )
            if not found.get("matched"):
                return

            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot = self.screenshot_dir / f"modal_dismissed_{stamp}.png"
            page.screenshot(path=str(screenshot), full_page=True)
            logging.info("Dismissed known release/update modal: %r. Screenshot: %s", found.get("text", ""), screenshot)
            page.locator("[data-codex-known-modal-ok='true']").first.click(timeout=self.args.default_timeout_ms)
            human_pause(self.args.min_wait, self.args.max_wait)
        except Exception as exc:
            logging.warning("Known-modal dismiss check failed without clicking a fallback: %s", exc)

    def upload_batch(self, page: Page, batch: Batch) -> None:
        try:
            file_input = page.locator("input[type='file']").first
            file_input.set_input_files(str(batch.upload_file))
            self.interaction_pause()
            click_first(page, ["a.button.ok[data-form-button='submit']", "input.button.ok[value='Upload']"], "Upload")
            page.wait_for_load_state("domcontentloaded", timeout=self.args.navigation_timeout_ms)
            self.interaction_pause()
            self.raise_if_systemic_problem(page, "after upload click")
            click_first(page, ["input.button.ok[value='Apply']"], "Apply mapping")
            self.interaction_pause()
            self.raise_if_systemic_problem(page, "after apply mapping click")
        except Exception as exc:
            self.capture_failure(page, "upload", batch.key, exc)
            raise

    def wait_for_results(self, page: Page, batch: Batch) -> None:
        try:
            wait_for_any_selector(page, RESULT_TABLE_SELECTORS, self.args.results_timeout_ms)
            self.clear_blocking_response_after_verified_success("batch results loaded")
            self.raise_if_systemic_problem(page, "after waiting for results")
        except Exception as exc:
            self.capture_failure(page, "wait_results", batch.key, exc)
            raise

    def set_page_size(self, page: Page, batch: Batch) -> int:
        fallback_page_size = 10
        try:
            locator = page.locator("select#pageSize")
            locator.wait_for(state="visible", timeout=self.args.default_timeout_ms)
            self.interaction_pause()
            locator.select_option("100")
            self.interaction_pause()
            page.wait_for_load_state("networkidle", timeout=self.args.navigation_timeout_ms)
            self.clear_blocking_response_after_verified_success("requested results page size loaded")
            self.raise_if_systemic_problem(page, "after setting page size")
            actual = read_page_size(page)
            if actual:
                if actual == 100:
                    logging.info("Resolved page size: 100 (requested)")
                else:
                    logging.warning("Resolved page size: %s (requested 100 but page reports different value)", actual)
                return actual
            logging.warning("Resolved page size: 10 (fallback, could not confirm)")
            return fallback_page_size
        except PlaywrightTimeoutError:
            logging.warning("Batch %s: page-size selector was not available; continuing.", batch.key)
            logging.warning("Resolved page size: 10 (fallback, could not confirm)")
            return fallback_page_size
        except Exception as exc:
            self.capture_failure(page, "page_size", batch.key, exc)
            logging.warning("Batch %s: could not set page size; continuing.", batch.key)
            logging.warning("Resolved page size: 10 (fallback, could not confirm)")
            return fallback_page_size

    def process_company(self, page: Page, batch: Batch, row: UploadRow) -> int:
        logging.info("own_id=%s: opening %s", row.own_id, row.company)
        if self.args.candidate_soft_cap is not None:
            company_soft_cap = self.args.candidate_soft_cap
            soft_cap_description = "fixed override"
        else:
            company_soft_cap = random.randint(
                self.args.candidate_soft_cap_min,
                self.args.candidate_soft_cap_max,
            )
            soft_cap_description = (
                f"random {self.args.candidate_soft_cap_min}-{self.args.candidate_soft_cap_max}"
            )
        logging.info(
            "own_id=%s: candidate soft cap for this company = %s (%s); hard cap = %s",
            row.own_id,
            company_soft_cap,
            soft_cap_description,
            self.args.candidate_hard_cap,
        )
        self.last_company_counts_as_collection_failure = False
        candidate_tray: Locator | None = None
        orbis_auto_match: dict[str, str] = {}
        try:
            close_open_candidate_tray(page)
            company_row = resolve_company_row(page, row)
            if company_row is None:
                raise RuntimeError(f"Could not resolve a unique result row for own_id={row.own_id} company={row.company}")
            orbis_auto_match = read_orbis_auto_match_from_company_row(company_row)

            self.interaction_pause()
            company_row.click(timeout=self.args.default_timeout_ms)
            self.interaction_pause()
            candidate_tray = wait_for_candidate_tray(page, self.args.results_timeout_ms, company_row, row)
            self.clear_blocking_response_after_verified_success("company candidate tray loaded")
            self.raise_if_systemic_problem(page, "after waiting for company candidate tray")
        except Exception as exc:
            if isinstance(exc, StopRun):
                incomplete_candidates = [
                    {
                        "candidate_index": "",
                        "bvdid": "",
                        "candidate_name": "",
                        "error": str(exc),
                        "reason_type": "collection_failure",
                    }
                ]
                decision = build_decision_record(
                    row,
                    batch,
                    select_best_match(self.reference_by_id.get(row.own_id), [], len(incomplete_candidates)),
                    [],
                    incomplete_candidates,
                )
                self.append_result(decision)
                self.collapse_company_tray(page, row)
                raise
            self.capture_failure(page, "open_company", row.own_id, exc)
            logging.exception("own_id=%s: could not open candidate list; continuing.", row.own_id)
            if orbis_auto_match.get("selected_name"):
                logging.warning(
                    "own_id=%s: row shows Orbis National ID auto-match (%s), but the candidate tray could not be opened; no auto-match decision will be written without a verified BvD ID.",
                    row.own_id,
                    orbis_auto_match.get("selected_name"),
                )
                self.collapse_company_tray(page, row)
                return 0
            incomplete_candidates = [
                {
                    "candidate_index": "",
                    "bvdid": "",
                    "candidate_name": "",
                    "error": str(exc),
                    "reason_type": "collection_failure",
                }
            ]
            decision = build_decision_record(
                row,
                batch,
                select_best_match(self.reference_by_id.get(row.own_id), [], len(incomplete_candidates)),
                [],
                incomplete_candidates,
            )
            self.append_result(decision)
            self.collapse_company_tray(page, row)
            return 0

        candidate_cells = candidate_tray.locator(CANDIDATE_NAME_SELECTOR)
        candidate_count = candidate_cells.count()
        if candidate_count == 0:
            if tray_has_explicit_no_result(candidate_tray):
                logging.info('own_id=%s: Orbis returned "No result" — no candidates to match.', row.own_id)
            else:
                logging.info("own_id=%s: no candidates found.", row.own_id)
            decision = build_decision_record(row, batch, select_best_match(self.reference_by_id.get(row.own_id), []), [])
            self.append_result(decision)
            self.mark_completed(row.own_id)
            self.collapse_company_tray(page, row)
            return 0

        if orbis_auto_match.get("selected_name"):
            try:
                if self.process_orbis_auto_match(page, batch, row, candidate_tray, orbis_auto_match):
                    return 1
            except StopRun:
                raise
            except Exception as exc:
                try:
                    close_snapshot_popup(page)
                except Exception:
                    pass
                logging.warning(
                    "own_id=%s: Orbis National ID auto-match shortcut failed; falling back to full candidate collection: %s",
                    row.own_id,
                    exc,
                )

        countries_seen: set[str] = set()
        candidate_stubs: list[dict[str, Any]] = []
        for candidate_index in range(candidate_count):
            try:
                candidate = extract_candidate(candidate_cells.nth(candidate_index))
                if candidate.get("country"):
                    countries_seen.add(str(candidate["country"]))
                candidate.update(
                    {
                        "own_id": row.own_id,
                        "company": row.company,
                        "source_file": row.source_file,
                        "source_row": row.source_row,
                        "batch_key": batch.key,
                        "source": "candidate_full_audit",
                        "candidate_index": candidate_index + 1,
                    }
                )
                candidate_stubs.append(candidate)
            except Exception as exc:
                self.capture_failure(page, f"candidate_stub_{candidate_index + 1}", row.own_id, exc)
                logging.exception("own_id=%s: candidate stub %s failed; continuing.", row.own_id, candidate_index + 1)
                candidate_stubs.append(
                    {
                        "own_id": row.own_id,
                        "company": row.company,
                        "source_file": row.source_file,
                        "source_row": row.source_row,
                        "batch_key": batch.key,
                        "candidate_index": candidate_index + 1,
                        "_stub_error": str(exc),
                    }
                )

        candidates: list[dict[str, Any]] = []
        incomplete_candidates: list[dict[str, Any]] = []
        failed_full_candidate_count = 0
        candidate_cap_reached = False
        cap_reason = ""
        runtime_hard_stop_reached = False
        ref_row = self.reference_by_id.get(row.own_id)
        for candidate_index, candidate in enumerate(candidate_stubs, start=1):
            if self.runtime_hard_limit_reached():
                runtime_hard_stop_reached = True
                hard_limit_minutes = self.args.max_runtime_minutes + RUNTIME_HARD_BUFFER_MINUTES
                cap_reason = (
                    f"Not visited - runtime hard stop reached at {hard_limit_minutes} minutes "
                    f"({RUNTIME_HARD_BUFFER_MINUTES}-minute buffer after soft cap)"
                )
                incomplete_candidates.extend(
                    incomplete_records_for_unvisited(
                        candidate_stubs[candidate_index - 1:],
                        cap_reason,
                        reason_type="runtime_hard_stop",
                        status="not_visited_runtime_hard_stop",
                    )
                )
                self.stop_reason = "runtime hard cap"
                logging.info(
                    "Runtime hard cap of %s min reached during own_id=%s after collecting %s candidate(s); finalizing the partial company and stopping run.",
                    hard_limit_minutes,
                    row.own_id,
                    len(candidates),
                )
                break

            if candidate.get("_stub_error"):
                incomplete_candidates.append(
                    {
                        "candidate_index": candidate.get("candidate_index", candidate_index),
                        "bvdid": candidate.get("bvdid", ""),
                        "candidate_name": candidate.get("candidate_name", ""),
                        "error": candidate.get("_stub_error", ""),
                        "reason_type": "collection_failure",
                    }
                )
                continue

            if should_stop_for_candidate_cap(
                candidates,
                ref_row,
                company_soft_cap,
                self.args.candidate_hard_cap,
            ):
                candidate_cap_reached = True
                cap_reason = cap_incomplete_reason(
                    candidates,
                    ref_row,
                    company_soft_cap,
                    self.args.candidate_hard_cap,
                )
                incomplete_candidates.extend(incomplete_records_for_unvisited(candidate_stubs[candidate_index - 1:], cap_reason))
                logging.info("own_id=%s: candidate collection cap reached: %s", row.own_id, cap_reason)
                break

            success = False
            last_error = ""
            for attempt in range(1, 3):
                try:
                    self.collect_candidate_snapshot_and_report(page, row, candidate)
                    self.append_result(candidate)
                    candidates.append(candidate)
                    human_pause(self.args.min_wait, self.args.max_wait)
                    success = True
                    break
                except StopRun as exc:
                    incomplete_candidates.append(
                        {
                            "candidate_index": candidate.get("candidate_index", candidate_index),
                            "bvdid": candidate.get("bvdid", ""),
                            "candidate_name": candidate.get("candidate_name", ""),
                            "error": str(exc),
                            "reason_type": "collection_failure",
                        }
                    )
                    match_decision = select_best_match(ref_row, candidates, len(incomplete_candidates))
                    decision = build_decision_record(row, batch, match_decision, candidates, incomplete_candidates)
                    self.append_result(decision)
                    self.collapse_company_tray(page, row)
                    raise
                except Exception as exc:
                    last_error = str(exc)
                    self.capture_failure(page, f"candidate_full_{candidate_index}_attempt_{attempt}", row.own_id, exc)
                    logging.exception(
                        "own_id=%s bvdid=%s: full candidate collection failed on attempt %s.",
                        row.own_id,
                        candidate.get("bvdid", ""),
                        attempt,
                    )
                    try:
                        close_snapshot_popup(page)
                        page.go_back(wait_until="domcontentloaded", timeout=self.args.navigation_timeout_ms)
                        wait_for_any_selector(page, RESULT_TABLE_SELECTORS, self.args.results_timeout_ms)
                        self.clear_blocking_response_after_verified_success(
                            "results list restored during candidate recovery"
                        )
                        self.interaction_pause()
                        self.raise_if_systemic_problem(page, "after recovery back navigation from candidate failure")
                    except StopRun as stop_exc:
                        incomplete_candidates.append(
                            {
                                "candidate_index": candidate.get("candidate_index", candidate_index),
                                "bvdid": candidate.get("bvdid", ""),
                                "candidate_name": candidate.get("candidate_name", ""),
                                "error": str(stop_exc),
                                "reason_type": "collection_failure",
                            }
                        )
                        match_decision = select_best_match(ref_row, candidates, len(incomplete_candidates))
                        decision = build_decision_record(row, batch, match_decision, candidates, incomplete_candidates)
                        self.append_result(decision)
                        self.collapse_company_tray(page, row)
                        raise
                    except Exception:
                        pass
            if not success:
                incomplete_candidates.append(
                    {
                        "candidate_index": candidate.get("candidate_index", candidate_index),
                        "bvdid": candidate.get("bvdid", ""),
                        "candidate_name": candidate.get("candidate_name", ""),
                        "error": last_error,
                        "reason_type": "collection_failure",
                    }
                )
                failed_full_candidate_count += 1
                if failed_full_candidate_count >= 3:
                    remaining = candidate_stubs[candidate_index:]
                    for remaining_candidate in remaining:
                        incomplete_candidates.append(
                            {
                                "candidate_index": remaining_candidate.get("candidate_index", ""),
                                "bvdid": remaining_candidate.get("bvdid", ""),
                                "candidate_name": remaining_candidate.get("candidate_name", ""),
                                "error": "Skipped after 3 candidates failed collection for this company.",
                                "reason_type": "collection_failure",
                            }
                        )
                    logging.warning(
                        "own_id=%s: stopping after 3 candidates failed collection; writing partial decision from %s successfully collected candidate(s)",
                        row.own_id,
                        len(candidates),
                    )
                    break
            if failed_full_candidate_count >= 3:
                break

        warn_if_country_spread(row, countries_seen)
        match_decision = select_best_match(ref_row, candidates, len(incomplete_candidates))
        decision = build_decision_record(row, batch, match_decision, candidates, incomplete_candidates)
        self.append_result(decision)
        if match_decision.decided and match_decision.selected:
            self.select_winning_candidate_in_orbis(page, row, match_decision.selected)
        self.collapse_company_tray(page, row)
        collection_failures = [item for item in incomplete_candidates if item.get("reason_type") == "collection_failure"]
        if collection_failures and not runtime_hard_stop_reached:
            self.last_company_counts_as_collection_failure = len(candidates) == 0
            logging.warning(
                "own_id=%s: %s candidate(s) failed collection; decision written but company will not be checkpointed.",
                row.own_id,
                len(collection_failures),
            )
            return len(candidates)
        if runtime_hard_stop_reached:
            logging.info(
                "own_id=%s: partial runtime-hard-stop decision written; checkpointing before graceful shutdown.",
                row.own_id,
            )
        logging.info(
            "own_id=%s: collected %s candidate(s), wrote decision; marking company complete.",
            row.own_id,
            len(candidates),
        )
        self.mark_completed(row.own_id)
        return len(candidates)

    def process_orbis_auto_match(
        self,
        page: Page,
        batch: Batch,
        row: UploadRow,
        candidate_tray: Locator,
        auto_match: dict[str, str],
    ) -> bool:
        selected_name = auto_match.get("selected_name", "")
        candidate_cell, candidate = find_candidate_by_selected_name(candidate_tray, selected_name)
        if candidate_cell is None or candidate is None:
            raise RuntimeError(f"Selected candidate {selected_name!r} was not found in the candidate tray.")

        snapshot = page.locator("div.owSnapshot").first
        self.interaction_pause()
        click_with_force_and_dom_fallback(
            candidate_cell,
            snapshot,
            row,
            candidate,
            action_label="Orbis auto-matched candidate snapshot",
            success_label="snapshot popup",
            force_timeout_ms=5000,
            force_verify_timeout_ms=8000,
            fallback_verify_timeout_ms=self.args.snapshot_timeout_ms,
            target_factory=lambda: find_candidate_by_selected_name(
                ensure_candidate_tray(
                    page,
                    row,
                    self.args.results_timeout_ms,
                    candidate,
                    max_pages=self.result_page_search_limit,
                    before_page_click=self.interaction_pause,
                    after_page_click=lambda: self.interaction_pause(),
                ),
                selected_name,
            )[0],
            identity_attr="data-snapshot-bvdid",
            expected_identity=str(candidate.get("bvdid") or ""),
        )
        self.clear_blocking_response_after_verified_success("Orbis auto-matched candidate snapshot opened")
        self.interaction_pause()
        self.raise_if_systemic_problem(page, "after opening Orbis auto-matched candidate snapshot")

        snapshot.wait_for(state="visible", timeout=self.args.snapshot_timeout_ms)
        snapshot_text = snapshot.inner_text(timeout=self.args.default_timeout_ms)
        bvdid = extract_bvdid_from_snapshot(snapshot_text)
        if not bvdid:
            close_snapshot_popup(page)
            raise RuntimeError("Could not extract BvD ID from the selected candidate snapshot.")

        candidate.update(
            {
                "bvdid": bvdid,
                "candidate_name": selected_name,
                "snapshot_text": snapshot_text,
                "national_id": auto_match.get("identifier", "") or candidate.get("national_id", ""),
                "identifier": auto_match.get("identifier", "") or candidate.get("identifier", ""),
                "score_letter": auto_match.get("score_letter", "") or candidate.get("score_letter", ""),
                "score_numeric": auto_match.get("score_numeric", "") or candidate.get("score_numeric", ""),
                "collection_status": "Orbis auto-match (remaining candidates not collected)",
            }
        )
        decision = build_orbis_auto_match_decision_record(row, batch, candidate, auto_match)
        self.append_result(decision)
        close_snapshot_popup(page)
        self.collapse_company_tray(page, row)
        self.mark_completed(row.own_id)
        self.orbis_auto_match_count += 1
        logging.info(
            "own_id=%s: auto-matched by Orbis via National ID (%s); skipping full candidate collection.",
            row.own_id,
            selected_name,
        )
        return True

    def collapse_company_tray(self, page: Page, row: UploadRow) -> None:
        try:
            close_snapshot_popup(page)
        except Exception:
            pass

        chevron: Locator | None = None
        try:
            chevron = locate_company_dropdown_chevron(page, row)
            if chevron is None:
                if collapse_expanded_dropdown_chevrons(page, timeout_ms=5000):
                    logging.debug(
                        "own_id=%s: company row was not on the current results page; any visible expanded dropdown was closed globally.",
                        row.own_id,
                    )
                else:
                    logging.warning("own_id=%s: candidate tray may still be expanded after cleanup.", row.own_id)
                return
            if chevron is not None:
                if get_locator_attribute(chevron, "aria-expanded").casefold() == "false":
                    logging.info("own_id=%s: candidate tray already collapsed.", row.own_id)
                    return
                self.interaction_pause()
                collapsed = click_company_dropdown_chevron_with_fallback(
                    page,
                    row,
                    chevron,
                    timeout_ms=8000,
                )
                self.interaction_pause()
                if collapsed:
                    logging.info("own_id=%s: collapsed candidate tray before moving on.", row.own_id)
                    return
        except Exception as exc:
            logging.debug("own_id=%s: dropdown-chevron collapse did not complete: %s", row.own_id, exc)

        try:
            close_open_candidate_tray(page)
            if wait_for_company_chevron_collapsed(page, row, timeout_ms=8000):
                logging.info("own_id=%s: collapsed candidate tray via fallback cleanup.", row.own_id)
                return
        except Exception as exc:
            logging.debug("own_id=%s: fallback cleanup did not collapse tray: %s", row.own_id, exc)

        try:
            collapse_expanded_dropdown_chevrons(page, timeout_ms=12_000)
            if wait_for_company_chevron_collapsed(page, row, timeout_ms=3000):
                logging.info("own_id=%s: collapsed candidate tray via expanded dropdown controls.", row.own_id)
                return
            collapse_visible_popping_trays(page, timeout_ms=12_000)
            if wait_for_company_chevron_collapsed(page, row, timeout_ms=3000):
                logging.info("own_id=%s: collapsed candidate tray via expanded-row controls.", row.own_id)
                return
            page.keyboard.press("Escape")
            page.mouse.wheel(0, -2000)
            page.wait_for_timeout(1000)
            if wait_for_company_chevron_collapsed(page, row, timeout_ms=5000):
                logging.info("own_id=%s: collapsed candidate tray after final Escape/scroll cleanup.", row.own_id)
                return
            logging.warning("own_id=%s: candidate tray may still be expanded after cleanup.", row.own_id)
        except Exception as exc:
            logging.warning("own_id=%s: could not collapse candidate tray before moving on: %s", row.own_id, exc)

    def ensure_no_open_trays_before_scan(self, page: Page, page_number: int) -> None:
        open_count = expanded_company_chevron_count(page)
        if open_count == 0:
            return
        logging.info(
            "Page %s: found %s open candidate tray/trays before row scan; forcing collapse before trusting visible row order.",
            page_number,
            open_count,
        )
        if not collapse_expanded_dropdown_chevrons(page, timeout_ms=12_000):
            try:
                close_open_candidate_tray(page)
                collapse_visible_popping_trays(page, timeout_ms=5000)
            except Exception:
                pass
        remaining = expanded_company_chevron_count(page)
        if remaining:
            logging.warning(
                "Page %s: %s candidate tray/trays still visible before row scan; row order may be unstable.",
                page_number,
                remaining,
            )

    def select_winning_candidate_in_orbis(self, page: Page, row: UploadRow, candidate: dict[str, Any]) -> bool:
        bvdid = str(candidate.get("bvdid") or "").strip()
        if not bvdid:
            logging.warning("could not select winning candidate <missing bvdid> in Orbis UI; JSON remains source of truth")
            return False
        try:
            tray = ensure_candidate_tray(
                page,
                row,
                self.args.results_timeout_ms,
                candidate,
                max_pages=self.result_page_search_limit,
                before_page_click=self.interaction_pause,
                after_page_click=self.interaction_pause,
            )
            self.clear_blocking_response_after_verified_success("candidate tray re-opened for selection")
            escaped_bvdid = json.dumps(bvdid)
            radio = tray.locator(
                f"input[type='radio'][value={escaped_bvdid}], "
                f"input[type='radio'][data-inputsnapshot-bvdid={escaped_bvdid}], "
                f"input[data-inputsnapshot-bvdid={escaped_bvdid}]"
            ).first
            if radio.count() == 0:
                raise RuntimeError(f"radio input for BvD ID {bvdid} was not found")
            self.interaction_pause()
            radio.click(timeout=self.args.default_timeout_ms, force=True)
            self.interaction_pause()
            self.raise_if_systemic_problem(page, "after selecting winning candidate radio")
            logging.info("selected winning candidate %s in Orbis UI", bvdid)
            return True
        except StopRun:
            raise
        except Exception as exc:
            logging.warning(
                "could not select winning candidate %s in Orbis UI; JSON remains source of truth: %s",
                bvdid,
                exc,
            )
            return False

    def collect_candidate_snapshot_and_report(
        self,
        page: Page,
        row: UploadRow,
        candidate: dict[str, Any],
    ) -> None:
        tray = ensure_candidate_tray(
            page,
            row,
            self.args.results_timeout_ms,
            candidate,
            max_pages=self.result_page_search_limit,
            before_page_click=self.interaction_pause,
            after_page_click=self.interaction_pause,
        )
        self.clear_blocking_response_after_verified_success("candidate tray ready for collection")
        candidate_cell = candidate_locator_in_tray(tray, candidate)
        snapshot = page.locator("div.owSnapshot").first
        self.interaction_pause()
        click_with_force_and_dom_fallback(
            candidate_cell,
            snapshot,
            row,
            candidate,
            action_label="candidate snapshot",
            success_label="snapshot popup",
            force_timeout_ms=5000,
            force_verify_timeout_ms=8000,
            fallback_verify_timeout_ms=self.args.snapshot_timeout_ms,
            target_factory=lambda: candidate_locator_in_tray(
                ensure_candidate_tray(
                    page,
                    row,
                    self.args.results_timeout_ms,
                    candidate,
                    max_pages=self.result_page_search_limit,
                    before_page_click=self.interaction_pause,
                    after_page_click=lambda: self.interaction_pause(),
                ),
                candidate,
            ),
            identity_attr="data-snapshot-bvdid",
            expected_identity=str(candidate.get("bvdid") or ""),
        )
        self.clear_blocking_response_after_verified_success("candidate snapshot opened")
        self.interaction_pause()
        self.raise_if_systemic_problem(page, "after opening candidate snapshot")

        snapshot.wait_for(state="visible", timeout=self.args.snapshot_timeout_ms)
        candidate["snapshot_text"] = snapshot.inner_text(timeout=self.args.default_timeout_ms)

        go_to_report = snapshot.locator("a[data-report-link], a:has-text('Go to report')").first
        if go_to_report.count() == 0:
            logging.warning(
                "own_id=%s bvdid=%s: snapshot had no Go to report link.",
                row.own_id,
                candidate.get("bvdid", ""),
            )
            close_snapshot_popup(page)
            return

        go_to_report.wait_for(state="visible", timeout=self.args.default_timeout_ms)
        self.interaction_pause()
        click_go_to_report_with_fallback(page, go_to_report, row, candidate, self.args.results_timeout_ms)
        self.clear_blocking_response_after_verified_success("candidate report opened")
        self.interaction_pause()
        self.raise_if_systemic_problem(page, "after clicking Go to report")

        candidate.update(extract_report_fields(page))
        if candidate.get("domain"):
            candidate["website"] = candidate["domain"]

        page.go_back(wait_until="domcontentloaded", timeout=self.args.navigation_timeout_ms)
        wait_for_any_selector(page, RESULT_TABLE_SELECTORS, self.args.results_timeout_ms)
        self.clear_blocking_response_after_verified_success("results list restored after candidate report")
        self.interaction_pause()
        self.raise_if_systemic_problem(page, "after returning to results list")

    def append_result(self, result: dict[str, Any]) -> None:
        result["scraped_at"] = datetime.now().isoformat(timespec="seconds")
        key = result_key(result)
        if key in self.seen_result_keys:
            logging.info(
                "own_id=%s bvdid=%s: result already present; skipping duplicate append.",
                result.get("own_id"),
                result.get("bvdid"),
            )
            return

        with self.result_path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(result, ensure_ascii=False) + "\n")
            handle.flush()
        self.seen_result_keys.add(key)

    def mark_completed(self, own_id: str) -> None:
        if own_id in self.completed_ids:
            return
        with self.checkpoint_path.open("a", encoding="utf-8") as handle:
            handle.write(f"{own_id}\n")
            handle.flush()
        self.completed_ids.add(own_id)

    def capture_failure(
        self,
        page: Page,
        label: str,
        own_id_or_batch: str,
        exc: Exception,
        log_level: int = logging.ERROR,
    ) -> None:
        safe_label = sanitize_filename(str(label))
        safe_id = sanitize_filename(str(own_id_or_batch))
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot = self.screenshot_dir / f"{stamp}_{safe_id}_{safe_label}.png"
        try:
            page.screenshot(path=str(screenshot), full_page=True)
            logging.log(log_level, "%s failed for %s: %s. Screenshot: %s", label, own_id_or_batch, exc, screenshot)
        except Exception as screenshot_exc:
            logging.log(
                log_level,
                "%s failed for %s: %s. Screenshot also failed: %s",
                label,
                own_id_or_batch,
                exc,
                screenshot_exc,
            )


def wait_for_ready_signal() -> None:
    print(
        "\nLog in through TUM eaccess / Shibboleth in the opened browser, reach the Orbis home page, "
        'then type "ready" here and press Enter.\n'
    )
    while True:
        answer = input("> ").strip().lower()
        if answer == "ready":
            return
        print('Waiting. Type "ready" only after Orbis is open and usable.')


def read_upload_file(path: Path) -> tuple[list[Any], list[UploadRow]]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=False)
    sheet = workbook.active
    values = list(sheet.iter_rows(values_only=True))
    if not values:
        return [], []

    header = list(values[0])
    header_map = {normalize_header(value): index for index, value in enumerate(header)}
    own_idx = header_map.get("ownid")
    company_idx = header_map.get("companyname")
    if own_idx is None or company_idx is None:
        raise ValueError(f"{path} must contain Own ID and Company name columns.")

    rows: list[UploadRow] = []
    for source_row, raw in enumerate(values[1:], start=2):
        row_values = list(raw)
        own_id = normalize_cell(row_values[own_idx])
        company = normalize_cell(row_values[company_idx])
        if not own_id or not company:
            continue
        rows.append(UploadRow(own_id, company, row_values, path.name, source_row))
    workbook.close()
    return header, rows


def write_upload_slice(header: list[Any], rows: list[UploadRow], path: Path) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append(header)
    for row in rows:
        sheet.append(row.values)
    workbook.save(path)
    workbook.close()


def load_reference_rows(path: Path) -> dict[str, dict[str, Any]]:
    if not path.exists() and path.name == "master_enriched_reference.xlsx":
        fallback = Path("files") / path.name
        if fallback.exists():
            path = fallback
    if not path.exists():
        logging.warning("Reference workbook not found at %s; decisions will require manual review.", path)
        return {}

    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    rows = sheet.iter_rows(values_only=True)
    try:
        headers = [normalize_cell(value) for value in next(rows)]
    except StopIteration:
        workbook.close()
        return {}
    if "Own ID" not in headers:
        workbook.close()
        raise ValueError(f"{path} must contain an Own ID column.")

    own_index = headers.index("Own ID")
    reference: dict[str, dict[str, Any]] = {}
    for raw in rows:
        own_id = normalize_cell(raw[own_index] if own_index < len(raw) else "")
        if not own_id:
            continue
        reference[own_id] = {
            header: raw[index] if index < len(raw) else ""
            for index, header in enumerate(headers)
            if header
        }
    workbook.close()
    return reference


def load_completed_ids(path: Path) -> set[str]:
    if not path.exists():
        return set()
    return {line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()}


def load_seen_result_keys(path: Path) -> set[tuple[str, str, str, str, str]]:
    if not path.exists():
        return set()
    seen: set[tuple[str, str, str, str, str]] = set()
    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            try:
                seen.add(result_key(json.loads(line)))
            except json.JSONDecodeError:
                continue
    return seen


def load_selected_bvdids(path: Path) -> dict[str, str]:
    selected_by_own_id: dict[str, str] = {}
    if not path.exists():
        return selected_by_own_id
    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                continue
            if record.get("source") != "decision":
                continue
            own_id = str(record.get("own_id") or "").strip()
            if not own_id:
                continue
            selected_bvdid = str(record.get("selected_bvdid") or "").strip()
            if selected_bvdid:
                selected_by_own_id[own_id] = selected_bvdid
            else:
                selected_by_own_id.pop(own_id, None)
    return selected_by_own_id


def result_key(result: dict[str, Any]) -> tuple[str, str, str, str, str]:
    if result.get("source") == "decision":
        status = "incomplete" if int(result.get("incomplete_candidate_count") or 0) else "complete"
        return ("decision", str(result.get("own_id", "")), status, "", "")
    return (
        str(result.get("own_id", "")),
        str(result.get("bvdid", "")),
        str(result.get("candidate_name", result.get("name", ""))),
        str(result.get("national_id", result.get("identifier", ""))),
        str(result.get("score_numeric", result.get("score", ""))),
    )


def chunks(items: list[UploadRow], size: int) -> Iterable[list[UploadRow]]:
    for start in range(0, len(items), size):
        yield items[start : start + size]


def click_first(page: Page, selectors: list[str], label: str) -> None:
    last_error: Exception | None = None
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            locator.wait_for(state="visible", timeout=10_000)
            locator.click(timeout=10_000)
            return
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"Could not click {label}") from last_error


def wait_for_any_selector(page: Page, selectors: list[str], timeout_ms: int) -> str:
    deadline = time.monotonic() + timeout_ms / 1000
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        for selector in selectors:
            try:
                if page.locator(selector).first.is_visible(timeout=1000):
                    return selector
            except Exception as exc:
                last_error = exc
        time.sleep(1)
    raise TimeoutError(f"Timed out waiting for any selector: {selectors}") from last_error


def is_execution_context_destroyed(exc: Exception) -> bool:
    return "execution context was destroyed" in str(exc).casefold()


def validate_results_page_dom_with_retries(page: Page, attempts: int = 3) -> tuple[bool, str]:
    last_reason = "Validation did not run."
    for attempt in range(1, attempts + 1):
        try:
            logging.info("Manual results page DOM validation attempt %s/%s.", attempt, attempts)
            ok, reason = validate_results_page_dom(page)
            logging.info("Manual results page DOM validation attempt %s/%s outcome: ok=%s reason=%s", attempt, attempts, ok, reason)
            return ok, reason
        except Exception as exc:
            if is_execution_context_destroyed(exc):
                last_reason = "Navigation was still in progress during DOM validation."
                logging.warning(
                    "Manual results page DOM validation attempt %s/%s hit transient navigation: %s",
                    attempt,
                    attempts,
                    exc,
                )
            else:
                last_reason = f"Could not inspect page DOM: {exc}"
                logging.warning(
                    "Manual results page DOM validation attempt %s/%s failed: %s",
                    attempt,
                    attempts,
                    exc,
                )

            if attempt < attempts:
                try:
                    page.wait_for_load_state("networkidle", timeout=8000)
                    logging.info("Manual results page DOM validation wait after attempt %s reached networkidle.", attempt)
                except Exception as wait_exc:
                    logging.warning(
                        "Manual results page DOM validation wait after attempt %s did not reach networkidle: %s",
                        attempt,
                        wait_exc,
                    )

    return False, last_reason


def read_page_size(page: Page) -> int | None:
    try:
        raw = page.locator("select#pageSize").first.input_value(timeout=3000)
        value = int(str(raw).strip())
        return value if value > 0 else None
    except Exception:
        return None


def detect_systemic_problem(page: Page, response_reason: str = "") -> str:
    if response_reason:
        return response_reason

    current_url = (page.url or "").casefold()
    login_url_markers = [
        "eaccess.tum.edu/login",
        "shibboleth",
        "/saml",
        "openid",
        "/idp/",
        "login.microsoftonline",
    ]
    if "orbismanda" in current_url:
        return f"redirected to Orbis M&A page ({page.url})"
    if any(marker in current_url for marker in login_url_markers):
        return f"redirected to login/SSO page ({page.url})"

    try:
        body_text = page.locator("body").inner_text(timeout=2000).casefold()
    except Exception:
        return ""
    for pattern in RATE_BLOCK_TEXT_PATTERNS:
        if pattern in body_text:
            return f"page text contains '{pattern}'"

    normalized_body = re.sub(r"\s+", " ", body_text).strip()
    captcha_phrase = re.search(
        r"\b(?:complete|solve|enter|load(?:ing)?|invalid)\s+(?:the\s+)?captcha\b"
        r"|\bcaptcha\s+(?:required|verification|challenge|failed|error)\b",
        normalized_body,
    )
    if captcha_phrase:
        return f"page text contains CAPTCHA challenge phrase '{captcha_phrase.group(0)}'"

    try:
        captcha_state = page.evaluate(
                """() => {
                    function visible(el) {
                        if (!el) return false;
                        const style = window.getComputedStyle(el);
                        const rect = el.getBoundingClientRect();
                        return style.visibility !== "hidden" && style.display !== "none"
                            && rect.width > 0 && rect.height > 0;
                    }
                    const captchaSelector = [
                        "iframe[src*='recaptcha' i]",
                        "iframe[src*='hcaptcha' i]",
                        "iframe[src*='captcha' i][title*='challenge' i]",
                        "form[id*='captcha' i]",
                        "form[class*='captcha' i]",
                        "input[name*='captcha' i]",
                        "textarea[name*='captcha' i]",
                        "[class~='g-recaptcha']",
                        "[class~='h-captcha']",
                        "[data-sitekey][data-callback]",
                        "[role='dialog'][aria-label*='captcha' i]"
                    ].join(",");
                    const usableOrbisSelector = [
                        "select#pageSize",
                        "input[data-action='navigate'][data-type='int']",
                        "img[data-action='next'][data-next]",
                        "tr.poppingRow",
                        "div.owSnapshot",
                        "td[data-map-trigger='true']"
                    ].join(",");
                    return {
                        captchaControl: [...document.querySelectorAll(captchaSelector)].some(visible),
                        usableOrbis: [...document.querySelectorAll(usableOrbisSelector)].some(visible)
                    };
                }"""
            )
    except Exception:
        captcha_state = {}
    visible_captcha_control = bool(captcha_state.get("captchaControl")) if isinstance(captcha_state, dict) else False
    usable_orbis_context = bool(captcha_state.get("usableOrbis")) if isinstance(captcha_state, dict) else False
    if visible_captcha_control and not usable_orbis_context:
        return "page contains a visible CAPTCHA control"
    if visible_captcha_control and usable_orbis_context:
        logging.debug("Ignoring CAPTCHA-like control because a usable Orbis results/snapshot context is visible.")

    # A challenge/interstitial can consist of only the word "CAPTCHA" without
    # exposing a recognizable widget. Treat it as blocking only on a short page,
    # not when the word appears incidentally in a normal results page.
    if re.search(r"\bcaptcha\b", normalized_body) and len(normalized_body) <= 500:
        return "short interstitial page contains 'captcha'"
    return ""


def diagnostic_excerpt(body_text: str, reason: str, radius: int = 600) -> str:
    normalized = re.sub(r"\s+", " ", str(body_text or "")).strip()
    if not normalized:
        return "<empty page body>"
    needles = [
        "captcha",
        "verify you are human",
        "verify that you are human",
        "too many requests",
        "temporarily blocked",
        "unusual traffic",
        "automated queries",
    ]
    lowered = normalized.casefold()
    positions = [lowered.find(needle) for needle in needles if lowered.find(needle) >= 0]
    if positions:
        center = min(positions)
        start = max(0, center - radius)
        end = min(len(normalized), center + radius)
        prefix = "..." if start else ""
        suffix = "..." if end < len(normalized) else ""
        return f"{prefix}{normalized[start:end]}{suffix}"
    return normalized[: radius * 2] + ("..." if len(normalized) > radius * 2 else "")


def advance_to_next_results_page(
    page: Page,
    previous_names: set[str],
    timeout_ms: int,
    before_click: Any | None = None,
    after_click: Any | None = None,
) -> bool:
    if not wait_for_results_page_settled(page, timeout_ms):
        raise StopRun(
            "Current results page did not settle before the page-advance check; "
            "the batch was left incomplete.",
            stop_reason="pagination failure",
        )

    raw_current_page, raw_total_pages, page_label = read_results_page_indicator(page)
    current_page, total_pages = parse_results_page_label(page_label)
    next_button = page.locator("img[data-action='next'][data-next], img[aria-label='Next page'][data-next]").first
    try:
        if next_button.count() == 0 or not next_button.is_visible(timeout=1000):
            logging.info(
                "page advance check: current=%s total=%s raw_current=%r raw_total=%r "
                "next_aria_disabled=%r next_src=%r",
                current_page if current_page is not None else "unknown",
                total_pages if total_pages is not None else "unknown",
                raw_current_page,
                raw_total_pages,
                "<missing>",
                "<missing>",
            )
            if current_page is not None and total_pages is not None and current_page < total_pages:
                logging.warning(
                    "Page indicator reports %s / %s, but the Next results-page control is missing or hidden; waiting for it to become ready.",
                    current_page,
                    total_pages,
                )
                try:
                    next_button.wait_for(state="visible", timeout=timeout_ms)
                except Exception as exc:
                    raise StopRun(
                        "Next results-page control did not become visible even though "
                        f"the page indicator shows {current_page} / {total_pages}; the batch was left incomplete.",
                        stop_reason="pagination failure",
                    ) from exc
            else:
                return False

        next_aria_disabled = get_locator_attribute(next_button, "aria-disabled")
        next_src = get_locator_attribute(next_button, "src").replace("\\", "/")
        logging.info(
            "page advance check: current=%s total=%s raw_current=%r raw_total=%r "
            "next_aria_disabled=%r next_src=%r",
            current_page if current_page is not None else "unknown",
            total_pages if total_pages is not None else "unknown",
            raw_current_page,
            raw_total_pages,
            next_aria_disabled,
            next_src,
        )
        next_looks_disabled = (
            next_aria_disabled.casefold() == "true"
            or "/next/blocked" in next_src.casefold()
        )

        if current_page is not None and total_pages is not None:
            if current_page >= total_pages:
                logging.debug(
                    "Results page indicator is %s / %s; no further result pages are available.",
                    current_page,
                    total_pages,
                )
                return False
            if next_looks_disabled:
                logging.warning(
                    "Page indicator reports %s / %s, overriding transient disabled state on the Next control.",
                    current_page,
                    total_pages,
                )
                try:
                    page.wait_for_function(
                        """() => {
                            const next = document.querySelector(
                                "img[data-action='next'][data-next], img[aria-label='Next page'][data-next]"
                            );
                            if (!next) return false;
                            const ariaDisabled = String(next.getAttribute("aria-disabled") || "").toLowerCase();
                            const src = String(next.getAttribute("src") || "").replace(/\\\\/g, "/").toLowerCase();
                            return ariaDisabled !== "true" && !src.includes("/next/blocked");
                        }""",
                        timeout=timeout_ms,
                    )
                    next_button = page.locator(
                        "img[data-action='next'][data-next], img[aria-label='Next page'][data-next]"
                    ).first
                except Exception:
                    logging.warning(
                        "Next control did not report an enabled state even though the page indicator shows more pages; attempting the advance."
                    )
        elif next_looks_disabled:
            logging.debug(
                "Next results-page control is disabled and no page-count indicator was readable; no further result pages are assumed."
            )
            return False
    except StopRun:
        raise
    except Exception as exc:
        raise StopRun(
            f"Could not inspect the Next results-page control: {exc}; the batch was left incomplete.",
            stop_reason="pagination failure",
        ) from exc

    try:
        signal = page.evaluate(
            """() => {
                function visible(el) {
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                }
                function text(el) {
                    return String(el.getAttribute("title") || el.innerText || el.textContent || "").replace(/\\s+/g, " ").trim();
                }
                const first = [...document.querySelectorAll("div[data-id='Name']:not([data-snapshot-bvdid]), div[title]:not([data-snapshot-bvdid])")]
                    .filter((el) => visible(el) && !el.closest("div.owSnapshot") && !el.closest("tr.poppingRow"))
                    .map(text)
                    .filter(Boolean)[0] || "";
                const next = document.querySelector("img[data-action='next'][data-next], img[aria-label='Next page'][data-next]");
                return {
                    firstName: first,
                    pageNumber: next ? (next.getAttribute("data-pagenumber") || "") : ""
                };
            }"""
        )
        first_name = str(signal.get("firstName") or "")
        page_number = str(signal.get("pageNumber") or "")
        if before_click:
            before_click()
        next_button.click(timeout=4000)
        if after_click:
            after_click()
        page.wait_for_function(
            """({previousNames, pageNumber}) => {
                function visible(el) {
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                }
                function text(el) {
                    return String(el.getAttribute("title") || el.innerText || el.textContent || "").replace(/\\s+/g, " ").trim();
                }
                const currentNames = [...new Set([...document.querySelectorAll("div[data-id='Name']:not([data-snapshot-bvdid]), div[title]:not([data-snapshot-bvdid])")]
                    .filter((el) => visible(el) && !el.closest("div.owSnapshot") && !el.closest("tr.poppingRow"))
                    .map(text)
                    .filter(Boolean))].sort();
                const priorNames = [...previousNames].sort();
                const next = document.querySelector("img[data-action='next'][data-next], img[aria-label='Next page'][data-next]");
                const currentPageNumber = next ? (next.getAttribute("data-pagenumber") || "") : "";
                return JSON.stringify(currentNames) !== JSON.stringify(priorNames)
                    || currentPageNumber !== pageNumber;
            }""",
            arg={"previousNames": sorted(previous_names), "pageNumber": page_number},
            timeout=timeout_ms,
        )
        if not wait_for_results_page_settled(page, timeout_ms, previous_names=previous_names):
            raise StopRun(
                "Results pagination changed, but the new page did not reach a stable company-name set; "
                "the batch was left incomplete.",
                stop_reason="pagination failure",
            )
    except StopRun:
        raise
    except Exception as exc:
        raise StopRun(
            f"Could not advance results pagination or detect page change: {exc}; the batch was left incomplete.",
            stop_reason="pagination failure",
        ) from exc

    new_names = set(collect_visible_company_names(page))
    if new_names == previous_names:
        raise StopRun(
            "The Next-page action completed without changing the visible result set; the batch was left incomplete.",
            stop_reason="pagination failure",
        )
    return True


def resolve_company_row_across_result_pages(
    page: Page,
    row: UploadRow,
    max_pages: int,
    timeout_ms: int,
    before_click: Any | None = None,
    after_click: Any | None = None,
    log_unresolved: bool = True,
) -> ElementHandle | None:
    current_names = collect_visible_company_names(page)
    current_page_label = results_page_label(page, fallback="current")
    company_row = resolve_company_row(page, row)
    log_company_page_check(row, current_page_label, current_names, company_row is not None)
    if company_row is not None:
        return company_row

    pages_checked = 1
    bounded_pages = max(1, min(int(max_pages or 1), 20))
    while pages_checked < bounded_pages:
        previous_names = set(collect_visible_company_names(page))
        advanced = advance_to_next_results_page(
            page,
            previous_names,
            timeout_ms,
            before_click=before_click,
            after_click=after_click,
        )
        if not advanced:
            break
        pages_checked += 1
        current_names = collect_visible_company_names(page)
        current_page_label = results_page_label(page, fallback=str(pages_checked))
        company_row = resolve_company_row(page, row)
        log_company_page_check(row, current_page_label, current_names, company_row is not None)
        if company_row is not None:
            logging.info(
                "own_id=%s: located company after advancing results pagination (%s page(s) checked).",
                row.own_id,
                pages_checked,
            )
            return company_row

    if log_unresolved:
        logging.warning(
            "own_id=%s: company %r remained unresolved after checking %s result page(s).",
            row.own_id,
            row.company,
            pages_checked,
        )
    return None


def company_name_is_visible_on_current_page(page: Page, row: UploadRow) -> bool:
    expected = normalize_loose(row.company)
    return bool(expected and expected in {normalize_loose(name) for name in collect_visible_company_names(page)})


def wait_for_results_page_settled(
    page: Page,
    timeout_ms: int,
    previous_names: set[str] | None = None,
) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    previous = set(previous_names or set())
    last_names: tuple[str, ...] = ()
    stable_since: float | None = None
    while time.monotonic() < deadline:
        current_names = tuple(collect_visible_company_names(page))
        current_set = set(current_names)
        changed = not previous or current_set != previous
        if current_names and changed:
            if current_names == last_names:
                if stable_since is not None and time.monotonic() - stable_since >= 2.0:
                    return True
            else:
                last_names = current_names
                stable_since = time.monotonic()
        else:
            last_names = current_names
            stable_since = None
        time.sleep(0.5)
    return False


def results_page_label(page: Page, fallback: str) -> str:
    _raw_current, _raw_total, input_label = read_results_page_indicator(page)
    if input_label:
        return input_label
    try:
        label = page.evaluate(
            """() => {
                function visible(node) {
                    const style = window.getComputedStyle(node);
                    const rect = node.getBoundingClientRect();
                    return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                }
                for (const node of document.querySelectorAll("span, div, td, label")) {
                    if (!visible(node)) continue;
                    const text = String(node.innerText || node.textContent || "").replace(/\\s+/g, " ").trim();
                    if (/^\\d+\\s*\\/\\s*\\d+$/.test(text)) return text;
                }
                return "";
            }"""
        )
        if label:
            return str(label)
    except Exception:
        pass
    return fallback


def read_results_page_indicator(page: Page) -> tuple[str, str, str]:
    try:
        indicator = page.evaluate(
            """() => {
                function visible(node) {
                    const style = window.getComputedStyle(node);
                    const rect = node.getBoundingClientRect();
                    return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                }
                function text(node) {
                    return String(node ? (node.innerText || node.textContent || "") : "")
                        .replace(/\\s+/g, " ")
                        .trim();
                }

                const pageInput = [...document.querySelectorAll(
                    "input[data-action='navigate'][data-type='int']"
                )].find(visible);
                if (!pageInput) return {currentRaw: "", nearbyTexts: []};

                const nearbyTexts = [];
                let sibling = pageInput.nextSibling;
                for (let checked = 0; sibling && checked < 6; checked += 1, sibling = sibling.nextSibling) {
                    const siblingText = text(sibling);
                    if (siblingText) nearbyTexts.push(siblingText);
                }

                let container = pageInput.parentElement;
                for (let depth = 0; container && depth < 3; depth += 1, container = container.parentElement) {
                    const containerText = text(container);
                    if (containerText) nearbyTexts.push(containerText);
                }

                return {
                    currentRaw: String(pageInput.getAttribute("value") || "").trim(),
                    nearbyTexts
                };
            }"""
        )
        raw_current = str(indicator.get("currentRaw") or "").strip()
        for nearby_text in indicator.get("nearbyTexts") or []:
            match = re.search(r"/\s*(\d+)\b", str(nearby_text))
            if match:
                raw_total = match.group(0)
                return raw_current, raw_total, f"{raw_current} / {match.group(1)}" if raw_current else ""
        return raw_current, "", ""
    except Exception:
        return "", "", ""


def parse_results_page_label(label: str) -> tuple[int | None, int | None]:
    match = re.fullmatch(r"\s*(\d+)\s*/\s*(\d+)\s*", str(label or ""))
    if not match:
        return None, None
    return int(match.group(1)), int(match.group(2))


def log_company_page_check(
    row: UploadRow,
    page_label: str,
    visible_names: list[str],
    found: bool,
) -> None:
    first_visible = visible_names[0] if visible_names else "<none>"
    logging.info(
        "own_id=%s: results page %s: %s (first visible: %s)",
        row.own_id,
        page_label,
        "found" if found else "not found",
        first_visible,
    )


def return_to_first_results_page(
    page: Page,
    timeout_ms: int,
    before_click: Any | None = None,
    after_click: Any | None = None,
) -> bool:
    first_button = page.locator(
        "img[data-action='first'][data-first], img[aria-label='First page'], img[title='First page']"
    ).first
    try:
        if first_button.count() == 0 or not first_button.is_visible(timeout=1000):
            return False
        previous_names = set(collect_visible_company_names(page))
        if before_click:
            before_click()
        first_button.click(timeout=10_000)
        if after_click:
            after_click()
        if wait_for_results_page_settled(page, timeout_ms, previous_names=previous_names):
            logging.info(
                "Results pagination reset completed and settled on page %s.",
                results_page_label(page, fallback="1"),
            )
            return True
    except StopRun:
        raise
    except Exception as exc:
        logging.debug("Could not return results pagination to the first page: %s", exc)
    return False


def restore_queued_company_results_page(
    page: Page,
    row: UploadRow,
    max_pages: int,
    timeout_ms: int,
    before_click: Any | None = None,
    after_click: Any | None = None,
) -> ElementHandle | None:
    company_row = resolve_company_row_across_result_pages(
        page,
        row,
        max_pages=max_pages,
        timeout_ms=timeout_ms,
        before_click=before_click,
        after_click=after_click,
        log_unresolved=False,
    )
    if company_row is not None:
        return company_row

    logging.info(
        "own_id=%s: target company was not found on any forward results page; resetting pagination to page 1.",
        row.own_id,
    )
    if not return_to_first_results_page(
        page,
        timeout_ms=timeout_ms,
        before_click=before_click,
        after_click=after_click,
    ):
        return None

    company_row = resolve_company_row_across_result_pages(
        page,
        row,
        max_pages=max_pages,
        timeout_ms=timeout_ms,
        before_click=before_click,
        after_click=after_click,
        log_unresolved=True,
    )
    if company_row is not None:
        return company_row

    logging.warning(
        "own_id=%s: company was not found on any results page; returning pagination to page 1 for the next company.",
        row.own_id,
    )
    if not return_to_first_results_page(
        page,
        timeout_ms=timeout_ms,
        before_click=before_click,
        after_click=after_click,
    ):
        logging.warning("own_id=%s: could not restore page 1 after the all-pages search failed.", row.own_id)
    return None


def validate_results_page_dom(page: Page) -> tuple[bool, str]:
    current_url = page.url.casefold()
    if "orbismanda" in current_url or "manda.bvdinfo" in current_url:
        return False, "This appears to be an Orbis M&A page, not plain Orbis."

    try:
        summary = page.evaluate(
            """() => {
                function visible(el) {
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                }
                function countVisible(selector) {
                    return [...document.querySelectorAll(selector)].filter(visible).length;
                }
                const bodyText = (document.body ? document.body.innerText || document.body.textContent || "" : "").slice(0, 5000);
                return {
                    pageSize: countVisible("select#pageSize"),
                    companyNames: countVisible("div[data-id='Name']:not([data-snapshot-bvdid])"),
                    titledRows: countVisible("div[title]:not([data-snapshot-bvdid])"),
                    candidateNames: countVisible("div[data-id='Name'][data-snapshot-bvdid]"),
                    tables: countVisible("table"),
                    bodyText
                };
            }"""
        )
    except Exception as exc:
        if is_execution_context_destroyed(exc):
            raise
        return False, f"Could not inspect page DOM: {exc}"

    text = str(summary.get("bodyText") or "").casefold()
    if "shibboleth" in text or "tum eaccess" in text or "login" in text and not summary.get("companyNames"):
        return False, "This looks like a login/eaccess page, not loaded Orbis results."

    company_count = int(summary.get("companyNames") or 0)
    titled_count = int(summary.get("titledRows") or 0)
    candidate_count = int(summary.get("candidateNames") or 0)
    table_count = int(summary.get("tables") or 0)
    page_size_count = int(summary.get("pageSize") or 0)
    if company_count or (titled_count >= 5 and table_count) or candidate_count or page_size_count:
        return True, (
            f"recognized results DOM: companyNames={company_count}, titledRows={titled_count}, "
            f"candidateNames={candidate_count}, pageSizeSelectors={page_size_count}"
        )
    return False, "No recognizable Orbis Batch Search results table, company rows, candidate rows, or page-size selector was visible."


def collect_visible_company_names(page: Page) -> list[str]:
    try:
        names = page.evaluate(
            """() => {
                function visible(el) {
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                }
                function text(el) {
                    return String(el.getAttribute("title") || el.innerText || el.textContent || "").replace(/\\s+/g, " ").trim();
                }
                const cells = [...document.querySelectorAll("div[data-id='Name']:not([data-snapshot-bvdid]), div[title]:not([data-snapshot-bvdid])")]
                    .filter((el) => visible(el) && !el.closest("div.owSnapshot") && !el.closest("tr.poppingRow"))
                    .map(text)
                    .filter(Boolean);
                return [...new Set(cells)];
            }"""
        )
    except Exception:
        return []
    return [str(name) for name in names]


def match_visible_rows_to_upload_rows(visible_names: list[str], rows: list[UploadRow]) -> list[UploadRow]:
    if not visible_names:
        return rows
    by_norm: dict[str, list[UploadRow]] = {}
    for row in rows:
        by_norm.setdefault(normalize_loose(row.company), []).append(row)

    matched: list[UploadRow] = []
    seen_ids: set[str] = set()
    # Queueing must use the same exact-match rule as the pre-processing visibility
    # recheck, or loosely matched rows are queued only to fail that recheck.
    for name in visible_names:
        normalized = normalize_loose(name)
        candidates = by_norm.get(normalized, [])
        if len(candidates) == 1 and candidates[0].own_id not in seen_ids:
            matched.append(candidates[0])
            seen_ids.add(candidates[0].own_id)
    return matched


def validate_plain_orbis_url(current_url: str) -> bool:
    return explain_plain_orbis_url(current_url)[0]


def explain_plain_orbis_url(current_url: str) -> tuple[bool, str]:
    lowered = current_url.casefold()
    parsed = urlparse(current_url)
    host = parsed.netloc.casefold()
    if "orbismanda" in lowered or "manda.bvdinfo" in lowered:
        return False, "URL appears to be Orbis M&A, not plain Orbis"
    if not host:
        return False, "URL has no host"
    plain_orbis_patterns = [
        r"(^|\.)orbis\d*\.bvdinfo\.com$",
        r"^orbis-r\d+-bvdinfo-com(\.|$)",
        r"^orbis\d*-bvdinfo-com(\.|$)",
    ]
    if any(re.search(pattern, host) for pattern in plain_orbis_patterns):
        return True, "plain Orbis host accepted"
    if "tum-eaccess.de" in host and "orbis" in host and "bvdinfo" in host:
        return True, "TUM proxy plain Orbis-like host accepted"
    return False, f"host {host!r} does not look like a plain Orbis host"


def build_orbis_version_url(current_url: str, suffix: str) -> str:
    parsed = urlparse(current_url)
    if not parsed.scheme or not parsed.netloc:
        raise ValueError(f"Cannot derive Orbis origin from current URL: {current_url}")
    match = re.search(r"(/version-[^/]+/)", parsed.path)
    if not match:
        raise ValueError(f"Cannot find Orbis version prefix in current URL: {current_url}")
    origin = f"{parsed.scheme}://{parsed.netloc}"
    return f"{origin}{match.group(1)}{suffix.lstrip('/')}"


def build_batch_search_url(current_url: str) -> str:
    return build_orbis_version_url(current_url, "Orbis/1/Companies/BatchSearch/Start")


def close_open_candidate_tray(page: Page) -> None:
    try:
        close_snapshot_popup(page)
    except Exception:
        pass
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    try:
        page.locator("body").click(position={"x": 5, "y": 5}, timeout=3000)
    except Exception:
        pass
    cleanup_snapshot_dialogs(page, remove_visible=True)


def resolve_company_row(page: Page, row: UploadRow) -> ElementHandle | None:
    exact_selector = f"div[title={json.dumps(row.company)}]:not([data-snapshot-bvdid])"
    exact = page.locator(exact_selector)
    exact_count = exact.count()
    if exact_count == 1:
        logging.info("own_id=%s: resolved company row by exact title.", row.own_id)
        return exact.first.element_handle(timeout=5000)
    if exact_count > 1:
        row_locations = exact.evaluate_all(
            """(cells) => {
                function text(node) {
                    if (!node) return "";
                    return String(node.getAttribute("title") || node.innerText || node.textContent || "")
                        .replace(/\\s+/g, " ")
                        .trim();
                }
                function locationCell(row, dataId, className) {
                    if (!row) return null;
                    return row.querySelector(
                        `div[data-id='${dataId}'], td.${className} div[title], td.${className}, .${className}`
                    );
                }
                const seenRows = new Set();
                const rows = [];
                cells.forEach((cell, index) => {
                    const companyRow = cell.closest("tr");
                    const rowIdentity = companyRow || cell;
                    if (seenRows.has(rowIdentity)) return;
                    seenRows.add(rowIdentity);
                    rows.push({
                        index,
                        city: text(locationCell(companyRow, "City", "City")),
                        country: text(locationCell(companyRow, "Country", "Country"))
                    });
                });
                return rows;
            }"""
        )
        if len(row_locations) == 1:
            logging.info(
                "own_id=%s: resolved company row by exact title after collapsing %s matching DOM elements in the same result row.",
                row.own_id,
                exact_count,
            )
            return exact.nth(int(row_locations[0]["index"])).element_handle(timeout=5000)

        reference_country = normalize_country_for_row_resolution(row.reference_country)
        reference_city = normalize_loose(row.reference_city)
        country_matches = [
            location
            for location in row_locations
            if reference_country
            and normalize_country_for_row_resolution(location.get("country")) == reference_country
        ]
        country_city_matches = [
            location
            for location in country_matches
            if reference_city and normalize_loose(location.get("city")) == reference_city
        ]
        resolved_location: dict[str, Any] | None = None
        resolution_detail = ""
        if len(country_city_matches) == 1:
            resolved_location = country_city_matches[0]
            resolution_detail = "country+city"
        elif len(country_matches) == 1:
            resolved_location = country_matches[0]
            resolution_detail = "country"

        if resolved_location is not None:
            logging.info(
                "own_id=%s: resolved ambiguous row for %s via country/city match (%s).",
                row.own_id,
                row.company,
                resolution_detail,
            )
            return exact.nth(int(resolved_location["index"])).element_handle(timeout=5000)

        logging.warning(
            "own_id=%s: ambiguity remained after country/city disambiguation; skipping for manual review "
            "(company=%r, unique exact-title rows=%s, matching DOM elements=%s, reference country=%r, reference city=%r).",
            row.own_id,
            row.company,
            len(row_locations),
            exact_count,
            row.reference_country,
            row.reference_city,
        )
        return None

    summary = page.evaluate(
        """({company}) => {
            const wantedName = normalizeText(company);
            const wantedTokens = tokenSet(company);
            const nameCells = collectCompanyNameCells();

            function normalizeText(value) {
                return String(value || "")
                    .toLowerCase()
                    .replace(/[\\p{P}\\p{S}]+/gu, " ")
                    .replace(/\\s+/g, " ")
                    .trim();
            }
            function tokenSet(value) {
                return new Set(normalizeText(value).split(" ").filter(Boolean));
            }
            function cellName(cell) {
                return normalizeText(cell ? (cell.getAttribute("title") || cell.innerText || cell.textContent) : "");
            }
            function cellText(cell) {
                return cell ? (cell.innerText || cell.textContent || cell.getAttribute("title") || "") : "";
            }
            function wordOverlap(cell) {
                if (!wantedTokens.size) return 0;
                const rowText = cellText(cell.closest("tr") || cell);
                const rowTokens = tokenSet(rowText);
                let hits = 0;
                for (const token of wantedTokens) {
                    if (rowTokens.has(token)) hits += 1;
                }
                return hits / wantedTokens.size;
            }
            function collectCompanyNameCells() {
                const seeds = [
                    ...document.querySelectorAll("div[data-id='Name']:not([data-snapshot-bvdid])"),
                    ...document.querySelectorAll("div[title]:not([data-snapshot-bvdid])")
                ].filter((el) => !el.closest("div.owSnapshot") && !el.closest("[aria-hidden='true']"));
                const seen = new Set();
                const collected = [];
                for (const seed of seeds) {
                    if (seen.has(seed)) continue;
                    seen.add(seed);
                    collected.push(seed);
                }
                return collected;
            }

            const exactNameMatches = nameCells
                .map((cell, index) => ({cell, index, name: cellName(cell), overlap: wordOverlap(cell)}))
                .filter((item) => item.name === wantedName);
            if (exactNameMatches.length) {
                return {mode: "normalized_name", count: exactNameMatches.length, index: exactNameMatches.length === 1 ? exactNameMatches[0].index : -1};
            }

            const containsMatches = nameCells
                .map((cell, index) => ({cell, index, name: cellName(cell), overlap: wordOverlap(cell)}))
                .filter((item) => item.overlap >= 0.8);
            if (containsMatches.length) {
                const bestOverlap = Math.max(...containsMatches.map((item) => item.overlap));
                return {
                    mode: "contains_name",
                    count: containsMatches.length,
                    index: containsMatches.length === 1 ? containsMatches[0].index : -1,
                    overlap_percent: Math.round(bestOverlap * 100)
                };
            }

            return {mode: "none", count: 0, index: -1, overlap_percent: 0};
        }""",
        {"company": row.company},
    )

    mode = str(summary.get("mode", "none"))
    count = int(summary.get("count", 0))
    index = int(summary.get("index", -1))
    overlap_percent = int(summary.get("overlap_percent", 0))
    if count != 1 or index < 0:
        if count > 1:
            if mode == "contains_name":
                logging.warning(
                    "own_id=%s: contains_name: %s rows meet 80%% threshold; ambiguous, skipping.",
                    row.own_id,
                    count,
                )
            else:
                logging.warning(
                    "own_id=%s: ambiguous result row lookup by %s returned %s rows for company %r; skipping for manual review.",
                    row.own_id,
                    mode,
                    count,
                    row.company,
                )
        else:
            logging.debug("own_id=%s: no result row found for company %r on the current results page.", row.own_id, row.company)
        return None

    if mode == "contains_name":
        logging.info(
            "own_id=%s: resolved company row by contains_name with %s%% word overlap.",
            row.own_id,
            overlap_percent,
        )
    else:
        logging.info("own_id=%s: resolved company row by %s.", row.own_id, mode)

    handle = page.evaluate_handle(
        """({targetIndex}) => {
            const seeds = [
                ...document.querySelectorAll("div[data-id='Name']:not([data-snapshot-bvdid])"),
                ...document.querySelectorAll("div[title]:not([data-snapshot-bvdid])")
            ].filter((el) => !el.closest("div.owSnapshot") && !el.closest("[aria-hidden='true']"));
            const seen = new Set();
            const cells = [];
            for (const seed of seeds) {
                if (seen.has(seed)) continue;
                seen.add(seed);
                cells.push(seed);
            }
            return cells[targetIndex] || null;
        }""",
        {"targetIndex": index},
    )
    return handle.as_element()


def wait_for_candidate_tray(
    page: Page,
    timeout_ms: int,
    company_row: ElementHandle | None = None,
    row: UploadRow | None = None,
) -> Locator:
    if company_row is not None:
        marker = f"codex-current-tray-{int(time.time() * 1000)}-{random.randint(1000, 9999)}"
        deadline = time.monotonic() + timeout_ms / 1000
        while time.monotonic() < deadline:
            try:
                found = company_row.evaluate(
                    """(el, args) => {
                        function visible(node) {
                            const style = window.getComputedStyle(node);
                            const rect = node.getBoundingClientRect();
                            return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                        }
                        document.querySelectorAll("tr.poppingRow[data-codex-current-tray]").forEach((node) => {
                            node.removeAttribute("data-codex-current-tray");
                        });
                        const companyTr = el.closest("tr");
                        let node = companyTr ? companyTr.nextElementSibling : null;
                        while (node) {
                            const noResultCell = node.querySelector(args.noResultSelector);
                            const noResultLoaded = noResultCell
                                && String(noResultCell.innerText || noResultCell.textContent || "").replace(/\\s+/g, " ").trim().toLowerCase() === "no result";
                            if (
                                node.matches
                                && node.matches("tr.poppingRow")
                                && visible(node)
                                && (node.querySelector(args.selector) || noResultLoaded)
                            ) {
                                node.setAttribute("data-codex-current-tray", args.marker);
                                return true;
                            }
                            if (node.matches && !node.matches("tr.poppingRow") && node.querySelector("div[data-id='Name']:not([data-snapshot-bvdid])")) {
                                break;
                            }
                            node = node.nextElementSibling;
                        }
                        return false;
                    }""",
                    {
                        "selector": CANDIDATE_NAME_SELECTOR,
                        "noResultSelector": NO_RESULT_SELECTOR,
                        "marker": marker,
                    },
                )
                if found:
                    tray = page.locator(f"tr.poppingRow[data-codex-current-tray={json.dumps(marker)}]").first
                    tray.wait_for(state="visible", timeout=5000)
                    return tray
            except Exception:
                pass
            time.sleep(0.5)
        label = f" for own_id={row.own_id}" if row else ""
        raise TimeoutError(f"Timed out waiting for current company's candidate tray{label}.")

    tray = page.locator("tr.poppingRow").first
    try:
        tray.wait_for(state="attached", timeout=timeout_ms)
        tray.wait_for(state="visible", timeout=timeout_ms)
        deadline = time.monotonic() + timeout_ms / 1000
        while time.monotonic() < deadline:
            if tray.locator(CANDIDATE_NAME_SELECTOR).count() > 0 or tray_has_explicit_no_result(tray):
                return tray
            time.sleep(0.5)
        raise TimeoutError("Candidate tray appeared without candidates or an explicit No result state.")
    except Exception:
        raise


def tray_has_explicit_no_result(tray: Locator) -> bool:
    try:
        cell = tray.locator(NO_RESULT_SELECTOR).first
        return cell.count() > 0 and normalize_loose(cell.inner_text(timeout=2000)) == "no result"
    except Exception:
        return False


def locate_company_dropdown_chevron(page: Page, row: UploadRow) -> Locator | None:
    company_cell = resolve_company_row(page, row)
    if company_cell is None:
        return None
    marker = f"codex-company-dropdown-{int(time.time() * 1000)}-{random.randint(1000, 9999)}"
    found = company_cell.evaluate(
        """(el, args) => {
            function clean(value) {
                return String(value || "").replace(/\\s+/g, " ").trim().toLowerCase();
            }
            document.querySelectorAll("[data-codex-company-dropdown]").forEach((node) => {
                node.removeAttribute("data-codex-company-dropdown");
            });
            const companyTr = el.closest("tr");
            if (!companyTr) return false;
            const controls = [...companyTr.querySelectorAll(args.selector)];
            if (!controls.length) return false;
            const expectedName = clean(args.company);
            const matchingControl = controls.find((control) =>
                clean(control.getAttribute("aria-label")).startsWith(expectedName)
            );
            const control = matchingControl || controls[0];
            control.setAttribute("data-codex-company-dropdown", args.marker);
            return true;
        }""",
        {"selector": COMPANY_DROPDOWN_SELECTOR, "company": row.company, "marker": marker},
    )
    if not found:
        return None
    chevron = page.locator(f"[data-codex-company-dropdown={json.dumps(marker)}]").first
    return chevron if chevron.count() > 0 else None


def wait_for_company_chevron_collapsed(page: Page, row: UploadRow, timeout_ms: int) -> bool:
    try:
        chevron = locate_company_dropdown_chevron(page, row)
    except Exception:
        return False
    if chevron is None:
        return False

    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        try:
            if get_locator_attribute(chevron, "aria-expanded").casefold() == "false":
                return True
        except Exception:
            return False
        time.sleep(0.25)
    return False


def click_company_dropdown_chevron_with_fallback(
    page: Page,
    row: UploadRow,
    chevron: Locator,
    timeout_ms: int,
) -> bool:
    if get_locator_attribute(chevron, "aria-expanded").casefold() == "false":
        return True
    try:
        chevron.scroll_into_view_if_needed(timeout=3000)
        chevron.click(timeout=5000, force=True)
        if wait_for_company_chevron_collapsed(page, row, timeout_ms):
            logging.info("own_id=%s: dropdown-chevron collapse succeeded via Playwright force click.", row.own_id)
            return True
    except Exception as exc:
        logging.debug("own_id=%s: dropdown-chevron force click did not collapse tray: %s", row.own_id, exc)

    try:
        fresh_chevron = locate_company_dropdown_chevron(page, row)
        if fresh_chevron is None:
            return False
        handle = fresh_chevron.element_handle(timeout=3000)
        if handle is None:
            return False
        handle.evaluate("el => el.click()")
        if wait_for_company_chevron_collapsed(page, row, timeout_ms):
            logging.info("own_id=%s: dropdown-chevron collapse succeeded via native DOM click fallback.", row.own_id)
            return True
    except Exception as exc:
        logging.debug("own_id=%s: dropdown-chevron DOM click did not collapse tray: %s", row.own_id, exc)
    return False


def expanded_company_chevron_count(page: Page) -> int:
    try:
        return int(
            page.evaluate(
                """(selector) => {
                    function visible(node) {
                        const style = window.getComputedStyle(node);
                        const rect = node.getBoundingClientRect();
                        return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                    }
                    return [...document.querySelectorAll(selector)]
                        .filter((node) => visible(node) && node.getAttribute("aria-expanded") === "true")
                        .length;
                }""",
                COMPANY_DROPDOWN_SELECTOR,
            )
        )
    except Exception:
        return 0


def collapse_expanded_dropdown_chevrons(page: Page, timeout_ms: int) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    expanded_selector = f'{COMPANY_DROPDOWN_SELECTOR}[aria-expanded="true"]:visible'
    while time.monotonic() < deadline:
        before = expanded_company_chevron_count(page)
        if before == 0:
            return True
        target = page.locator(expanded_selector).first
        try:
            target.scroll_into_view_if_needed(timeout=2000)
            target.click(timeout=3000, force=True)
        except Exception:
            try:
                fresh_target = page.locator(expanded_selector).first
                handle = fresh_target.element_handle(timeout=2000)
                if handle is not None:
                    handle.evaluate("el => el.click()")
            except Exception:
                pass
        settle_deadline = min(deadline, time.monotonic() + 3)
        while time.monotonic() < settle_deadline:
            if expanded_company_chevron_count(page) < before:
                break
            time.sleep(0.25)
    return expanded_company_chevron_count(page) == 0


def visible_popping_tray_count(page: Page) -> int:
    try:
        return int(
            page.evaluate(
                """() => {
                    function visible(node) {
                        const style = window.getComputedStyle(node);
                        const rect = node.getBoundingClientRect();
                        return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                    }
                    return [...document.querySelectorAll("tr.poppingRow")].filter(visible).length;
                }"""
            )
        )
    except Exception:
        return page.locator("tr.poppingRow").count()


def wait_for_no_visible_popping_trays(page: Page, timeout_ms: int) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        if visible_popping_tray_count(page) == 0:
            return True
        time.sleep(0.5)
    return visible_popping_tray_count(page) == 0


def collapse_visible_popping_trays(page: Page, timeout_ms: int) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        if visible_popping_tray_count(page) == 0:
            return True
        try:
            clicked = page.evaluate(
                """() => {
                    function visible(node) {
                        if (!node) return false;
                        const style = window.getComputedStyle(node);
                        const rect = node.getBoundingClientRect();
                        return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                    }
                    function clickNode(node) {
                        if (!node) return false;
                        node.scrollIntoView({block: "center", inline: "nearest"});
                        node.click();
                        return true;
                    }
                    const trays = [...document.querySelectorAll("tr.poppingRow")].filter(visible);
                    for (const tray of trays) {
                        const companyRow = tray.previousElementSibling;
                        if (!companyRow) continue;
                        const controls = [
                            ...companyRow.querySelectorAll(
                                "img[role='button'][aria-expanded='true'][src*='Icons/Orbis/dropdown'], " +
                                "a[aria-expanded='true'], button[aria-expanded='true'], " +
                                "img[aria-label*='Collapse' i], img[title*='Collapse' i], img[alt*='Collapse' i], " +
                                "a[title*='Collapse' i], button[title*='Collapse' i], " +
                                ".collapse, .expanded, .toggle, .treeCollapse"
                            )
                        ].filter(visible);
                        if (controls.length && clickNode(controls[0])) return true;

                        const nameCell = companyRow.querySelector("div[data-id='Name']:not([data-snapshot-bvdid]), div[title]:not([data-snapshot-bvdid])");
                        if (visible(nameCell) && clickNode(nameCell)) return true;

                        if (visible(companyRow) && clickNode(companyRow)) return true;
                    }
                    return false;
                }"""
            )
            if clicked and wait_for_no_visible_popping_trays(page, timeout_ms=3000):
                return True
        except Exception as exc:
            logging.debug("Expanded-row collapse attempt failed: %s", exc)
        try:
            page.keyboard.press("Escape")
            page.locator("body").click(position={"x": 5, "y": 5}, timeout=1000)
        except Exception:
            pass
        time.sleep(0.75)
    return visible_popping_tray_count(page) == 0


def tray_contains_expected_candidate(tray: Locator, candidate: dict[str, Any]) -> bool:
    bvdid = str(candidate.get("bvdid") or "")
    if bvdid:
        try:
            if tray.locator(f"{CANDIDATE_NAME_SELECTOR}[data-snapshot-bvdid={json.dumps(bvdid)}]").count() > 0:
                return True
        except Exception:
            pass
    name = str(candidate.get("candidate_name") or "").strip()
    if name:
        try:
            return tray.locator(CANDIDATE_NAME_SELECTOR).filter(has_text=name).count() > 0
        except Exception:
            return False
    return False


def ensure_candidate_tray(
    page: Page,
    row: UploadRow,
    timeout_ms: int,
    expected_candidate: dict[str, Any] | None = None,
    max_pages: int = 1,
    before_page_click: Any | None = None,
    after_page_click: Any | None = None,
) -> Locator:
    visible_trays = page.locator("tr.poppingRow")
    try:
        for index in range(visible_trays.count()):
            tray = visible_trays.nth(index)
            if not tray.is_visible(timeout=500) or tray.locator(CANDIDATE_NAME_SELECTOR).count() == 0:
                continue
            if expected_candidate is None or tray_contains_expected_candidate(tray, expected_candidate):
                return tray
        if expected_candidate is not None and visible_trays.count() > 0:
            logging.warning(
                "Located tr.poppingRow does not contain expected candidates for own_id=%s; likely a stale row from a previous company.",
                row.own_id,
            )
    except Exception:
        pass

    close_open_candidate_tray(page)
    company_row = resolve_company_row_across_result_pages(
        page,
        row,
        max_pages=max_pages,
        timeout_ms=timeout_ms,
        before_click=before_page_click,
        after_click=after_page_click,
    )
    if company_row is None:
        raise RuntimeError(f"Could not re-open candidate tray for own_id={row.own_id} company={row.company}")
    company_row.click(timeout=30_000)
    tray = wait_for_candidate_tray(page, timeout_ms, company_row, row)
    if expected_candidate is not None and not tray_contains_expected_candidate(tray, expected_candidate):
        logging.warning(
            "Located tr.poppingRow does not contain expected candidates for own_id=%s; likely a stale row from a previous company.",
            row.own_id,
        )
        raise RuntimeError(f"Wrong candidate tray found for own_id={row.own_id}; expected candidate was not present.")
    return tray


def candidate_locator_in_tray(tray: Locator, candidate: dict[str, Any]) -> Locator:
    bvdid = str(candidate.get("bvdid") or "")
    if bvdid:
        locator = tray.locator(f"{CANDIDATE_NAME_SELECTOR}[data-snapshot-bvdid={json.dumps(bvdid)}]").first
        if locator.count():
            return locator

    name = str(candidate.get("candidate_name") or "")
    if name:
        return tray.locator(CANDIDATE_NAME_SELECTOR).filter(has_text=name).first
    return tray.locator(CANDIDATE_NAME_SELECTOR).first


def click_go_to_report_with_fallback(
    page: Page,
    go_to_report: Locator,
    row: UploadRow,
    candidate: dict[str, Any],
    results_timeout_ms: int,
) -> None:
    report_marker = page.locator("td[data-map-trigger='true']").first
    click_with_force_and_dom_fallback(
        go_to_report,
        report_marker,
        row,
        candidate,
        action_label="Go to report",
        success_label="report page",
        force_timeout_ms=5000,
        force_verify_timeout_ms=8000,
        fallback_verify_timeout_ms=results_timeout_ms,
        target_factory=lambda: page.locator("div.owSnapshot").first.locator("a[data-report-link], a:has-text('Go to report')").first,
        identity_attr="data-report-link",
        expected_identity=str(candidate.get("bvdid") or ""),
    )


def click_with_force_and_dom_fallback(
    target: Locator,
    success_locator: Locator,
    row: UploadRow,
    candidate: dict[str, Any],
    action_label: str,
    success_label: str,
    force_timeout_ms: int,
    force_verify_timeout_ms: int,
    fallback_verify_timeout_ms: int,
    target_factory: Any | None = None,
    identity_attr: str = "",
    expected_identity: str = "",
) -> None:
    try:
        target = ensure_click_target_identity(
            target,
            target_factory,
            identity_attr,
            expected_identity,
            row,
            candidate,
        )
        target.click(timeout=force_timeout_ms, force=True)
        success_locator.wait_for(state="visible", timeout=force_verify_timeout_ms)
        logging.info(
            "own_id=%s bvdid=%s: %s succeeded via Playwright force click.",
            row.own_id,
            candidate.get("bvdid", ""),
            action_label,
        )
        return
    except Exception as force_exc:
        logging.warning(
            "own_id=%s bvdid=%s: %s force click did not reach %s; falling back to DOM click: %s",
            row.own_id,
            candidate.get("bvdid", ""),
            action_label,
            success_label,
            force_exc,
        )

    target = ensure_click_target_identity(
        target,
        target_factory,
        identity_attr,
        expected_identity,
        row,
        candidate,
    )
    handle = target.element_handle(timeout=3000)
    if handle is None:
        raise RuntimeError(f"{action_label} target element could not be resolved for DOM click fallback.")
    handle.evaluate("el => el.click()")
    success_locator.wait_for(state="visible", timeout=fallback_verify_timeout_ms)
    logging.info(
        "own_id=%s bvdid=%s: %s succeeded via native DOM click fallback.",
        row.own_id,
        candidate.get("bvdid", ""),
        action_label,
    )


def ensure_click_target_identity(
    target: Locator,
    target_factory: Any | None,
    identity_attr: str,
    expected_identity: str,
    row: UploadRow,
    candidate: dict[str, Any],
) -> Locator:
    if not identity_attr or not expected_identity:
        return target

    found_identity = get_locator_attribute(target, identity_attr)
    if click_identity_matches(found_identity, expected_identity):
        return target

    logging.warning(
        "own_id=%s bvdid=%s: target element identity mismatch before click (expected %s, found %s) - likely stale DOM node from virtualized list; re-locating",
        row.own_id,
        candidate.get("bvdid", ""),
        expected_identity,
        found_identity,
    )
    if target_factory is None:
        raise RuntimeError(f"Click target identity mismatch: expected {expected_identity}, found {found_identity}.")

    fresh_target = target_factory()
    if fresh_target is None:
        raise RuntimeError(f"Could not re-locate click target after identity mismatch for expected {expected_identity}.")
    fresh_identity = get_locator_attribute(fresh_target, identity_attr)
    if not click_identity_matches(fresh_identity, expected_identity):
        raise RuntimeError(f"Re-located click target identity mismatch: expected {expected_identity}, found {fresh_identity}.")
    return fresh_target


def click_identity_matches(found_identity: str, expected_identity: str) -> bool:
    found = str(found_identity or "").strip()
    expected = str(expected_identity or "").strip()
    if found == expected:
        return True
    found_prefix = found.split("_", 1)[0]
    expected_prefix = expected.split("_", 1)[0]
    return bool(found_prefix and expected_prefix and found_prefix == expected_prefix)


def get_locator_attribute(locator: Locator, attribute: str) -> str:
    try:
        return str(locator.get_attribute(attribute, timeout=2000) or "")
    except Exception:
        return ""


def extract_report_fields(page: Page) -> dict[str, Any]:
    return page.evaluate(
        """() => {
            function clean(value) {
                return String(value || "").replace(/\\s+/g, " ").trim();
            }
            const address = [...document.querySelectorAll('td[data-map-trigger="true"] span[lang]')]
                .map((node) => clean(node.innerText || node.textContent))
                .filter(Boolean)
                .join(", ");
            const domainNode = document.querySelector('td[data-qc-id="Midget.DomainLink"] a');
            const domain = clean(domainNode ? (domainNode.innerText || domainNode.textContent) : "");
            const stateText = [...document.querySelectorAll("td.state")]
                .map((node) => node.innerText || node.textContent || "")
                .join("\\n");
            const orbisMatch = stateText.match(/Orbis ID:\\s*(\\d+)/i);
            const statusNode = document.querySelector("td.state span");
            const companyStatus = clean(statusNode ? (statusNode.innerText || statusNode.textContent) : "");
            const headquartersCell = document.querySelector("td[data-qc-id='Report.HeadQuaters']");
            let ownerName = "";
            if (headquartersCell && /Global Ultimate Owner/i.test(headquartersCell.innerText || headquartersCell.textContent || "")) {
                const ownerLink = headquartersCell.querySelector("a");
                ownerName = clean(ownerLink ? (ownerLink.getAttribute("title") || ownerLink.innerText || ownerLink.textContent) : "");
            }

            let management = "";
            const title = [...document.querySelectorAll("td.midgetTitle")]
                .find((node) => clean(node.innerText || node.textContent) === "Management");
            if (title) {
                const table = title.closest("table");
                const titleRow = title.closest("tr");
                const rows = table ? [...table.querySelectorAll("tr")] : [];
                const start = titleRow ? rows.indexOf(titleRow) + 1 : 0;
                management = rows.slice(start)
                    .map((tr) => clean(tr.innerText || tr.textContent))
                    .filter(Boolean)
                    .join("; ");
            }

            const reportFields = {};
            for (const tr of document.querySelectorAll("tr")) {
                const cells = [...tr.querySelectorAll("th, td")].map((cell) => clean(cell.innerText || cell.textContent)).filter(Boolean);
                if (cells.length >= 2 && cells[0].length <= 80) {
                    const key = cells[0];
                    const value = cells.slice(1).join(" | ");
                    if (key && value && !reportFields[key]) reportFields[key] = value;
                }
            }
            const reportText = clean(document.body ? (document.body.innerText || document.body.textContent) : "");
            const dateOfIncorporationEntry = Object.entries(reportFields)
                .find(([key]) => /date of incorporation/i.test(key));
            const dateOfIncorporation = dateOfIncorporationEntry
                ? clean(dateOfIncorporationEntry[1])
                : "";

            return {
                full_address: address,
                domain,
                website: domain,
                orbis_id: orbisMatch ? orbisMatch[1] : "",
                company_status: companyStatus,
                owner_name: ownerName,
                management,
                date_of_incorporation: dateOfIncorporation,
                report_fields: reportFields,
                report_text: reportText
            };
        }"""
    )


def warn_if_country_spread(row: UploadRow, countries: set[str]) -> None:
    normalized = {normalize_loose(country) for country in countries if normalize_loose(country)}
    if len(normalized) >= 4:
        logging.warning(
            "own_id=%s: candidates span %s countries (%s); manual review recommended.",
            row.own_id,
            len(normalized),
            ", ".join(sorted(countries)),
        )


def should_stop_for_candidate_cap(
    candidates: list[dict[str, Any]],
    ref_row: dict[str, Any] | None,
    soft_cap: int,
    hard_cap: int,
) -> bool:
    if len(candidates) >= hard_cap:
        return True
    if len(candidates) >= soft_cap and has_convincing_candidate(candidates, ref_row):
        return True
    return False


def cap_incomplete_reason(
    candidates: list[dict[str, Any]],
    ref_row: dict[str, Any] | None,
    soft_cap: int,
    hard_cap: int,
) -> str:
    if len(candidates) >= soft_cap and has_convincing_candidate(candidates, ref_row):
        return f"Not visited - candidate cap reached (matched within first {len(candidates)})"
    return f"Not visited - cap of {hard_cap} reached without a {MIN_WITNESSES}-witness match"


def incomplete_records_for_unvisited(
    candidates: list[dict[str, Any]],
    reason: str,
    reason_type: str = "cap_skip",
    status: str = "not_visited_candidate_cap",
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for candidate in candidates:
        records.append(
            {
                "candidate_index": candidate.get("candidate_index", ""),
                "bvdid": candidate.get("bvdid", ""),
                "candidate_name": candidate.get("candidate_name", ""),
                "error": reason,
                "status": status,
                "reason_type": reason_type,
            }
        )
    return records


def normalize_selected_label(value: Any) -> str:
    return re.sub(r"\s+", " ", normalize_cell(value)).casefold()


def read_orbis_auto_match_from_company_row(company_row: ElementHandle) -> dict[str, str]:
    return company_row.evaluate(
        """(el) => {
            function clean(value) {
                return String(value || "").replace(/\\s+/g, " ").trim();
            }
            function text(node) {
                return clean(node ? (node.getAttribute("title") || node.innerText || node.textContent) : "");
            }
            const tr = el.closest("tr");
            if (!tr) return {selected_name: "", identifier: "", score_letter: "", score_numeric: ""};
            const selectedNodes = [...tr.querySelectorAll("div#text[title]")];
            const selected = selectedNodes
                .map((node) => clean(node.getAttribute("title")))
                .find(Boolean) || "";
            const identifierNode =
                tr.querySelector("div[data-id='Identifier'][title], div[data-id='NationalId'][title], td.Identifier div[title], td.NationalId div[title], td.Identifier, td.NationalId");
            const score = tr.querySelector("span[data-id='Score'], td.Score span, .Score span, td.Score, .Score");
            return {
                selected_name: selected,
                identifier: text(identifierNode),
                score_letter: text(score),
                score_numeric: score ? clean(score.getAttribute("data-score")) : ""
            };
        }"""
    )


def find_candidate_by_selected_name(tray: Locator, selected_name: str) -> tuple[Locator | None, dict[str, Any] | None]:
    wanted = normalize_selected_label(selected_name)
    if not wanted:
        return None, None
    candidate_cells = tray.locator(CANDIDATE_NAME_SELECTOR)
    for index in range(candidate_cells.count()):
        cell = candidate_cells.nth(index)
        candidate = extract_candidate(cell)
        if normalize_selected_label(candidate.get("candidate_name", "")) == wanted:
            candidate["candidate_index"] = index + 1
            return cell, candidate
    return None, None


def extract_bvdid_from_snapshot(snapshot_text: str) -> str:
    match = re.search(r"BVD ID:\s*([A-Z0-9*]+)", snapshot_text or "", flags=re.IGNORECASE)
    return match.group(1) if match else ""


def extract_candidate(candidate_cell: Any) -> dict[str, Any]:
    return candidate_cell.evaluate(
        """(el) => {
            function text(node) {
                if (!node) return "";
                return (node.getAttribute("title") || node.innerText || node.textContent || "").trim();
            }
            function cell(row, dataId, className) {
                if (!row) return null;
                return row.querySelector(`div[data-id='${dataId}'], td.${className}, .${className}`);
            }
            const row = el.closest("tr");
            const score = row ? row.querySelector("span[data-id='Score'], td.Score span, .Score span, td.Score, .Score") : null;
            return {
                bvdid: el.getAttribute("data-snapshot-bvdid") || "",
                candidate_name: text(el),
                city: text(cell(row, "City", "City")),
                country: text(cell(row, "Country", "Country")),
                national_id: text(cell(row, "NationalId", "NationalId")),
                score_letter: text(score),
                score_numeric: score ? (score.getAttribute("data-score") || "") : ""
            };
        }"""
    )


def build_decision_record(
    row: UploadRow,
    batch: Batch,
    decision: Any,
    candidates: list[dict[str, Any]],
    incomplete_candidates: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    selected = decision.selected
    evidence = decision.evidence
    incomplete_candidates = incomplete_candidates or []
    return {
        "source": "decision",
        "own_id": row.own_id,
        "company": row.company,
        "source_file": row.source_file,
        "source_row": row.source_row,
        "batch_key": batch.key,
        "selected_bvdid": selected.get("bvdid", "") if selected else "",
        "selected_company_name": selected.get("candidate_name", "") if selected else "",
        "orbis_score_letter": selected.get("score_letter", "") if selected else "",
        "orbis_score_numeric": selected.get("score_numeric", "") if selected else "",
        "matched_via": decision.matched_via,
        "reason_1": decision.reason_1,
        "reason_2": decision.reason_2,
        "comments": decision.comments,
        "evidence_count": evidence.count if evidence else 0,
        "all_evidence_used": list(evidence.witnesses) if evidence else [],
        "candidate_count": len(candidates),
        "incomplete_candidates": incomplete_candidates,
        "incomplete_candidate_count": len(incomplete_candidates),
        "best_candidate_if_no_match": summarize_candidate(decision.best_candidate_if_no_match),
        "decided_at": datetime.now().isoformat(timespec="seconds"),
    }


def build_orbis_auto_match_decision_record(
    row: UploadRow,
    batch: Batch,
    candidate: dict[str, Any],
    auto_match: dict[str, str],
) -> dict[str, Any]:
    comments = (
        "Matched automatically by Orbis Batch Search using the supplied National ID (trade register number). "
        "Remaining candidates were not collected and our two-witness rule was not independently applied."
    )
    return {
        "source": "decision",
        "decision_type": "orbis_national_id_auto_match",
        "own_id": row.own_id,
        "company": row.company,
        "source_file": row.source_file,
        "source_row": row.source_row,
        "batch_key": batch.key,
        "selected_bvdid": candidate.get("bvdid", ""),
        "selected_company_name": auto_match.get("selected_name", "") or candidate.get("candidate_name", ""),
        "orbis_selected_name": auto_match.get("selected_name", ""),
        "orbis_row_identifier": auto_match.get("identifier", ""),
        "orbis_score_letter": candidate.get("score_letter", ""),
        "orbis_score_numeric": candidate.get("score_numeric", ""),
        "matched_via": "Orbis National ID auto-match",
        "reason_1": "Matched automatically by Orbis Batch Search using the supplied National ID (trade register number).",
        "reason_2": "Remaining candidates were not collected and our two-witness rule was not independently applied.",
        "comments": comments,
        "evidence_count": 0,
        "all_evidence_used": [],
        "candidate_count": 1,
        "incomplete_candidates": [],
        "incomplete_candidate_count": 0,
        "auto_match_candidate": candidate,
        "best_candidate_if_no_match": "",
        "decided_at": datetime.now().isoformat(timespec="seconds"),
    }


def summarize_candidate(candidate: dict[str, Any] | None) -> dict[str, Any] | str:
    if not candidate:
        return ""
    return {
        "bvdid": candidate.get("bvdid", ""),
        "candidate_name": candidate.get("candidate_name", ""),
        "orbis_score_letter": candidate.get("score_letter", ""),
        "orbis_score_numeric": candidate.get("score_numeric", ""),
    }


def close_snapshot_popup(page: Page) -> None:
    selectors = [
        "div.owSnapshot button[title='Close']",
        "button[title='Close']",
        "a[title='Close']",
        ".ui-dialog-titlebar-close",
        "button:has-text('Close')",
    ]
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            if locator.is_visible(timeout=1000):
                locator.click(timeout=3000)
                cleanup_snapshot_dialogs(page, remove_visible=True)
                return
        except Exception:
            continue
    try:
        page.keyboard.press("Escape")
    finally:
        cleanup_snapshot_dialogs(page, remove_visible=True)


def cleanup_snapshot_dialogs(page: Page, remove_visible: bool = False) -> None:
    try:
        removed = page.evaluate(
            """({removeVisible}) => {
                function visible(el) {
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
                }
                const selectors = [
                    ".owSnapshot",
                    ".ui-dialog:has(.owSnapshot)",
                    ".openerContainer.ui-draggable.ui-draggable-handle[role='dialog']"
                ];
                const nodes = [...document.querySelectorAll(selectors.join(","))];
                let removed = 0;
                for (const node of nodes) {
                    if (!removeVisible && visible(node)) continue;
                    node.remove();
                    removed += 1;
                }
                return removed;
            }""",
            {"removeVisible": remove_visible},
        )
        if removed:
            logging.info("Removed %s stale snapshot/dialog DOM element(s).", removed)
    except Exception as exc:
        logging.warning("Could not inspect/remove stale snapshot dialogs: %s", exc)


def human_pause(min_seconds: float, max_seconds: float) -> None:
    low = min(min_seconds, max_seconds)
    high = max(min_seconds, max_seconds)
    if high <= low:
        pause_seconds = max(0.0, low)
    else:
        split = low + (high - low) * 0.4
        if random.random() < 0.8:
            pause_seconds = random.uniform(low, split)
        else:
            pause_seconds = random.uniform(split, high)

    if random.random() < 0.05:
        pause_seconds += random.uniform(20.0, 60.0)
    time.sleep(pause_seconds)


def random_company_rest_threshold(min_companies: int, max_companies: int) -> int:
    low = max(1, min(min_companies, max_companies))
    high = max(low, max(min_companies, max_companies))
    return random.randint(low, high)


def normalize_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", normalize_cell(value).lower())


def normalize_loose(value: Any) -> str:
    return re.sub(r"\s+", " ", normalize_cell(value).casefold()).strip()


def normalize_country_for_row_resolution(value: Any) -> str:
    normalized = normalize_loose(value)
    if not normalized:
        return ""
    if len(normalized) == 2 and normalized.isalpha():
        return normalized.upper()
    return COUNTRY_TO_ISO2.get(normalized, normalized)


def natural_key(path: Path) -> list[Any]:
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", path.name)]


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_.-]+", "_", value).strip("._")
    return cleaned[:180] or "item"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scrape Orbis Batch Search candidates with resume-safe checkpoints.")
    parser.add_argument("--input-dir", default="files", help="Directory containing orbis_upload_*.xlsx files.")
    parser.add_argument("--reference", default="files/master_enriched_reference.xlsx", help="Reference workbook for per-company matching decisions.")
    parser.add_argument("--run-dir", default="orbis_run", help="Directory for temp uploads, logs, screenshots.")
    parser.add_argument("--output", default="orbis_run/results.jsonl", help="Crash-safe JSONL result output.")
    parser.add_argument("--checkpoint", default="orbis_run/completed_own_ids.txt", help="Completed own_id checkpoint file.")
    parser.add_argument("--user-data-dir", default="orbis_run/playwright_profile", help="Persistent browser profile directory.")
    parser.add_argument("--start-url", default=DEFAULT_START_URL, help="TUM eaccess / Orbis start URL.")
    parser.add_argument("--batch-size", type=int, default=100, help="Rows per generated upload workbook.")
    parser.add_argument("--only-file", action="append", help="Restrict to one or more orbis_upload_*.xlsx file names.")
    parser.add_argument(
        "--current-upload-file",
        help="Manual handoff safety: file name of the Excel workbook the human uploaded in Orbis.",
    )
    parser.add_argument("--max-batches", type=int, default=0, help="Stop after N batches, useful for smoke tests.")
    parser.add_argument(
        "--max-companies",
        type=int,
        default=0,
        help="Stop after N companies attempted in this run; useful for smoke tests.",
    )
    parser.add_argument(
        "--max-runtime-minutes",
        type=int,
        default=None,
        help="Soft wall-clock cap for per-company processing; finishes the current company, with a fixed 15-minute hard-stop buffer.",
    )
    parser.add_argument(
        "--start-from-current-page",
        "--attach-current-results-page",
        action="store_true",
        help="Do not login, navigate, upload, or map. Wait for the human to place the browser on a loaded Batch Search results page, then process it.",
    )
    parser.add_argument(
        "--open-start-url-in-manual-mode",
        action="store_true",
        help="With --start-from-current-page, open --start-url in the Playwright Chromium window before waiting for manual login/upload/results.",
    )
    parser.add_argument(
        "--reselect-checkpointed",
        action="store_true",
        help="Re-apply stored selected BvD IDs in the Orbis UI for checkpointed companies without collecting or writing candidate data.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Create/list pending 100-row uploads without launching browser.")
    parser.add_argument("--slow-mo", type=int, default=0, help="Playwright slow_mo in milliseconds.")
    parser.add_argument("--min-wait", type=float, default=0.5, help="Minimum randomized wait between clicks.")
    parser.add_argument("--max-wait", type=float, default=2.0, help="Maximum randomized wait between clicks.")
    parser.add_argument("--interaction-min-wait", type=float, default=5.0, help="Minimum cautious randomized wait around browser interactions.")
    parser.add_argument("--interaction-max-wait", type=float, default=13.0, help="Maximum cautious randomized wait around browser interactions.")
    parser.add_argument(
        "--candidate-soft-cap",
        type=int,
        default=None,
        help="Optional fixed per-company soft-cap override; disables the default randomized soft cap.",
    )
    parser.add_argument("--candidate-soft-cap-min", type=int, default=5, help="Minimum randomized candidate soft cap per company.")
    parser.add_argument("--candidate-soft-cap-max", type=int, default=10, help="Maximum randomized candidate soft cap per company.")
    parser.add_argument("--candidate-hard-cap", type=int, default=15, help="Never visit more than this many candidates per company.")
    parser.add_argument("--rest-every-min-companies", type=int, default=8, help="Minimum companies processed before a longer randomized rest break.")
    parser.add_argument("--rest-every-max-companies", type=int, default=12, help="Maximum companies processed before a longer randomized rest break.")
    parser.add_argument("--rest-min-seconds", type=float, default=300.0, help="Minimum longer rest-break duration in seconds.")
    parser.add_argument("--rest-max-seconds", type=float, default=600.0, help="Maximum longer rest-break duration in seconds.")
    parser.add_argument("--default-timeout-ms", type=int, default=30_000)
    parser.add_argument("--navigation-timeout-ms", type=int, default=90_000)
    parser.add_argument("--results-timeout-ms", type=int, default=300_000)
    parser.add_argument("--snapshot-timeout-ms", type=int, default=60_000)
    args = parser.parse_args()
    if args.candidate_hard_cap < 1:
        parser.error("--candidate-hard-cap must be at least 1")

    soft_min = max(1, args.candidate_soft_cap_min)
    soft_max = max(1, args.candidate_soft_cap_max)
    if soft_min > soft_max:
        soft_min, soft_max = soft_max, soft_min
    if soft_max > args.candidate_hard_cap:
        parser.error(
            "--candidate-hard-cap must be greater than or equal to --candidate-soft-cap-max "
            f"({args.candidate_hard_cap} < {soft_max})"
        )
    args.candidate_soft_cap_min = soft_min
    args.candidate_soft_cap_max = soft_max

    if args.candidate_soft_cap is not None:
        args.candidate_soft_cap = max(1, args.candidate_soft_cap)
        if args.candidate_soft_cap > args.candidate_hard_cap:
            parser.error(
                "--candidate-soft-cap fixed override must not exceed --candidate-hard-cap "
                f"({args.candidate_soft_cap} > {args.candidate_hard_cap})"
            )
    return args


if __name__ == "__main__":
    OrbisScraper(parse_args()).run()
