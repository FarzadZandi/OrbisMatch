Recommended manual handoff run from this folder:

python .\orbis_candidate_scraper.py --start-from-current-page --current-upload-file orbis_upload_0001_1000.xlsx --max-companies 3
In this mode, you manually open plain Orbis, navigate to Batch Search, upload the Excel file, apply mapping if needed, and wait until the matching results are visible. Then type ready in the terminal. The scraper validates the current page by DOM content, not by a hardcoded versioned URL. Before processing, it prints the assumed upload file, first/last expected own_id, and company count. Type yes only if this matches the file you manually uploaded.

Important: the scraper opens its own Playwright Chromium window using the profile at orbis_run/playwright_profile. It does not attach to your normal Chrome or Edge window. If that Playwright window starts blank, that is expected: manually open Orbis in that blank Playwright window and continue the login and upload workflow there.

Optional convenience: open the configured TUM/Orbis start URL in the Playwright window before waiting for manual work:

python .\orbis_candidate_scraper.py --start-from-current-page --open-start-url-in-manual-mode --current-upload-file orbis_upload_0001_1000.xlsx --max-companies 3
Normal automated upload mode is still available:

python .\orbis_candidate_scraper.py
The script opens a headed persistent Chromium profile at orbis_run/playwright_profile. In manual handoff mode, it does not open the login URL, navigate to Batch Search, upload files, or click mapping controls. It only waits for ready, validates that the current page looks like loaded Batch Search results, and starts collecting. The exception is --open-start-url-in-manual-mode, which only opens --start-url in the Playwright window to save typing; login, upload, mapping, and waiting for results remain manual. If --current-upload-file is omitted, the script warns that it will assume the first pending upload batch and still requires the same yes confirmation before processing.

Useful smoke test:

python .\orbis_candidate_scraper.py --start-from-current-page --current-upload-file orbis_upload_0001_1000.xlsx --max-companies 3
Useful dry run:

python .\orbis_candidate_scraper.py --dry-run
State and outputs:

orbis_run/results.jsonl: append-only raw candidate records plus one source: "decision" record per completed company.
orbis_run/completed_own_ids.txt: company-level checkpoint.
orbis_run/temp_uploads/: generated 100-row upload files by default.
orbis_run/screenshots/: failure screenshots.
orbis_run/logs/orbis_scraper.log: run log without credentials.
Confirmed rate-block, CAPTCHA, or session-loss stops also write a full-page screenshot and a short URL/text diagnostic to the screenshot and log folders.
CAPTCHA-like names or markup do not trigger a stop while a usable Orbis results table, candidate tray, snapshot, or report is visibly loaded.
A transient HTTP 403/429 marker is cleared when the requested UI action is subsequently verified as successful (for example, a fallback click opens the expected snapshot). A later failed request can set the marker again.
Restart behavior:

Completed own_ids are skipped.
Existing candidate rows in results.jsonl are deduplicated by own_id, BvD ID, candidate name, national ID, and score.
Decision rows are deduplicated by own_id + source.
If a company crashes before its decision row is written, rerunning will upload it again, avoid appending duplicate candidate rows already captured, finish the decision, and only then checkpoint the company.
A failed or blocked results-page transition stops the run with the reason pagination failure or systemic safety stop. It is not treated as the end of the batch, so untouched companies remain pending and receive no synthetic row-resolution decisions.
Row/candidate safety:

Manual handoff mode validates by visible Orbis Batch Search results DOM: result/company rows, candidate/result elements, tables, or the page-size selector. It rejects obvious Orbis M&A, login, and generic eaccess pages and allows you to retry after typing ready again.
Normal mode navigates directly to the configured plain Orbis Batch Search URL and checks again that it did not drift into Orbis M&A before uploading.
Default upload slices are 100 companies per Batch Search.
At the start of each batch, the scraper checks only for known Orbis release/update notice modals with an in-modal OK button. If found, it screenshots the modal as orbis_run/screenshots/modal_dismissed_*.png, logs it, and clicks only that OK.
Before opening each company, the script closes any previous snapshot/tray.
It matches the Orbis result row by exact company title, then normalized company-name matching, then contains matching. Duplicate title elements in different columns of the same table row are collapsed before ambiguity checks.
Ambiguous loosened matches are logged and skipped for manual review.
Candidate capture is scoped to the single visible tr.poppingRow.
Every candidate is visited sequentially. The scraper opens the snapshot, captures the snapshot text, clicks Go to report, captures structured report fields plus full report text, returns to the candidate list, and then moves to the next candidate.
No candidate is skipped because of the Orbis A/B/C score, and the score is stored only for documentation.
The Orbis candidate NationalId column is captured as national_id for display only.
Candidate country spread of 4+ distinct countries is logged as a manual review warning.
matching_logic.py is used live: candidates are scored against files/master_enriched_reference.xlsx only after all candidates for the current company are collected. Website evidence has highest priority, broader project-reference fields are used when present, and one source: "decision" JSONL record is written per completed company.
matching_pipeline.py is an optional offline audit/rebuild tool. It reads the same JSONL and reference workbook and uses the same matching_logic.py rules to regenerate orbis_run/final_matches.xlsx.
