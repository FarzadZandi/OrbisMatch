import unittest
from unittest.mock import patch

import orbis_candidate_scraper as scraper


class FakeBody:
    def __init__(self, text: str) -> None:
        self.text = text

    def inner_text(self, timeout: int = 0) -> str:
        return self.text


class DetectionPage:
    def __init__(self, body_text: str, visible_captcha: bool = False) -> None:
        self.url = "https://orbis4.bvdinfo.com/results"
        self.body_text = body_text
        self.visible_captcha = visible_captcha

    def locator(self, selector: str) -> FakeBody:
        self.last_selector = selector
        return FakeBody(self.body_text)

    def evaluate(self, script: str) -> dict[str, bool]:
        return {"captchaControl": self.visible_captcha, "usableOrbis": False}


class FakeNextButton:
    def __init__(self, click_error: Exception | None = None) -> None:
        self.first = self
        self.click_error = click_error

    def count(self) -> int:
        return 1

    def is_visible(self, timeout: int = 0) -> bool:
        return True

    def click(self, timeout: int = 0) -> None:
        if self.click_error:
            raise self.click_error


class PaginationPage:
    def __init__(self, button: FakeNextButton) -> None:
        self.button = button

    def locator(self, selector: str) -> FakeNextButton:
        return self.button

    def evaluate(self, script: str):
        return {"firstName": "Alpha", "pageNumber": "1"}

    def wait_for_function(self, *args, **kwargs) -> None:
        return None


class FakeExactLocator:
    def __init__(self, handle: object) -> None:
        self.first = self
        self.handle = handle

    def count(self) -> int:
        return 2

    def evaluate_all(self, script: str):
        # Two exact-title DOM elements (company and selected-match columns)
        # collapse to one unique table row in the browser-side script.
        return [{"index": 0, "city": "Paris", "country": "FR"}]

    def nth(self, index: int):
        if index != 0:
            raise AssertionError(f"unexpected locator index {index}")
        return self

    def element_handle(self, timeout: int = 0):
        return self.handle


class ResolutionPage:
    def __init__(self, locator: FakeExactLocator) -> None:
        self.exact_locator = locator

    def locator(self, selector: str) -> FakeExactLocator:
        return self.exact_locator


class SystemicProblemDetectionTests(unittest.TestCase):
    def test_incidental_captcha_word_on_normal_page_is_not_a_block(self) -> None:
        body = "Normal Orbis results. A documentation note mentions captcha. " + ("row data " * 100)
        self.assertEqual(scraper.detect_systemic_problem(DetectionPage(body)), "")

    def test_short_captcha_interstitial_is_a_block(self) -> None:
        reason = scraper.detect_systemic_problem(DetectionPage("CAPTCHA"))
        self.assertEqual(reason, "short interstitial page contains 'captcha'")

    def test_captcha_like_company_name_is_not_a_block(self) -> None:
        reason = scraper.detect_systemic_problem(DetectionPage("Captchaexample Exampleville XY REG123456"))
        self.assertEqual(reason, "")

    def test_visible_captcha_control_is_a_block(self) -> None:
        reason = scraper.detect_systemic_problem(
            DetectionPage("Normal-looking long page " * 100, visible_captcha=True)
        )
        self.assertEqual(reason, "page contains a visible CAPTCHA control")


class CaptchaSelectorBrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.playwright = scraper.sync_playwright().start()
        cls.browser = cls.playwright.chromium.launch(headless=True)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.browser.close()
        cls.playwright.stop()

    def test_captcha_like_result_row_does_not_match_challenge_controls(self) -> None:
        page = self.browser.new_page()
        try:
            page.set_content(
                '<table><tr class="company-Captchaexample"><td id="Captchaexample">Captchaexample</td>'
                '<td>Exampleville</td><td>XY</td></tr></table>'
            )
            self.assertEqual(scraper.detect_systemic_problem(page), "")
        finally:
            page.close()

    def test_known_recaptcha_control_is_still_detected(self) -> None:
        page = self.browser.new_page()
        try:
            page.set_content('<div class="g-recaptcha" style="width: 20px; height: 20px"></div>')
            self.assertEqual(
                scraper.detect_systemic_problem(page),
                "page contains a visible CAPTCHA control",
            )
        finally:
            page.close()

    def test_captcha_named_candidate_markup_is_ignored_inside_usable_orbis_ui(self) -> None:
        page = self.browser.new_page()
        try:
            page.set_content(
                '<select id="pageSize"><option selected>100</option></select>'
                '<div class="owSnapshot" style="width: 200px; height: 100px">'
                '<form id="CAPTCHAEXAMPLE-candidate"><div>CAPTCHAEXAMPLE LTD</div></form></div>'
            )
            self.assertEqual(scraper.detect_systemic_problem(page), "")
        finally:
            page.close()


class BlockingResponseRecoveryTests(unittest.TestCase):
    def test_verified_ui_success_clears_transient_http_block_marker(self) -> None:
        instance = object.__new__(scraper.OrbisScraper)
        instance.last_block_response_reason = "HTTP 403 from https://example.test/snapshot"
        instance.clear_blocking_response_after_verified_success("snapshot popup opened")
        self.assertEqual(instance.last_block_response_reason, "")


class PaginationSafetyTests(unittest.TestCase):
    def pagination_patches(self):
        return (
            patch.object(scraper, "wait_for_results_page_settled", return_value=True),
            patch.object(scraper, "read_results_page_indicator", return_value=("1", "/ 3", "1 / 3")),
            patch.object(scraper, "get_locator_attribute", side_effect=["", "/next/-default"]),
        )

    def test_systemic_stop_from_after_click_is_not_swallowed(self) -> None:
        page = PaginationPage(FakeNextButton())
        patches = self.pagination_patches()
        with patches[0], patches[1], patches[2]:
            with self.assertRaises(scraper.StopRun) as raised:
                scraper.advance_to_next_results_page(
                    page,
                    {"Alpha"},
                    1000,
                    after_click=lambda: (_ for _ in ()).throw(scraper.StopRun("blocked")),
                )
        self.assertEqual(raised.exception.stop_reason, "systemic safety stop")

    def test_click_failure_stops_as_pagination_failure(self) -> None:
        page = PaginationPage(FakeNextButton(RuntimeError("click failed")))
        patches = self.pagination_patches()
        with patches[0], patches[1], patches[2]:
            with self.assertRaises(scraper.StopRun) as raised:
                scraper.advance_to_next_results_page(page, {"Alpha"}, 1000)
        self.assertEqual(raised.exception.stop_reason, "pagination failure")
        self.assertIn("batch was left incomplete", str(raised.exception))


class CompanyResolutionTests(unittest.TestCase):
    def test_duplicate_title_elements_in_one_table_row_are_not_ambiguous(self) -> None:
        expected_handle = object()
        page = ResolutionPage(FakeExactLocator(expected_handle))
        row = scraper.UploadRow(
            own_id="test-001",
            company="Example Company",
            values=[],
            source_file="example_upload.xlsx",
            source_row=7,
            reference_country="XY",
            reference_city="Exampleville",
        )
        self.assertIs(scraper.resolve_company_row(page, row), expected_handle)


if __name__ == "__main__":
    unittest.main()
