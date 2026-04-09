from __future__ import annotations

import argparse
import contextlib
import json
import logging
import re
import sys
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


SKILL_DIR = Path(__file__).resolve().parent.parent
MMIS_CORE_DIR = Path(r"C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts")
if str(MMIS_CORE_DIR) not in sys.path:
    sys.path.insert(0, str(MMIS_CORE_DIR))

from mmisClient import MMISClient  # noqa: E402


START_CENTER_URL = "https://ap.nmmis.railway.gov.tw/maximo/ui/?event=loadapp&value=startcntr"
DOWNLOADS_DIR = Path(r"C:\Users\NMMIS\Downloads")
LOG_DIR = SKILL_DIR / "logs"
LOG_FILE = LOG_DIR / "mmis_1a_linked_fault_notices.log"
DEBUG_DIR = SKILL_DIR / "debug"
DEFAULT_TIMEOUT_MS = 30000
WORK_ORDER_PATTERN = re.compile(r"^\d{3}-1A-\d+$")
FAULT_NOTICE_PATTERN = re.compile(r"^\d{7,8}-\d+$")

LOGIN_USERNAME_SELECTOR = "#username"
LOGIN_PASSWORD_SELECTOR = "#password"
LOGIN_BUTTON_SELECTOR = "#loginbutton"
LOGIN_ROBOT_SELECTOR = "#iamnotrobot"
APP_LINK_XPATH = (
    "xpath=//a[contains(@href,'ZZ_PMWO1A') or contains(normalize-space(.),'動力車日檢(1A)') "
    "or contains(normalize-space(.),'動力車日檢')]"
)
WORK_ORDER_INPUT_XPATH = (
    "xpath=//input[@id='m6a7dfd2f_tfrow_[C:5]_txt-tb' "
    "or (contains(@id,'tfrow_') and contains(@id,'[C:5]_txt-tb'))]"
)
DEPOT_INPUT_XPATH = (
    "xpath=//input[@id='m6a7dfd2f_tfrow_[C:2]_txt-tb' "
    "or (contains(@id,'tfrow_') and contains(@id,'[C:2]_txt-tb'))]"
)
FIRST_RESULT_CELL_XPATH = (
    "xpath=(//td[contains(@id,'_tdrow_') and contains(@id,'[C:5]-c[R:')])[1]"
)
DETAIL_WORK_ORDER_XPATH = (
    "xpath=//input[@id='mac5943f6-tb' or contains(@id,'ac5943f6-tb')]"
)
FAULT_NOTICE_SPANS_XPATH = "xpath=//td[.//span[@title]]//span[@title]"


@dataclass
class StepMetric:
    name: str
    status: str
    duration_ms: int
    details: dict[str, Any]


class LinkedFaultNoticeQueryError(RuntimeError):
    pass


def sanitize_filename(value: str) -> str:
    return re.sub(r'[^0-9A-Za-z_\-]+', "_", value)


def build_logger() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("mmis_1a_linked_fault_notices")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logger


def write_debug_artifacts(page, name: str) -> None:
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    safe_name = sanitize_filename(name)
    screenshot_path = DEBUG_DIR / f"{safe_name}.png"
    html_path = DEBUG_DIR / f"{safe_name}.html"
    page.screenshot(path=str(screenshot_path), full_page=True)
    html_path.write_text(page.content(), encoding="utf-8")


class LinkedFaultNoticeQuery:
    def __init__(self, work_order_no: str) -> None:
        work_order = work_order_no.strip()
        if not work_order:
            raise LinkedFaultNoticeQueryError("work_order_no 不可為空")
        self.work_order_no = work_order
        self.client = MMISClient()
        self.logger = build_logger()
        self.step_metrics: list[StepMetric] = []
        self.screenshot_path: str | None = None

    def _format_details(self, details: dict[str, Any]) -> str:
        normalized = {k: v for k, v in details.items() if v not in (None, "", [], {})}
        return "" if not normalized else " " + json.dumps(normalized, ensure_ascii=False, sort_keys=True)

    @contextlib.contextmanager
    def timed_step(self, name: str, **details: Any):
        started_at = time.perf_counter()
        self.logger.info("START %s%s", name, self._format_details(details))
        status = "success"
        try:
            yield
        except Exception:
            status = "failed"
            raise
        finally:
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            self.step_metrics.append(StepMetric(name=name, status=status, duration_ms=duration_ms, details=details))
            level = logging.INFO if status == "success" else logging.ERROR
            self.logger.log(level, "END %s status=%s duration_ms=%s", name, status, duration_ms)

    def _result(self, **extra: Any) -> dict[str, Any]:
        return {
            "work_order": self.work_order_no,
            "log_file": str(LOG_FILE),
            "screenshot": self.screenshot_path,
            "step_metrics": [asdict(metric) for metric in self.step_metrics],
            **extra,
        }

    def _cookies_for_browser(self) -> list[dict[str, Any]]:
        cookies: list[dict[str, Any]] = []
        for cookie in self.client.session.cookies:
            if not cookie.domain:
                continue
            cookies.append(
                {
                    "name": cookie.name,
                    "value": cookie.value,
                    "domain": cookie.domain,
                    "path": cookie.path or "/",
                    "httpOnly": False,
                    "secure": True,
                }
            )
        return cookies

    def _ensure_browser_authenticated(self, page) -> bool:
        assert self.client.state is not None
        page.goto(self.client.state.page_url, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")
        if page.locator(APP_LINK_XPATH).count() > 0:
            return True

        page.goto(START_CENTER_URL, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")
        return page.locator(APP_LINK_XPATH).count() > 0

    def _browser_login_fallback(self, page) -> None:
        self.client.require_credentials()
        page.goto(
            "https://ap.nmmis.railway.gov.tw/maximo/webclient/login/login.jsp?welcome=true",
            wait_until="domcontentloaded",
        )
        page.locator(LOGIN_USERNAME_SELECTOR).wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
        page.locator(LOGIN_USERNAME_SELECTOR).fill(self.client.username)
        page.locator(LOGIN_PASSWORD_SELECTOR).fill(self.client.password)
        if page.locator(LOGIN_ROBOT_SELECTOR).count() > 0:
            page.locator(LOGIN_ROBOT_SELECTOR).click()
        page.locator(LOGIN_BUTTON_SELECTOR).click()
        page.wait_for_load_state("networkidle")
        page.locator(APP_LINK_XPATH).first.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)

    def _collect_fault_notices(self, page) -> list[str]:
        notices: list[str] = []
        spans = page.locator(FAULT_NOTICE_SPANS_XPATH)
        count = spans.count()
        for index in range(count):
            span = spans.nth(index)
            title = normalize_value(span.get_attribute("title"))
            text = normalize_value(span.inner_text())
            candidate = title or text
            if not candidate or not FAULT_NOTICE_PATTERN.match(candidate):
                continue
            notices.append(candidate)
        seen: set[str] = set()
        ordered: list[str] = []
        for item in notices:
            if item not in seen:
                seen.add(item)
                ordered.append(item)
        return ordered

    def _save_screenshot(self, page) -> str:
        DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
        target = DOWNLOADS_DIR / f"workOrder_{sanitize_filename(self.work_order_no)}.png"
        page.screenshot(path=str(target), full_page=True)
        self.logger.info("[INFO] screenshot saved")
        self.screenshot_path = str(target)
        return str(target)

    def getLinkedFaultNotices(self) -> dict[str, Any]:
        return self.run()

    def run(self) -> dict[str, Any]:
        login_result = self.client.login()
        reused_session = bool(login_result.get("reused_session"))

        try:
            from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
            from playwright.sync_api import sync_playwright
        except Exception as exc:  # noqa: BLE001
            raise LinkedFaultNoticeQueryError(f"無法載入 Playwright: {exc}") from exc

        try:
            with sync_playwright() as playwright:
                browser = playwright.chromium.launch(headless=True)
                context = browser.new_context(accept_downloads=False)
                page = context.new_page()
                page.set_default_timeout(DEFAULT_TIMEOUT_MS)
                try:
                    cookies = self._cookies_for_browser()
                    if cookies:
                        context.add_cookies(cookies)

                    with self.timed_step("bootstrap_browser_session", cookie_count=len(cookies)):
                        authenticated = self._ensure_browser_authenticated(page)
                        if not authenticated:
                            self._browser_login_fallback(page)
                        self.logger.info("[INFO] login success")

                    with self.timed_step("open_start_center"):
                        page.goto(START_CENTER_URL, wait_until="domcontentloaded")
                        page.wait_for_load_state("networkidle")
                        page.locator(APP_LINK_XPATH).first.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
                        self.logger.info("[INFO] page loaded")

                    with self.timed_step("open_1a_app"):
                        app_link = page.locator(APP_LINK_XPATH).first
                        app_link.click()
                        page.wait_for_load_state("networkidle")
                        try:
                            page.wait_for_url(re.compile(r".*value=zz_pmwo1a.*"), timeout=DEFAULT_TIMEOUT_MS)
                            page.locator(WORK_ORDER_INPUT_XPATH).first.wait_for(
                                state="visible", timeout=DEFAULT_TIMEOUT_MS
                            )
                        except Exception:
                            write_debug_artifacts(page, f"open_1a_app_{self.work_order_no}")
                            raise

                    with self.timed_step("search_work_order", work_order_no=self.work_order_no):
                        self.logger.info("[INFO] searching work order: %s", self.work_order_no)
                        work_order_input = page.locator(WORK_ORDER_INPUT_XPATH).first
                        depot_input = page.locator(DEPOT_INPUT_XPATH).first

                        depot_input.fill("")
                        depot_input.press("Tab")
                        work_order_input.fill("")
                        work_order_input.fill(self.work_order_no)
                        work_order_input.press("Enter")
                        page.wait_for_load_state("networkidle")

                    with self.timed_step("select_first_result"):
                        first_result = page.locator(FIRST_RESULT_CELL_XPATH).first
                        if first_result.count() == 0:
                            self.logger.info("[INFO] result found / not found")
                            return {"error": "找不到工作單", "log_file": str(LOG_FILE)}
                        first_result.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
                        self.logger.info("[INFO] result found / not found")
                        first_result.click()
                        page.wait_for_load_state("networkidle")

                    with self.timed_step("verify_detail_page"):
                        detail_input = page.locator(DETAIL_WORK_ORDER_XPATH).first
                        detail_input.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
                        detail_value = normalize_value(detail_input.input_value())
                        if detail_value != self.work_order_no:
                            raise LinkedFaultNoticeQueryError(
                                f"工單明細頁工作單號不一致: expected={self.work_order_no}, actual={detail_value}"
                            )
                        self.logger.info("[INFO] entering detail page")

                    with self.timed_step("save_screenshot"):
                        self._save_screenshot(page)

                    with self.timed_step("collect_fault_notices"):
                        fault_notices = self._collect_fault_notices(page)
                        self.logger.info("[INFO] fault notices count: %s", len(fault_notices))
                        return self._result(
                            work_order=self.work_order_no,
                            fault_notices=fault_notices,
                            reused_session=reused_session,
                        )
                except PlaywrightTimeoutError as exc:
                    raise LinkedFaultNoticeQueryError(f"頁面操作逾時: {exc}") from exc
                finally:
                    context.close()
                    browser.close()
        except LinkedFaultNoticeQueryError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise LinkedFaultNoticeQueryError(f"查詢失敗: {exc}") from exc


def normalize_value(value: str | None) -> str:
    return (value or "").strip()


def getLinkedFaultNotices(work_order_no: str) -> dict[str, Any]:
    query = LinkedFaultNoticeQuery(work_order_no)
    return query.getLinkedFaultNotices()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Query linked fault notices for MMIS 1A work order")
    parser.add_argument("--work-order-no", required=True, help="Work order number, e.g. 115-1A-23391")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        result = getLinkedFaultNotices(args.work_order_no)
        print(json.dumps(result, ensure_ascii=False))
        return 0 if "error" not in result else 1
    except LinkedFaultNoticeQueryError as exc:
        print(json.dumps({"error": str(exc), "log_file": str(LOG_FILE)}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main())
