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
from typing import Any, Callable, Iterable


SKILL_DIR = Path(__file__).resolve().parent.parent
MMIS_CORE_DIR = Path(r"C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts")
if str(MMIS_CORE_DIR) not in sys.path:
    sys.path.insert(0, str(MMIS_CORE_DIR))

from mmisClient import MMISClient  # noqa: E402


START_CENTER_URL = "https://ap.nmmis.railway.gov.tw/maximo/ui/?event=loadapp&value=startcntr"
LOG_DIR = SKILL_DIR / "logs"
LOG_FILE = LOG_DIR / "mmis_1a_linked_fault_notices.log"
DEFAULT_TIMEOUT_MS = 30000
MAX_RETRIES = 3
WORK_ORDER_PATTERN = re.compile(r"^\d{3}-1A-\d+$")
FAULT_NOTICE_PATTERN = re.compile(r"^\d{7,8}-\d+$")

LOGIN_USERNAME_SELECTOR = "#username"
LOGIN_PASSWORD_SELECTOR = "#password"
LOGIN_BUTTON_SELECTOR = "#loginbutton"
LOGIN_ROBOT_SELECTOR = "#iamnotrobot"
APP_LINK_SELECTOR = (
    "xpath=(//a[contains(@href,'ZZ_PMWO1A') "
    "or contains(@href,'zz_pmwo1a') "
    "or contains(normalize-space(.),'動力車日檢(1A)') "
    "or contains(normalize-space(.),'動力車日檢')])[1]"
)
WORK_ORDER_INPUT_SELECTOR = (
    "xpath=(//input[@id='m6a7dfd2f_tfrow_[C:5]_txt-tb' "
    "or (contains(@id,'tfrow_') and contains(@id,'[C:5]_txt-tb'))])[1]"
)
DEPOT_INPUT_SELECTOR = (
    "xpath=(//input[@id='m6a7dfd2f_tfrow_[C:2]_txt-tb' "
    "or (contains(@id,'tfrow_') and contains(@id,'[C:2]_txt-tb'))])[1]"
)
RESULT_ROW_SELECTOR = "xpath=//td[contains(@id,'_tdrow_') and contains(@id,'[C:5]-c[R:')]"
DETAIL_WORK_ORDER_SELECTOR = (
    "xpath=(//input[@id='mac5943f6-tb' or contains(@id,'ac5943f6-tb')])[1]"
)
FAULT_NOTICE_SPANS_SELECTOR = "xpath=//td[.//span[@title]]//span[@title]"


@dataclass
class StepMetric:
    name: str
    status: str
    duration_ms: int
    details: dict[str, Any]


class LinkedFaultNoticeQueryError(RuntimeError):
    pass


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


def normalize_value(value: str | None) -> str:
    return (value or "").strip()


class LinkedFaultNoticeQuery:
    def __init__(self) -> None:
        self.client = MMISClient()
        self.logger = build_logger()
        self.step_metrics: list[StepMetric] = []
        self.page = None
        self.context = None
        self.browser = None
        self.playwright = None
        self._browser_authenticated = False
        self._on_1a_page = False
        self._last_login_reused = False

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

    def _result(
        self,
        work_order_no: str,
        ok: bool,
        fault_notices: list[str],
        error: str | None = None,
        elapsed_ms: int | None = None,
    ) -> dict[str, Any]:
        result: dict[str, Any] = {
            "ok": ok,
            "work_order": work_order_no,
            "fault_notices": fault_notices,
            "count": len(fault_notices),
            "log_file": str(LOG_FILE),
            "step_metrics": [asdict(metric) for metric in self.step_metrics],
        }
        if error:
            result["error"] = error
        if elapsed_ms is not None:
            result["elapsed_ms"] = elapsed_ms
        return result

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

    def _load_playwright(self) -> None:
        if self.playwright is not None:
            return
        try:
            from playwright.sync_api import sync_playwright  # type: ignore
        except Exception as exc:  # noqa: BLE001
            raise LinkedFaultNoticeQueryError(f"無法載入 Playwright: {exc}") from exc
        self.playwright = sync_playwright().start()

    def _ensure_browser(self) -> None:
        if self.page is not None:
            return
        self._load_playwright()
        assert self.playwright is not None
        self.browser = self.playwright.chromium.launch(headless=True)
        self.context = self.browser.new_context(accept_downloads=False)
        self.page = self.context.new_page()
        self.page.set_default_timeout(DEFAULT_TIMEOUT_MS)

    def close(self) -> None:
        if self.context is not None:
            self.context.close()
            self.context = None
        if self.browser is not None:
            self.browser.close()
            self.browser = None
        if self.playwright is not None:
            self.playwright.stop()
            self.playwright = None
        self.page = None
        self._browser_authenticated = False
        self._on_1a_page = False

    def __enter__(self) -> "LinkedFaultNoticeQuery":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def _with_retry(self, name: str, action: Callable[[], Any], retries: int = MAX_RETRIES) -> Any:
        last_error: Exception | None = None
        for attempt in range(1, retries + 1):
            try:
                if attempt > 1:
                    self.logger.info("[INFO] retry attempt: %s step=%s", attempt, name)
                return action()
            except Exception as exc:  # noqa: BLE001
                last_error = exc
                if attempt >= retries:
                    break
                if name in {"ensureLoggedIn", "open1AWorkOrderPage"}:
                    self._browser_authenticated = False
                    self._on_1a_page = False
        assert last_error is not None
        raise last_error

    def ensureLoggedIn(self) -> bool:
        def action() -> bool:
            with self.timed_step("ensureLoggedIn"):
                login_result = self.client.login()
                self._last_login_reused = bool(login_result.get("reused_session"))
                self._ensure_browser()
                assert self.page is not None
                assert self.context is not None

                cookies = self._cookies_for_browser()
                if cookies:
                    with contextlib.suppress(Exception):
                        self.context.clear_cookies()
                    self.context.add_cookies(cookies)

                authenticated = self._ensure_browser_authenticated()
                if not authenticated:
                    self._browser_login_fallback()
                    self._browser_authenticated = True
                    self.logger.info("[INFO] login performed")
                else:
                    self._browser_authenticated = True
                    self.logger.info("[INFO] login reused")
                return self._last_login_reused

        return self._with_retry("ensureLoggedIn", action)

    def _ensure_browser_authenticated(self) -> bool:
        assert self.page is not None
        assert self.client.state is not None

        candidate_urls = [self.client.state.page_url, START_CENTER_URL]
        for url in candidate_urls:
            self.page.goto(url, wait_until="domcontentloaded")
            self.page.wait_for_load_state("networkidle")
            if self.page.locator(APP_LINK_SELECTOR).count() > 0:
                return True
        return False

    def _browser_login_fallback(self) -> None:
        assert self.page is not None
        self.client.require_credentials()
        self.page.goto(
            "https://ap.nmmis.railway.gov.tw/maximo/webclient/login/login.jsp?welcome=true",
            wait_until="domcontentloaded",
        )
        self.page.locator(LOGIN_USERNAME_SELECTOR).wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
        self.page.locator(LOGIN_USERNAME_SELECTOR).fill(self.client.username)
        self.page.locator(LOGIN_PASSWORD_SELECTOR).fill(self.client.password)
        if self.page.locator(LOGIN_ROBOT_SELECTOR).count() > 0:
            self.page.locator(LOGIN_ROBOT_SELECTOR).click()
        self.page.locator(LOGIN_BUTTON_SELECTOR).click()
        self.page.wait_for_load_state("networkidle")
        self.page.locator(APP_LINK_SELECTOR).wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)

    def open1AWorkOrderPage(self) -> None:
        def action() -> None:
            with self.timed_step("open1AWorkOrderPage"):
                if not self._browser_authenticated:
                    self.ensureLoggedIn()
                assert self.page is not None

                if self._on_1a_page and self.page.locator(WORK_ORDER_INPUT_SELECTOR).count() > 0:
                    self.logger.info("[INFO] entered 1A page")
                    return

                self.page.goto(START_CENTER_URL, wait_until="domcontentloaded")
                self.page.wait_for_load_state("networkidle")
                app_link = self.page.locator(APP_LINK_SELECTOR)
                app_link.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
                app_link.click()
                self.page.wait_for_url(re.compile(r".*value=zz_pmwo1a.*"), timeout=DEFAULT_TIMEOUT_MS)
                self.page.locator(WORK_ORDER_INPUT_SELECTOR).wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
                self.page.locator(DEPOT_INPUT_SELECTOR).wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
                self._on_1a_page = True
                self.logger.info("[INFO] entered 1A page")

        self._with_retry("open1AWorkOrderPage", action)

    def searchWorkOrder(self, work_order_no: str) -> bool:
        def action() -> bool:
            with self.timed_step("searchWorkOrder", work_order_no=work_order_no):
                self.open1AWorkOrderPage()
                assert self.page is not None
                self.logger.info("[INFO] searching work order: %s", work_order_no)

                depot_input = self.page.locator(DEPOT_INPUT_SELECTOR)
                work_order_input = self.page.locator(WORK_ORDER_INPUT_SELECTOR)
                depot_input.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
                work_order_input.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)

                depot_input.fill("")
                depot_input.press("Tab")
                work_order_input.fill("")
                work_order_input.fill(work_order_no)
                work_order_input.press("Enter")
                self.page.wait_for_load_state("networkidle")

                if self._has_no_results():
                    self.logger.info("[INFO] work order found / not found")
                    return False

                result_rows = self.page.locator(RESULT_ROW_SELECTOR)
                result_rows.first.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
                self.logger.info("[INFO] work order found / not found")
                return result_rows.count() > 0

        return bool(self._with_retry("searchWorkOrder", action))

    def _has_no_results(self) -> bool:
        assert self.page is not None
        no_result_patterns = [
            "查無資料",
            "沒有資料",
            "No records to display",
        ]
        page_text = self.page.locator("body").inner_text(timeout=DEFAULT_TIMEOUT_MS)
        return any(pattern in page_text for pattern in no_result_patterns) and self.page.locator(RESULT_ROW_SELECTOR).count() == 0

    def openFirstResult(self, work_order_no: str) -> bool:
        def action() -> bool:
            with self.timed_step("openFirstResult", work_order_no=work_order_no):
                assert self.page is not None
                result_rows = self.page.locator(RESULT_ROW_SELECTOR)
                if result_rows.count() == 0:
                    self.logger.info("[INFO] work order found / not found")
                    return False
                first_result = result_rows.first
                first_result.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
                first_result.click()
                self.page.locator(DETAIL_WORK_ORDER_SELECTOR).wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
                detail_value = normalize_value(self.page.locator(DETAIL_WORK_ORDER_SELECTOR).input_value())
                if detail_value != work_order_no:
                    raise LinkedFaultNoticeQueryError(
                        f"工單明細頁工作單號不一致: expected={work_order_no}, actual={detail_value}"
                    )
                self.logger.info("[INFO] entered detail page")
                return True

        return bool(self._with_retry("openFirstResult", action))

    def extractLinkedFaultNotices(self) -> list[str]:
        def action() -> list[str]:
            with self.timed_step("extractLinkedFaultNotices"):
                assert self.page is not None
                spans = self.page.locator(FAULT_NOTICE_SPANS_SELECTOR)
                notices: list[str] = []
                for index in range(spans.count()):
                    span = spans.nth(index)
                    title = normalize_value(span.get_attribute("title"))
                    text = normalize_value(span.inner_text())
                    candidate = title or text
                    if candidate and FAULT_NOTICE_PATTERN.match(candidate):
                        notices.append(candidate)

                seen: set[str] = set()
                ordered: list[str] = []
                for item in notices:
                    if item not in seen:
                        seen.add(item)
                        ordered.append(item)
                self.logger.info("[INFO] fault notices count: %s", len(ordered))
                return ordered

        return list(self._with_retry("extractLinkedFaultNotices", action))

    def getLinkedFaultNotices(self, work_order_no: str) -> dict[str, Any]:
        started_at = time.perf_counter()
        self.step_metrics = []
        normalized_work_order = normalize_value(work_order_no)
        if not normalized_work_order:
            return self._result("", False, [], error="work_order_no 不可為空", elapsed_ms=0)
        if not WORK_ORDER_PATTERN.match(normalized_work_order):
            return self._result(normalized_work_order, False, [], error="工單號格式不正確", elapsed_ms=0)

        try:
            self.ensureLoggedIn()
            found = self.searchWorkOrder(normalized_work_order)
            if not found:
                elapsed_ms = int((time.perf_counter() - started_at) * 1000)
                self.logger.info("[INFO] total elapsed time: %s ms", elapsed_ms)
                return self._result(
                    normalized_work_order,
                    False,
                    [],
                    error="找不到工作單",
                    elapsed_ms=elapsed_ms,
                )

            opened = self.openFirstResult(normalized_work_order)
            if not opened:
                elapsed_ms = int((time.perf_counter() - started_at) * 1000)
                self.logger.info("[INFO] total elapsed time: %s ms", elapsed_ms)
                return self._result(
                    normalized_work_order,
                    False,
                    [],
                    error="找不到工作單",
                    elapsed_ms=elapsed_ms,
                )

            fault_notices = self.extractLinkedFaultNotices()
            elapsed_ms = int((time.perf_counter() - started_at) * 1000)
            self.logger.info("[INFO] total elapsed time: %s ms", elapsed_ms)
            return self._result(normalized_work_order, True, fault_notices, elapsed_ms=elapsed_ms)
        except LinkedFaultNoticeQueryError as exc:
            elapsed_ms = int((time.perf_counter() - started_at) * 1000)
            self.logger.info("[INFO] total elapsed time: %s ms", elapsed_ms)
            return self._result(normalized_work_order, False, [], error=str(exc), elapsed_ms=elapsed_ms)
        except Exception as exc:  # noqa: BLE001
            elapsed_ms = int((time.perf_counter() - started_at) * 1000)
            self.logger.info("[INFO] total elapsed time: %s ms", elapsed_ms)
            return self._result(normalized_work_order, False, [], error=f"查詢失敗: {exc}", elapsed_ms=elapsed_ms)

    def getLinkedFaultNoticesBatch(self, work_order_list: Iterable[str]) -> list[dict[str, Any]]:
        results: list[dict[str, Any]] = []
        for work_order_no in work_order_list:
            results.append(self.getLinkedFaultNotices(work_order_no))
        return results


def getLinkedFaultNotices(work_order_no: str) -> dict[str, Any]:
    with LinkedFaultNoticeQuery() as query:
        return query.getLinkedFaultNotices(work_order_no)


def getLinkedFaultNoticesBatch(work_order_list: Iterable[str]) -> list[dict[str, Any]]:
    with LinkedFaultNoticeQuery() as query:
        return query.getLinkedFaultNoticesBatch(work_order_list)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Query linked fault notices for MMIS 1A work order")
    parser.add_argument("--work-order-no", required=True, help="Work order number, e.g. 115-1A-23391")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    result = getLinkedFaultNotices(args.work_order_no)
    print(json.dumps(result, ensure_ascii=False))
    if result.get("ok") is False and result.get("error") not in {"找不到工作單", None}:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
