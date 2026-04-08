from __future__ import annotations

import argparse
import contextlib
import json
import logging
import re
import sys
import tempfile
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any


SKILL_DIR = Path(__file__).resolve().parent.parent
MMIS_CORE_DIR = Path(r"C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts")
if str(MMIS_CORE_DIR) not in sys.path:
    sys.path.insert(0, str(MMIS_CORE_DIR))

from mmisClient import (  # noqa: E402
    MMISClient,
    MMISClientError,
    generate_level_filename,
    parse_fault_levels,
    resolve_filename_conflict,
)


TARGET_DIR = Path(
    r"C:\Users\NMMIS\OneDrive - Ministry of Transportation and Communications-7280502-Taiwan Railways Administration, MOTC\文件\MMIS桌面"
)
LOG_DIR = SKILL_DIR / "logs"
LOG_FILE = LOG_DIR / "mmis_open_b_level_fault_notices.log"
APP_SELECTOR = "#FavoriteApp_ZZ_FNM"
QUERY_DROPDOWN_SELECTOR = "#toolbar2_tbs_0_tbcb_0_query-img"
DOWNLOAD_IMAGE_SELECTOR = "#m6a7dfd2f-lb4_image"
RESULT_COUNT_SELECTOR = "#m6a7dfd2f-lb3"
DEPOT_INPUT_XPATH = (
    "xpath=//input[contains(@id,'m6a7dfd2f_tfrow_') and contains(@id,'[C:15]_txt-tb')]"
)
LEVEL_INPUT_XPATH = (
    "xpath=//input[contains(@id,'m6a7dfd2f_tfrow_') and contains(@id,'[C:5]_txt-tb')]"
)
LOGIN_USERNAME_SELECTOR = "#username"
LOGIN_PASSWORD_SELECTOR = "#password"
LOGIN_BUTTON_SELECTOR = "#loginbutton"
LOGIN_ROBOT_SELECTOR = "#iamnotrobot"
TARGET_QUERY_NAME = "故障通報未結案清單"
DEFAULT_DEPOT = "新竹機務段"
DEFAULT_LEVEL = "B"
DEFAULT_TIMEOUT_MS = 15000
COUNT_RE = re.compile(r"\b\d+\s*-\s*\d+/(\d+)\b")


@dataclass
class StepMetric:
    name: str
    status: str
    duration_ms: int
    details: dict[str, Any]


class SkillError(RuntimeError):
    pass


def normalize_depot_name(depot_name: str) -> str:
    normalized = depot_name.strip()
    if not normalized:
        raise SkillError("depot 不可為空字串")
    return normalized


def build_logger() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("mmis_open_b_level_fault_notices")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logger


class PlaywrightOpenBLevelFaultNoticeDownloader:
    def __init__(self, *, level: str = DEFAULT_LEVEL, depot_name: str = DEFAULT_DEPOT) -> None:
        self.logger = build_logger()
        self.client = MMISClient()
        self.step_metrics: list[StepMetric] = []
        try:
            self.level_selection = parse_fault_levels(level)
        except MMISClientError as exc:
            raise SkillError(str(exc)) from exc
        self.depot_name = normalize_depot_name(depot_name)

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

    def _session_result(self, **extra: Any) -> dict[str, Any]:
        return {
            "success": True,
            "query_name": TARGET_QUERY_NAME,
            "filters": {
                "配屬段": self.depot_name,
                "等級": self.level_selection.display_level,
                "等級查詢值": self.level_selection.query_level,
            },
            "log_file": str(LOG_FILE),
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

    def _extract_result_count(self, text: str | None) -> int | None:
        if not text:
            return None
        match = COUNT_RE.search(re.sub(r"\s+", " ", text))
        return int(match.group(1)) if match else None

    def _ensure_browser_authenticated(self, page) -> bool:
        assert self.client.state is not None
        page.goto(self.client.state.page_url, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")
        if page.locator(APP_SELECTOR).count() > 0:
            page.locator(APP_SELECTOR).first.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
            return True

        start_center_url = (
            f"https://ap.nmmis.railway.gov.tw/maximo/ui/login"
            f"?uisessionid={self.client.state.ui_session_id}&event=loadapp&value=startcntr"
        )
        page.goto(start_center_url, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")
        if page.locator(APP_SELECTOR).count() > 0:
            page.locator(APP_SELECTOR).first.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
            return True
        return False

    def _browser_login_fallback(self, page) -> None:
        self.client.require_credentials()
        with self.timed_step("browser_login_fallback"):
            page.context.clear_cookies()
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
            page.locator(APP_SELECTOR).first.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)

    def run(self) -> dict[str, Any]:
        login_result = self.client.login()
        TARGET_DIR.mkdir(parents=True, exist_ok=True)
        base_filename = generate_level_filename(
            depot_name=self.depot_name,
            level_display=self.level_selection.display_level,
        )
        target_path = resolve_filename_conflict(TARGET_DIR, base_filename)

        reused_session = bool(login_result.get("reused_session"))
        self.logger.info("input level = %s", self.level_selection.display_level)
        self.logger.info("parsed query = %s", self.level_selection.query_level)
        self.logger.info("filename level = %s", self.level_selection.display_level)

        try:
            from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
            from playwright.sync_api import sync_playwright
        except Exception as exc:  # noqa: BLE001
            raise SkillError(f"無法載入 Playwright: {exc}") from exc

        try:
            with sync_playwright() as playwright:
                with self.timed_step("playwright_query_download", headless=True):
                    browser = playwright.chromium.launch(headless=True)
                    context = browser.new_context(accept_downloads=True)
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
                            self.logger.info("登入成功")

                        with self.timed_step("open_fault_notice_management", selector=APP_SELECTOR):
                            page.locator(APP_SELECTOR).first.click()
                            page.wait_for_url(re.compile(r".*value=zz_fnm.*"), timeout=DEFAULT_TIMEOUT_MS)
                            page.locator(QUERY_DROPDOWN_SELECTOR).wait_for(
                                state="visible", timeout=DEFAULT_TIMEOUT_MS
                            )

                        with self.timed_step("open_query_dropdown", selector=QUERY_DROPDOWN_SELECTOR):
                            page.locator(QUERY_DROPDOWN_SELECTOR).click()
                            page.get_by_text(TARGET_QUERY_NAME, exact=True).first.wait_for(
                                state="visible", timeout=DEFAULT_TIMEOUT_MS
                            )

                        with self.timed_step("select_query", query_name=TARGET_QUERY_NAME):
                            page.get_by_text(TARGET_QUERY_NAME, exact=True).first.click()
                            page.locator(DEPOT_INPUT_XPATH).first.wait_for(
                                state="visible", timeout=DEFAULT_TIMEOUT_MS
                            )
                            page.locator(LEVEL_INPUT_XPATH).first.wait_for(
                                state="visible", timeout=DEFAULT_TIMEOUT_MS
                            )

                        with self.timed_step(
                            "fill_filters",
                            depot=self.depot_name,
                            level=self.level_selection.display_level,
                        ):
                            depot_input = page.locator(DEPOT_INPUT_XPATH).first
                            level_input = page.locator(LEVEL_INPUT_XPATH).first
                            depot_input.fill(self.depot_name)
                            level_input.fill(self.level_selection.query_level)
                            if depot_input.input_value() != self.depot_name:
                                raise SkillError("配屬段欄位未成功寫入")
                            if level_input.input_value() != self.level_selection.query_level:
                                raise SkillError("等級欄位未成功寫入")
                            self.logger.info(
                                "查詢條件已成功套用: 配屬段=%s, 等級=%s",
                                self.depot_name,
                                self.level_selection.query_level,
                            )

                        with self.timed_step("run_query"):
                            level_input = page.locator(LEVEL_INPUT_XPATH).first
                            level_input.press("Enter")
                            page.wait_for_load_state("networkidle")
                            page.locator(DOWNLOAD_IMAGE_SELECTOR).wait_for(
                                state="visible", timeout=DEFAULT_TIMEOUT_MS
                            )

                        result_count = None
                        if page.locator(RESULT_COUNT_SELECTOR).count() > 0:
                            result_count_text = page.locator(RESULT_COUNT_SELECTOR).first.inner_text()
                            result_count = self._extract_result_count(result_count_text)
                            if result_count is not None:
                                self.logger.info("查詢結果筆數: %s", result_count)
                                if result_count == 0:
                                    raise SkillError("查詢成功但結果筆數為 0")
                            else:
                                self.logger.info("查詢結果筆數: 無法直接解析")
                        else:
                            self.logger.info("查詢結果筆數: 找不到結果計數元素")

                        with self.timed_step("download_file", selector=DOWNLOAD_IMAGE_SELECTOR):
                            self.logger.info("base filename = %s", base_filename)
                            if target_path.name != base_filename:
                                self.logger.info("file exists -> resolving conflict")
                            self.logger.info("final filename = %s", target_path.name)
                            temp_path: Path | None = None
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                                temp_path = Path(tmp.name)
                            try:
                                with page.expect_download(timeout=20000) as download_info:
                                    page.locator(DOWNLOAD_IMAGE_SELECTOR).click()
                                download = download_info.value
                                download.save_as(str(temp_path))
                                temp_path.replace(target_path)
                            finally:
                                if temp_path is not None and temp_path.exists():
                                    temp_path.unlink()
                            self.logger.info("下載成功: %s", target_path)

                        return self._session_result(
                            reused_session=reused_session,
                            result_count=result_count,
                            filename=target_path.name,
                            path=str(target_path),
                            execution_mode="playwright",
                        )
                    finally:
                        browser.close()
        except PlaywrightTimeoutError as exc:
            raise SkillError(f"Playwright 執行逾時: {exc}") from exc
        except MMISClientError as exc:
            raise SkillError(str(exc)) from exc


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--level", default=DEFAULT_LEVEL)
    parser.add_argument("--depot", default=DEFAULT_DEPOT)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    downloader = PlaywrightOpenBLevelFaultNoticeDownloader(level=args.level, depot_name=args.depot)
    try:
        result = downloader.run()
    except Exception as exc:  # noqa: BLE001
        downloader.logger.error("MMIS 執行失敗: %s", exc)
        print(
            json.dumps(
                {
                    "success": False,
                    "query_name": TARGET_QUERY_NAME,
                    "filters": {
                        "配屬段": getattr(downloader, "depot_name", args.depot),
                        "等級": getattr(getattr(downloader, "level_selection", None), "display_level", args.level),
                        "等級查詢值": getattr(
                            getattr(downloader, "level_selection", None),
                            "query_level",
                            args.level,
                        ),
                    },
                    "reason": str(exc),
                    "log_file": str(LOG_FILE),
                    "step_metrics": [asdict(metric) for metric in downloader.step_metrics],
                },
                ensure_ascii=False,
            )
        )
        return 1

    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
