from __future__ import annotations

import argparse
import contextlib
import html
import json
import logging
import os
import re
import subprocess
import sys
import time
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import requests
import urllib3
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://ap.nmmis.railway.gov.tw"
LOGIN_PAGE_URL = f"{BASE_URL}/maximo/webclient/login/login.jsp?welcome=true"
LOGIN_POST_URL = f"{BASE_URL}/maximo/webclient/login/mxlogin.jsp?welcome=true"
UI_EVENT_URL = f"{BASE_URL}/maximo/ui/maximo.jsp"
TARGET_DIR = Path(
    r"C:\Users\NMMIS\OneDrive - Ministry of Transportation and Communications-7280502-Taiwan Railways Administration, MOTC\文件\MMIS桌面"
)
SCRIPT_DIR = Path(__file__).resolve().parent
LOG_DIR = SCRIPT_DIR.parent / "logs"
CACHE_DIR = SCRIPT_DIR.parent / "cache"
LOG_FILE = LOG_DIR / "mmis.log"
SESSION_CACHE_FILE = CACHE_DIR / "session.json"
EXCEL_FORMATTER = (
    Path(r"C:\Users\NMMIS\.codex\skills\mmis-excel-formatting\scripts\format_mmis_excel.py")
)
OPEN_B_LEVEL_PLAYWRIGHT_SCRIPT = Path(
    r"C:\Users\NMMIS\.codex\skills\mmis-query-open-b-level-fault-notices\scripts\playwright_open_b_level_fault_notice_download.py"
)
REQUEST_TIMEOUT_SECONDS = 20
REQUEST_RETRIES = 3
REQUEST_RETRY_DELAY_SECONDS = 1.0
SESSION_MAX_AGE_MINUTES = 50
PLAYWRIGHT_FALLBACK_ENABLED = (
    os.environ.get("MMIS_ENABLE_PLAYWRIGHT_FALLBACK", "true").lower() == "true"
)
DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/147.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}
PAGE_SEQ_RE = re.compile(r'var\s+PAGESEQNUM\s*=\s*"(?P<page_seq>\d+)"')
UI_SESSION_RE = re.compile(
    r'var\s+UISESSIONID\s*=\s*decodeURIComponent\("(?P<ui_session_id>\d+)"\)'
)
CSRF_TOKEN_RE = re.compile(r'var\s+CSRFTOKEN\s*=\s*"(?P<csrf_token>[a-z0-9]+)"')
APP_ID_RE = re.compile(r'var\s+APPID\s*=\s*"(?P<app_id>[^"]+)"')
REDIRECT_SESSION_RE = re.compile(r"uisessionid=(?P<ui_session_id>\d+)")
EVENT_REDIRECT_RE = re.compile(r"<redirect><!\[CDATA\[(?P<url>[^\]]+)\]\]></redirect>")
DOWNLOAD_URL_RE = re.compile(
    r'openEncodedURL\((?:"|&quot;)(?P<url>https://[^"]+?_tbldnld=[^"]+)'
)
COUNT_RE = re.compile(r"\b\d+\s*-\s*\d+/(\d+)\b")
DEPOT_ABBREVIATIONS = {
    "七堵機務段": "七機",
    "臺北機務段": "北機",
    "新竹機務段": "本段",
    "彰化機務段": "彰機",
    "嘉義機務段": "嘉機",
    "高雄機務段": "高機",
    "宜蘭機務段": "宜機",
    "花蓮機務段": "花機",
    "臺東機務段": "東機",
}

SAVED_QUERIES: dict[str, dict[str, str]] = {
    "本段未處理通報(車輛配屬段)": {
        "menu_value": "本段未處理通報_query",
        "focus_id": "menu0_本段未處理通報_query_a",
    },
    "故障通報未結案清單": {
        "menu_value": "故障通報未結案清單_query",
        "focus_id": "menu0_故障通報未結案清單_query_a",
    }
}


@dataclass
class PageState:
    ui_session_id: str
    page_seq: int
    csrf_token: str
    app_id: str
    page_url: str


@dataclass
class StepMetric:
    name: str
    status: str
    duration_ms: int
    details: dict[str, Any]


@dataclass(frozen=True)
class FaultLevelSelection:
    display_level: str
    query_level: str


class MMISClientError(RuntimeError):
    pass


def mmdd_today() -> str:
    return datetime.now().strftime("%m%d")


def build_target_filename() -> str:
    return f"本段未處理故障通報{mmdd_today()}.xlsx"


def depot_short_name(depot_name: str) -> str:
    return DEPOT_ABBREVIATIONS.get(depot_name, depot_name)


def generate_level_filename(*, depot_name: str, level_display: str) -> str:
    return f"{depot_short_name(depot_name)}{level_display}級故障通報管理{mmdd_today()}.xlsx"


def resolve_filename_conflict(target_dir: Path, filename: str) -> Path:
    candidate = target_dir / filename
    if not candidate.exists():
        return candidate

    stem = candidate.stem
    suffix = candidate.suffix
    counter = 2
    while True:
        numbered = target_dir / f"{stem}({counter}){suffix}"
        if not numbered.exists():
            return numbered
        counter += 1


def parse_fault_levels(level: str) -> FaultLevelSelection:
    raw = (level or "").strip().upper()
    compact = raw.replace(",", "").replace(" ", "")
    if not compact:
        compact = "B"

    seen: set[str] = set()
    ordered_levels: list[str] = []
    for ch in compact:
        if ch not in {"A", "B", "C"}:
            raise MMISClientError("level 僅允許 A、B、C 或其組合")
        if ch in seen:
            raise MMISClientError("level 不可重複指定相同等級")
        seen.add(ch)
        ordered_levels.append(ch)

    display_level = "".join(ordered_levels)
    query_level = ",".join(ordered_levels)
    return FaultLevelSelection(display_level=display_level, query_level=query_level)


def normalize_depot_name(depot_name: str) -> str:
    normalized = depot_name.strip()
    if not normalized:
        raise MMISClientError("depot 不可為空字串")
    return normalized


def build_logger() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("mmis")
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


def parse_page_state(text: str, page_url: str) -> PageState:
    page_seq = PAGE_SEQ_RE.search(text)
    ui_session_id = UI_SESSION_RE.search(text)
    csrf_token = CSRF_TOKEN_RE.search(text)
    app_id = APP_ID_RE.search(text)
    if not all([page_seq, ui_session_id, csrf_token, app_id]):
        raise MMISClientError("無法從頁面內容解析 MMIS session state")
    return PageState(
        ui_session_id=ui_session_id.group("ui_session_id"),
        page_seq=int(page_seq.group("page_seq")),
        csrf_token=csrf_token.group("csrf_token"),
        app_id=app_id.group("app_id"),
        page_url=page_url,
    )


def extract_download_url(text: str) -> str | None:
    decoded = html.unescape(text)
    match = DOWNLOAD_URL_RE.search(decoded)
    return match.group("url") if match else None


def extract_event_redirect_url(text: str) -> str | None:
    decoded = html.unescape(text)
    match = EVENT_REDIRECT_RE.search(decoded)
    return match.group("url") if match else None


def extract_result_count(text: str) -> int | None:
    decoded = re.sub(r"\s+", " ", html.unescape(text))
    match = COUNT_RE.search(decoded)
    return int(match.group(1)) if match else None


class MMISClient:
    def __init__(
        self,
        username: str | None = None,
        password: str | None = None,
        verify_ssl: bool | None = None,
        session_cache_file: Path = SESSION_CACHE_FILE,
    ) -> None:
        self.username = username or os.environ.get("MMIS_USERNAME")
        self.password = password or os.environ.get("MMIS_PASSWORD")
        self.verify_ssl = (
            verify_ssl
            if verify_ssl is not None
            else os.environ.get("MMIS_VERIFY_SSL", "").lower() == "true"
        )
        self.session_cache_file = session_cache_file
        self.logger = build_logger()
        self.session = requests.Session()
        self.session.verify = self.verify_ssl
        self.session.headers.update(DEFAULT_HEADERS)
        self.state: PageState | None = None
        self.step_metrics: list[StepMetric] = []
        self.request_retry_count = 0

        retry = Retry(
            total=3,
            connect=3,
            read=3,
            backoff_factor=0.4,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET", "POST"],
        )
        adapter = HTTPAdapter(max_retries=retry)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)

    def require_credentials(self) -> None:
        if not self.username or not self.password:
            raise MMISClientError("環境變數 MMIS_USERNAME 或 MMIS_PASSWORD 未設定")

    def _format_details(self, details: dict[str, Any]) -> str:
        normalized = {k: v for k, v in details.items() if v not in (None, "", [], {})}
        if not normalized:
            return ""
        return " " + json.dumps(normalized, ensure_ascii=False, sort_keys=True)

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
            self.step_metrics.append(
                StepMetric(
                    name=name,
                    status=status,
                    duration_ms=duration_ms,
                    details=details,
                )
            )
            level = logging.INFO if status == "success" else logging.ERROR
            self.logger.log(level, "END %s status=%s duration_ms=%s", name, status, duration_ms)

    def _load_cached_state(self) -> PageState | None:
        if not self.session_cache_file.exists():
            return None
        try:
            payload = json.loads(self.session_cache_file.read_text(encoding="utf-8"))
            cached_at = datetime.fromisoformat(payload["cached_at"])
            if datetime.now() - cached_at > timedelta(minutes=SESSION_MAX_AGE_MINUTES):
                return None
            for cookie in payload.get("cookies", []):
                self.session.cookies.set(
                    cookie["name"],
                    cookie["value"],
                    domain=cookie.get("domain"),
                    path=cookie.get("path", "/"),
                )
            return PageState(**payload["state"])
        except Exception as exc:  # noqa: BLE001
            self.logger.warning("忽略失敗的 session cache: %s", exc)
            return None

    def _save_cached_state(self, state: PageState) -> None:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        payload = {
            "cached_at": datetime.now().isoformat(),
            "state": asdict(state),
            "cookies": [
                {
                    "name": cookie.name,
                    "value": cookie.value,
                    "domain": cookie.domain,
                    "path": cookie.path,
                }
                for cookie in self.session.cookies
            ],
        }
        self.session_cache_file.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _request(
        self,
        method: str,
        url: str,
        *,
        expected_status: int | tuple[int, ...] = (200,),
        **kwargs: Any,
    ) -> requests.Response:
        expected = (expected_status,) if isinstance(expected_status, int) else expected_status
        timeout = kwargs.pop("timeout", REQUEST_TIMEOUT_SECONDS)
        last_error: Exception | None = None

        for attempt in range(1, REQUEST_RETRIES + 1):
            started_at = time.perf_counter()
            try:
                response = self.session.request(method, url, timeout=timeout, **kwargs)
                duration_ms = int((time.perf_counter() - started_at) * 1000)
                if response.status_code in expected:
                    self.logger.info(
                        "REQUEST %s %s status=%s duration_ms=%s attempt=%s",
                        method,
                        url,
                        response.status_code,
                        duration_ms,
                        attempt,
                    )
                    return response
                last_error = MMISClientError(
                    f"HTTP {response.status_code} for {method} {url}: {response.text[:300]}"
                )
                retryable = response.status_code >= 500 or response.status_code in {408, 409, 425, 429}
                self.logger.warning(
                    "REQUEST %s %s unexpected_status=%s duration_ms=%s attempt=%s",
                    method,
                    url,
                    response.status_code,
                    duration_ms,
                    attempt,
                )
            except requests.RequestException as exc:
                duration_ms = int((time.perf_counter() - started_at) * 1000)
                last_error = exc
                retryable = True
                self.logger.warning(
                    "REQUEST %s %s error=%s duration_ms=%s attempt=%s",
                    method,
                    url,
                    exc,
                    duration_ms,
                    attempt,
                )

            if attempt < REQUEST_RETRIES and retryable:
                self.request_retry_count += 1
                sleep_seconds = REQUEST_RETRY_DELAY_SECONDS * attempt
                self.logger.warning(
                    "RETRY %s %s next_attempt=%s sleep_s=%.1f",
                    method,
                    url,
                    attempt + 1,
                    sleep_seconds,
                )
                time.sleep(sleep_seconds)
                continue
            break

        raise MMISClientError(str(last_error))

    def _session_result(self, reused_session: bool) -> dict[str, Any]:
        return {
            "logged_in": True,
            "reused_session": reused_session,
            "uisessionid": self.state.ui_session_id if self.state else None,
            "app_id": self.state.app_id if self.state else None,
            "page_url": self.state.page_url if self.state else None,
            "log_file": str(LOG_FILE),
            "retry_count": self.request_retry_count,
            "step_metrics": [asdict(metric) for metric in self.step_metrics],
        }

    def login(self, force: bool = False) -> dict[str, Any]:
        self.require_credentials()
        result: dict[str, Any] | None = None
        with self.timed_step("login", force=force):
            if force:
                self.session.cookies.clear()
                self.state = None

            if not force:
                cached = self._load_cached_state()
                if cached and self._validate_state(cached):
                    self.logger.info("沿用既有 MMIS session: uisessionid=%s", self.state.ui_session_id)
                    result = self._session_result(reused_session=True)
            if result is None:
                login_page = self._request("GET", LOGIN_PAGE_URL).text
                soup = BeautifulSoup(login_page, "html.parser")
                form = soup.find("form", {"id": "loginform"})
                if form is None:
                    try:
                        self.state = parse_page_state(login_page, LOGIN_PAGE_URL)
                    except MMISClientError as exc:
                        raise MMISClientError("找不到 loginform，無法登入 MMIS") from exc
                    self._save_cached_state(self.state)
                    self.logger.info("MMIS 已在登入狀態，直接沿用現有頁面 session")
                    result = self._session_result(reused_session=True)
                else:
                    form_data: dict[str, str] = {}
                    for element in form.find_all("input"):
                        name = element.get("name")
                        if name:
                            form_data[name] = element.get("value", "")
                    form_data["username"] = self.username
                    form_data["password"] = self.password

                    response = self._request(
                        "POST",
                        urljoin(LOGIN_PAGE_URL, form.get("action", LOGIN_POST_URL)),
                        data=form_data,
                        headers={"Referer": LOGIN_PAGE_URL},
                        allow_redirects=False,
                        expected_status=(302, 303),
                    )
                    redirect_url = response.headers.get("Location") or response.headers.get("location")
                    if not redirect_url:
                        raise MMISClientError("登入後未取得 MMIS redirect URL")

                    match = REDIRECT_SESSION_RE.search(redirect_url)
                    if not match:
                        raise MMISClientError("登入成功但無法解析 uisessionid")

                    ui_session_id = match.group("ui_session_id")
                    page_url = (
                        f"{BASE_URL}/maximo/ui/login?uisessionid={ui_session_id}"
                        "&event=loadapp&value=startcntr"
                    )
                    page_response = self._request("GET", page_url)
                    if "使用者登錄" in page_response.text and "loginform" in page_response.text:
                        raise MMISClientError("登入失敗，MMIS 將請求導回登入頁")

                    self.state = parse_page_state(page_response.text, page_response.url)
                    self._save_cached_state(self.state)
                    self.logger.info("登入成功: uisessionid=%s", self.state.ui_session_id)
                    result = self._session_result(reused_session=False)
        assert result is not None
        result["step_metrics"] = [asdict(metric) for metric in self.step_metrics]
        return result

    def _validate_state(self, state: PageState) -> bool:
        try:
            page_url = (
                f"{BASE_URL}/maximo/ui/login?uisessionid={state.ui_session_id}"
                "&event=loadapp&value=startcntr"
            )
            response = self._request("GET", page_url)
            if "使用者登錄" in response.text and "loginform" in response.text:
                return False
            self.state = parse_page_state(response.text, response.url)
            self._save_cached_state(self.state)
            return True
        except Exception:  # noqa: BLE001
            return False

    def _refresh_state(self, page_url: str | None = None) -> PageState:
        if self.state is None:
            raise MMISClientError("MMIS state 尚未建立，無法刷新頁面狀態")
        response = self._request("GET", page_url or self.state.page_url)
        self.state = parse_page_state(response.text, response.url)
        self._save_cached_state(self.state)
        return self.state

    def _change_app_from_start_center(self, app_value: str) -> PageState:
        if self.state is None:
            raise MMISClientError("MMIS state 尚未建立，無法切換應用程式")
        if self.state.app_id != "startcntr":
            raise MMISClientError("目前不在啟動中心，無法使用 changeapp 切換應用程式")

        # Start center must be refreshed once so PAGESEQNUM/CSRFTOKEN align with
        # the same state used by the browser before posting changeapp.
        state = self._refresh_state(self.state.page_url)
        app_value_upper = app_value.upper()
        response_text = self._post_event(
            currentfocus=f"FavoriteApp_{app_value_upper}",
            event_type="changeapp",
            target_id="startcntr",
            value=app_value_upper,
            xhr_seq=1,
            referer=state.page_url,
        )
        redirect_url = extract_event_redirect_url(response_text)
        if not redirect_url:
            raise MMISClientError(f"changeapp 成功送出但未取得 {app_value} redirect URL")
        return self._refresh_state(redirect_url)

    def load_app(self, app_value: str, *, force_refresh: bool = False) -> PageState:
        with self.timed_step("load_app", app_value=app_value):
            if self.state is None:
                self.login()
            assert self.state is not None

            if self.state.app_id == app_value:
                return self._refresh_state() if force_refresh else self.state

            if app_value != "startcntr" and self.state.app_id == "startcntr":
                return self._change_app_from_start_center(app_value)

            if app_value == "startcntr":
                page_url = (
                    f"{BASE_URL}/maximo/ui/login?uisessionid={self.state.ui_session_id}"
                    "&event=loadapp&value=startcntr"
                )
            else:
                page_url = (
                    f"{BASE_URL}/maximo/ui/?event=loadapp&value={app_value}"
                    f"&uisessionid={self.state.ui_session_id}"
                )
            response = self._request("GET", page_url)
            self.state = parse_page_state(response.text, response.url)
            self._save_cached_state(self.state)
            return self.state

    def _post_event(
        self,
        *,
        currentfocus: str,
        event_type: str | None = None,
        target_id: str | None = None,
        value: str = "",
        xhr_seq: int,
        referer: str | None = None,
        events: list[dict[str, Any]] | None = None,
    ) -> str:
        if self.state is None:
            raise MMISClientError("MMIS state 尚未建立，無法送出事件")

        if events is None:
            if event_type is None or target_id is None:
                raise MMISClientError("未提供 events 時，event_type 與 target_id 為必填")
            events = [
                {
                    "type": event_type,
                    "targetId": target_id,
                    "value": value,
                    "requestType": "SYNC",
                    "csrftokenholder": self.state.csrf_token,
                }
            ]

        with self.timed_step(
            "post_event",
            event_type=event_type or "custom",
            target_id=target_id or "multi",
            xhr_seq=xhr_seq,
        ):
            referer_url = referer or self.state.page_url
            payload = {
                "uisessionid": self.state.ui_session_id,
                "csrftoken": self.state.csrf_token,
                "currentfocus": currentfocus,
                "scrollleftpos": "0",
                "localStorage": "true",
                "scrolltoppos": "0",
                "requesttype": "SYNC",
                "responsetype": "text/xml",
                "events": json.dumps(events, ensure_ascii=False, separators=(",", ":")),
            }
            headers = {
                "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
                "X-Requested-With": "XMLHttpRequest",
                "Origin": BASE_URL,
                "Referer": referer_url,
                "qmtry": "1",
                "pageseqnum": str(self.state.page_seq),
                "xhrseqnum": str(xhr_seq),
            }
            response = self._request("POST", UI_EVENT_URL, data=payload, headers=headers)
            if "exit.jsp?sharedSession=1" in response.text:
                raise MMISClientError("MMIS session 已失效，伺服器要求重新登入")
            try:
                self.state = parse_page_state(response.text, self.state.page_url)
                self._save_cached_state(self.state)
            except MMISClientError:
                self.logger.warning("事件回應未帶完整 page state，沿用上一個 state")
            return response.text

    def _open_saved_query_menu(self, *, state: PageState) -> str:
        return self._post_event(
            currentfocus="toolbar2_tbs_0_tbcb_0_query-tb",
            event_type="click",
            target_id="toolbar2_tbs_0_tbcb_0_query-img",
            value="",
            xhr_seq=1,
            referer=state.page_url,
        )

    def _apply_saved_query(self, *, state: PageState, query_name: str, xhr_seq: int = 2) -> str:
        if query_name not in SAVED_QUERIES:
            raise MMISClientError(f"尚未定義的 MMIS 儲存查詢: {query_name}")
        query_meta = SAVED_QUERIES[query_name]
        return self._post_event(
            currentfocus=query_meta["focus_id"],
            event_type="click",
            target_id="mainrec_menus",
            value=query_meta["menu_value"],
            xhr_seq=xhr_seq,
            referer=state.page_url,
        )

    def _download_to_path(
        self,
        *,
        download_url: str,
        target_path: Path,
        base_filename: str,
        referer: str,
        execution_mode: str,
    ) -> None:
        self.logger.info("base filename = %s", base_filename)
        if target_path.name != base_filename:
            self.logger.info("file exists -> resolving conflict")
        self.logger.info("final filename = %s", target_path.name)
        with self.timed_step("download_result_file", execution_mode=execution_mode):
            download_response = self._request("GET", download_url, headers={"Referer": referer})
            with target_path.open("wb") as file_obj:
                file_obj.write(download_response.content)
            self.logger.info("下載成功: %s", target_path)

    def _run_python_fallback_script(self, script_path: Path) -> dict[str, Any]:
        process = subprocess.run(
            [sys.executable, str(script_path)],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        stdout = process.stdout.strip()
        stderr = process.stderr.strip()
        if process.returncode != 0:
            raise MMISClientError(stderr or stdout or f"fallback 腳本失敗: {script_path}")
        try:
            return json.loads(stdout)
        except json.JSONDecodeError as exc:
            raise MMISClientError(f"fallback 腳本未輸出有效 JSON: {exc}") from exc

    def get_unprocessed_fault_reports(
        self,
        query_name: str = "本段未處理通報(車輛配屬段)",
        target_dir: Path = TARGET_DIR,
        *,
        run_excel_formatter: bool = False,
    ) -> dict[str, Any]:
        if query_name not in SAVED_QUERIES:
            raise MMISClientError(f"尚未定義的 MMIS 儲存查詢: {query_name}")

        result: dict[str, Any] | None = None
        with self.timed_step("download_unprocessed_fault_reports", query_name=query_name):
            self.login()
            target_dir.mkdir(parents=True, exist_ok=True)
            base_filename = build_target_filename()
            target_path = resolve_filename_conflict(target_dir, base_filename)

            download_url: str | None = None
            result_count: int | None = None
            execution_mode = "http"
            self.logger.info("開始查詢故障通報管理")
            self.logger.info("查詢條件: %s", query_name)

            try:
                apply_response = ""
                state: PageState | None = None
                for attempt in range(2):
                    state = self.load_app("zz_fnm")
                    try:
                        menu_response = self._open_saved_query_menu(state=state)
                        if "mainrec_menus" not in menu_response:
                            raise MMISClientError("MMIS 未回傳查詢選單內容，無法套用儲存查詢")
                        apply_response = self._apply_saved_query(state=state, query_name=query_name, xhr_seq=2)
                        break
                    except MMISClientError as exc:
                        if attempt == 0 and "重新登入" in str(exc):
                            self.logger.warning("偵測到 session 失效，重新登入後重試查詢")
                            self.login(force=True)
                            continue
                        raise

                if query_name not in html.unescape(apply_response):
                    raise MMISClientError("查詢套用失敗，回應中找不到已套用的查詢名稱")

                download_url = extract_download_url(apply_response)
                if not download_url:
                    raise MMISClientError("查詢已套用，但找不到下載 URL")

                result_count = extract_result_count(apply_response)
                if result_count is not None:
                    self.logger.info("查詢結果筆數: %s", result_count)
                else:
                    self.logger.info("查詢結果筆數: 無法從回應直接解析")

                self._download_to_path(
                    download_url=download_url,
                    target_path=target_path,
                    base_filename=base_filename,
                    referer=state.page_url if state else self.state.page_url,
                    execution_mode=execution_mode,
                )
            except MMISClientError as exc:
                execution_mode = "browser-fallback"
                self.logger.warning("HTTP 流程失敗，改用最小化瀏覽器 fallback: %s", exc)
                fallback_result = self._playwright_download_unprocessed_fault_reports(
                    query_name=query_name,
                    target_path=target_path,
                )
                result_count = fallback_result["result_count"]
                download_url = fallback_result["download_url"]

            excel_result: dict[str, Any] | None = None
            if run_excel_formatter:
                excel_result = self.run_excel_formatter()

            result = {
                "success": True,
                "query_name": query_name,
                "uisessionid": self.state.ui_session_id if self.state else None,
                "page_url": self.state.page_url if self.state else None,
                "result_count": result_count,
                "download_url": download_url,
                "path": str(target_path),
                "filename": target_path.name,
                "execution_mode": execution_mode,
                "formatted": run_excel_formatter,
                "excel_result": excel_result,
                "log_file": str(LOG_FILE),
                "retry_count": self.request_retry_count,
                "step_metrics": [asdict(metric) for metric in self.step_metrics],
            }
        assert result is not None
        result["step_metrics"] = [asdict(metric) for metric in self.step_metrics]
        return result

    def get_open_b_level_fault_reports(
        self,
        *,
        target_dir: Path = TARGET_DIR,
        depot_name: str = "新竹機務段",
        level: str = "B",
        query_name: str = "故障通報未結案清單",
    ) -> dict[str, Any]:
        level_selection = parse_fault_levels(level)
        depot_name = normalize_depot_name(depot_name)
        result: dict[str, Any] | None = None
        with self.timed_step(
            "download_open_b_level_fault_reports",
            query_name=query_name,
            depot_name=depot_name,
            level=level_selection.display_level,
        ):
            self.login()
            target_dir.mkdir(parents=True, exist_ok=True)
            base_filename = generate_level_filename(
                depot_name=depot_name,
                level_display=level_selection.display_level,
            )
            target_path = resolve_filename_conflict(target_dir, base_filename)

            execution_mode = "http"
            result_count: int | None = None
            download_url: str | None = None
            self.logger.info("開始查詢故障通報管理")
            self.logger.info("input level = %s", level)
            self.logger.info("parsed query = %s", level_selection.query_level)
            self.logger.info("filename level = %s", level_selection.display_level)
            self.logger.info(
                "查詢條件: %s, 配屬段=%s, 等級=%s",
                query_name,
                depot_name,
                level_selection.display_level,
            )

            try:
                response_text = ""
                state: PageState | None = None
                for attempt in range(2):
                    state = self.load_app("zz_fnm")
                    try:
                        menu_response = self._open_saved_query_menu(state=state)
                        if "mainrec_menus" not in menu_response:
                            raise MMISClientError("MMIS 未回傳查詢選單內容，無法套用儲存查詢")
                        response_text = self._apply_saved_query(state=state, query_name=query_name, xhr_seq=2)
                        if query_name not in html.unescape(response_text):
                            raise MMISClientError("查詢套用失敗，回應中找不到已套用的查詢名稱")

                        response_text = self._post_event(
                            currentfocus="m6a7dfd2f_tfrow_[C:5]_txt-tb",
                            xhr_seq=3,
                            referer=state.page_url,
                            events=[
                                {
                                    "type": "setvalue",
                                    "targetId": "m6a7dfd2f_tfrow_[C:15]_txt-tb",
                                    "value": depot_name,
                                    "requestType": "SYNC",
                                    "csrftokenholder": self.state.csrf_token,
                                }
                            ],
                        )
                        response_text = self._post_event(
                            currentfocus="m6a7dfd2f_tfrow_[C:5]_txt-tb",
                            xhr_seq=4,
                            referer=state.page_url,
                            events=[
                                {
                                    "type": "setvalue",
                                    "targetId": "m6a7dfd2f_tfrow_[C:5]_txt-tb",
                                    "value": level_selection.query_level,
                                    "requestType": "SYNC",
                                    "csrftokenholder": self.state.csrf_token,
                                },
                                {
                                    "type": "filterrows",
                                    "targetId": "m6a7dfd2f_tbod_tfrow-tr",
                                    "value": "",
                                    "requestType": "SYNC",
                                    "csrftokenholder": self.state.csrf_token,
                                },
                            ],
                        )
                        self.logger.info(
                            "查詢條件已成功套用: 配屬段=%s, 等級=%s",
                            depot_name,
                            level_selection.query_level,
                        )
                        break
                    except MMISClientError as exc:
                        if attempt == 0 and "重新登入" in str(exc):
                            self.logger.warning("偵測到 session 失效，重新登入後重試 B 級未結案查詢")
                            self.login(force=True)
                            continue
                        raise

                result_count = extract_result_count(response_text)
                if result_count is not None:
                    self.logger.info("查詢結果筆數: %s", result_count)
                else:
                    self.logger.info("查詢結果筆數: 無法從回應直接解析")

                download_url = extract_download_url(response_text)
                if not download_url:
                    raise MMISClientError("查詢完成，但找不到下載 URL")

                self._download_to_path(
                    download_url=download_url,
                    target_path=target_path,
                    base_filename=base_filename,
                    referer=state.page_url if state else self.state.page_url,
                    execution_mode=execution_mode,
                )
            except MMISClientError as exc:
                execution_mode = "browser-fallback"
                self.logger.warning("HTTP 流程失敗，改用 B 級未結案 Playwright fallback: %s", exc)
                fallback_result = self._run_python_fallback_script(OPEN_B_LEVEL_PLAYWRIGHT_SCRIPT)
                result_count = fallback_result.get("result_count")
                download_url = fallback_result.get("download_url")

            result = {
                "success": True,
                "query_name": query_name,
                "filters": {
                    "配屬段": depot_name,
                    "等級": level_selection.display_level,
                    "等級查詢值": level_selection.query_level,
                },
                "uisessionid": self.state.ui_session_id if self.state else None,
                "page_url": self.state.page_url if self.state else None,
                "result_count": result_count,
                "download_url": download_url,
                "path": str(target_path),
                "filename": target_path.name,
                "execution_mode": execution_mode,
                "log_file": str(LOG_FILE),
                "retry_count": self.request_retry_count,
                "step_metrics": [asdict(metric) for metric in self.step_metrics],
            }
        assert result is not None
        result["step_metrics"] = [asdict(metric) for metric in self.step_metrics]
        return result

    def _playwright_download_unprocessed_fault_reports(
        self,
        *,
        query_name: str,
        target_path: Path,
    ) -> dict[str, Any]:
        self.require_credentials()
        if not PLAYWRIGHT_FALLBACK_ENABLED:
            raise MMISClientError("HTTP 失敗且 Playwright fallback 已停用")

        try:
            from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
            from playwright.sync_api import sync_playwright
        except Exception as exc:  # noqa: BLE001
            raise MMISClientError(f"無法載入 Playwright fallback: {exc}") from exc

        download_url = None
        result_count = None
        with self.timed_step("playwright_fallback", query_name=query_name):
            self.logger.info("啟動 headless browser fallback")
            with sync_playwright() as playwright:
                browser = playwright.chromium.launch(headless=True)
                context = browser.new_context(accept_downloads=True)
                page = context.new_page()
                page.set_default_timeout(10000)
                try:
                    page.goto(LOGIN_PAGE_URL, wait_until="domcontentloaded")
                    page.locator("#username").fill(self.username)
                    page.locator("#password").fill(self.password)
                    page.locator("#iamnotrobot").click()
                    page.locator("#loginbutton").click()
                    page.wait_for_timeout(2500)

                    session_match = REDIRECT_SESSION_RE.search(page.url)
                    if not session_match:
                        raise MMISClientError("瀏覽器 fallback 登入後無法取得 uisessionid")
                    ui_session_id = session_match.group("ui_session_id")
                    start_url = (
                        f"{BASE_URL}/maximo/ui/login?uisessionid={ui_session_id}"
                        "&event=loadapp&value=startcntr"
                    )
                    page.goto(start_url, wait_until="domcontentloaded")
                    page.locator("#FavoriteApp_ZZ_FNM").first.click()
                    page.wait_for_timeout(1200)
                    page.locator("#toolbar2_tbs_0_tbcb_0_query-co_2").click()
                    page.wait_for_timeout(600)
                    page.get_by_text(query_name, exact=True).first.click()
                    page.wait_for_timeout(1500)

                    query_value = page.locator("#toolbar2_tbs_0_tbcb_0_query-tb").input_value()
                    if query_name not in query_value:
                        raise MMISClientError("瀏覽器 fallback 未成功套用查詢")

                    href = page.locator("#m6a7dfd2f-lb4").get_attribute("href") or ""
                    decoded_href = html.unescape(href)
                    match = re.search(r'openEncodedURL\("(?P<url>https://[^"]+)"', decoded_href)
                    download_url = match.group("url") if match else None
                    result_count = extract_result_count(page.locator("#m6a7dfd2f-lb3").inner_text())

                    with page.expect_download(timeout=20000) as download_info:
                        page.locator("#m6a7dfd2f-lb4").click()
                    download = download_info.value
                    download.save_as(str(target_path))
                except PlaywrightTimeoutError as exc:
                    raise MMISClientError(f"瀏覽器 fallback 逾時: {exc}") from exc
                finally:
                    browser.close()

        self.logger.info("瀏覽器 fallback 下載成功: %s", target_path)
        return {
            "download_url": download_url,
            "result_count": result_count,
        }

    def run_excel_formatter(self) -> dict[str, Any]:
        with self.timed_step("excel_formatter"):
            self.logger.info("開始整理 Excel 檔案")
            process = subprocess.run(
                [sys.executable, str(EXCEL_FORMATTER)],
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
            )
            if process.returncode != 0:
                self.logger.error("Excel 整理失敗: %s", process.stderr.strip() or process.stdout.strip())
                raise MMISClientError(process.stderr.strip() or process.stdout.strip() or "Excel 整理失敗")
            try:
                result = json.loads(process.stdout.strip())
            except json.JSONDecodeError as exc:
                raise MMISClientError(f"Excel 格式化腳本未輸出有效 JSON: {exc}") from exc
            self.logger.info("Excel 整理完成: saved=%s", result.get("saved"))
            return result


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="MMIS request-driven automation client")
    subparsers = parser.add_subparsers(dest="command", required=True)

    login_parser = subparsers.add_parser("login", help="Login to MMIS and cache the session")
    login_parser.add_argument("--force", action="store_true", help="Ignore cached session and login again")

    query_parser = subparsers.add_parser(
        "download-unprocessed-fault-reports",
        help="Download fixed saved query result from 故障通報管理",
    )
    query_parser.add_argument(
        "--query-name",
        default="本段未處理通報(車輛配屬段)",
        help="Saved query name",
    )
    query_parser.add_argument(
        "--target-dir",
        default=str(TARGET_DIR),
        help="Directory to save the downloaded file",
    )
    query_parser.add_argument(
        "--format-excel",
        action="store_true",
        help="Run Excel formatter after download",
    )

    b_level_parser = subparsers.add_parser(
        "download-open-b-level-fault-reports",
        help="Download B-level open fault notice report from 故障通報管理",
    )
    b_level_parser.add_argument(
        "--target-dir",
        default=str(TARGET_DIR),
        help="Directory to save the downloaded file",
    )
    b_level_parser.add_argument(
        "--depot-name",
        "--depot",
        dest="depot_name",
        default="新竹機務段",
        help="Depot filter value",
    )
    b_level_parser.add_argument("--level", default="B", help="Fault level filter value")
    return parser


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    parser = build_parser()
    args = parser.parse_args()
    client = MMISClient()

    try:
        if args.command == "login":
            result = client.login(force=args.force)
        elif args.command == "download-unprocessed-fault-reports":
            result = client.get_unprocessed_fault_reports(
                query_name=args.query_name,
                target_dir=Path(args.target_dir),
                run_excel_formatter=args.format_excel,
            )
        elif args.command == "download-open-b-level-fault-reports":
            result = client.get_open_b_level_fault_reports(
                target_dir=Path(args.target_dir),
                depot_name=args.depot_name,
                level=args.level,
            )
        else:
            raise MMISClientError(f"未知命令: {args.command}")
    except Exception as exc:  # noqa: BLE001
        client.logger.error("MMIS 執行失敗: %s", exc)
        print(
            json.dumps(
                {
                    "success": False,
                    "command": args.command,
                    "reason": str(exc),
                    "log_file": str(LOG_FILE),
                    "retry_count": client.request_retry_count,
                    "step_metrics": [asdict(metric) for metric in client.step_metrics],
                },
                ensure_ascii=False,
            )
        )
        return 1

    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
