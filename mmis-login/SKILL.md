---
name: mmis-login
description: 以程式化方式登入 MMIS 並建立可重用 session。當使用者提到「登入 MMIS」、「進入 MMIS」或任何需要先取得 MMIS 有效 session 的工作時使用。預設走 `mmisClient.py` 的 HTTP/session 流程；只有在使用者明確要求截圖或純瀏覽器驗證時，才退回最小化 Playwright。
---

# MMIS Login

此 skill 不再以逐步 Playwright 操作為主流程。預設做法是直接呼叫核心程式：

```powershell
python C:\Users\NMMIS\.codex\skills\mmis-login\scripts\run_mmis_login.py
```

實際登入核心位於：

- `C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts\mmisClient.py`

## 目標

- 直接用 HTTP/session 登入 MMIS
- 建立並快取可重用 session
- 將執行狀態寫入 console 與 log 檔
- 回傳 JSON 結果

## 前置條件

- `MMIS_USERNAME` 已設定
- `MMIS_PASSWORD` 已設定

若任一缺失，立即停止並回報缺少哪些環境變數。

## 執行原則

- 預設只呼叫核心程式，不要再逐步 click 登入畫面。
- 優先沿用快取 session；只有快取失效時才重新登入。
- 登入成功判斷以程式回傳 JSON 為準，不靠人工觀察頁面。
- log 必須同時寫到 console 與 `logs/mmis.log`。

## 成功判斷

以下條件成立即視為成功：

- 回傳 JSON 中 `logged_in=true`
- 有 `uisessionid`
- 有 `log_file`

## 失敗回報

回覆時至少包含：

- 是否成功登入
- 是否沿用既有 session
- `uisessionid`
- `log_file`
- 若失敗，明確錯誤原因

## Playwright 例外情況

只有在下列情況才使用 Playwright：

- 使用者明確要求登入後截圖
- 使用者要求檢視登入後畫面
- HTTP/session 流程失敗，且需要最小化瀏覽器作為 fallback 驗證
