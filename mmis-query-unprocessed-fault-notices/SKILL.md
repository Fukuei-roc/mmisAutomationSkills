---
name: mmis-query-unprocessed-fault-notices
description: 以程式化方式查詢 MMIS 的「本段未處理通報(車輛配屬段)」、下載結果並可選擇直接整理 Excel。預設走 `mmisClient.py` 的 HTTP/session/event 流程，不再依賴逐步 Playwright 點選。
---

# MMIS Query Unprocessed Fault Notices

此 skill 的主流程是呼叫核心程式，而不是逐步操作瀏覽器：

```powershell
python C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts\run_unprocessed_fault_notice_download.py
```

若使用者同時要求下載後整理 Excel，改用：

```powershell
python C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts\run_unprocessed_fault_notice_download.py --format-excel
```

核心程式：

- `C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts\mmisClient.py`

## 分工

- `mmisClient.py`
  - 登入 MMIS
  - 重用 session
  - 送出 `maximo.jsp` event request
  - 套用儲存查詢 `本段未處理通報(車輛配屬段)`
  - 下載結果檔案
  - 視需要串接 Excel 格式化腳本
- `format_mmis_excel.py`
  - 專責整理下載後的 Excel 檔案

## 執行原則

- 能直接走 HTTP/session 就不要開瀏覽器。
- 能直接送 `maximo.jsp` event request 就不要 click UI。
- 優先沿用快取 session，避免重複登入。
- 下載檔名固定為 `本段未處理故障通報MMDD.xlsx`。
- log 必須寫到 console 與 `logs/mmis.log`。

## 成功判斷

以下條件成立即視為成功：

- 查詢條件成功套用為 `本段未處理通報(車輛配屬段)`
- 成功解析下載 URL
- 成功下載到 `C:\Users\NMMIS\OneDrive - Ministry of Transportation and Communications-7280502-Taiwan Railways Administration, MOTC\文件\MMIS桌面`
- 回傳 JSON 含 `success=true`

若有加上 `--format-excel`，還必須包含：

- `formatted=true`
- `excel_result.saved=true`

## 失敗回報

回覆時至少包含：

- 是否沿用既有 session
- 查詢條件
- 查詢結果筆數
- 最終儲存完整路徑
- `log_file`
- 若失敗，明確錯誤原因

## Playwright 例外情況

只有在以下情況才退回最小化 Playwright：

- MMIS request/event pattern 改版，`mmisClient.py` 無法完成查詢
- 使用者明確要求畫面驗證或截圖
