---
name: mmis-auto-link-unprocessed-fault-notices
description: 批次讀取當日本段未處理故障通報 Excel，重用 MMIS session 與同一個 1A 頁面，逐筆查詢動力車日檢(1A)工單，並把第一筆找到的工作單號回填到 Excel I 欄。當使用者要求自動勾稽未處理故障通報對應的 1A 日檢單、批次寫回 Excel、或更新當日未處理故障通報檔時使用。
---

# MMIS Auto Link Unprocessed Fault Notices

此 skill 是批次工具，不是單筆查詢工具。

執行方式：

```powershell
python C:\Users\NMMIS\.codex\skills\mmis-auto-link-unprocessed-fault-notices\scripts\run_auto_link_unprocessed_fault_notices.py
```

可選參數：

- `--skip-filled`
  - 若 I 欄已有資料則跳過
- `--file <path>`
  - 指定 Excel 檔案，否則自動找當日 `本段未處理故障通報MMDD.xlsx`

預設續跑行為：

- 即使不帶 `--skip-filled`，也會自動略過已完成列：
  - 已寫入工單號
  - `找不到日檢單`
  - `缺少查詢條件`
- `查詢失敗` 或空白列會在下次重跑時再嘗試

核心腳本：

- `C:\Users\NMMIS\.codex\skills\mmis-auto-link-unprocessed-fault-notices\scripts\auto_link_unprocessed_fault_notices.py`

重用的既有能力：

- HTTP login / session cache：
  - `C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts\mmisClient.py`
- Playwright 1A 頁面與 session reuse：
  - `C:\Users\NMMIS\.codex\skills\mmis-query-1a-work-order-linked-fault-notices\scripts\playwright_linked_fault_notices_query.py`

## 功能

- 自動尋找當日 `本段未處理故障通報MMDD.xlsx`
- 讀取 Excel 並解析欄位名稱
- 從第 2 列逐筆查詢
- 進入 `動力車日檢(1A)`
- 設定查詢模式為 `所有記錄`
- 依 `新竹機務段 + 發生日期 + 車號` 查詢
- 找到第一筆工單後回填到 Excel `I` 欄
- 找不到時寫入 `找不到日檢單`
- 單筆失敗不中斷整批
- 每筆查詢在按 Enter 前輸出完整頁面 debug 截圖到 `C:\Users\NMMIS\Downloads\mmis_query_debug\query_debug_<row>.png`
- 每列處理完立即存檔，避免批次中斷時丟失進度

## 執行原則

- 穩定優先於速度
- 不使用固定 sleep
- 重用同一個 browser / context / page
- 不可每筆重啟 browser
- 不可每筆重新登入
- 只寫入 `I` 欄，不改其他資料欄
- 若單列查詢失敗，先寫回 `查詢失敗` 並立即存檔，再重建 browser/session 狀態後繼續下一列

## 主要流程

1. 找當日 Excel
2. 讀取工作表與標題列
3. 初始化 MMIS session 與 Playwright
4. 進入 `動力車日檢(1A)`
5. 逐列清空舊 filter，填入新條件並查詢
6. 取得第一筆工單號或寫入 `找不到日檢單`
7. 定期 autosave，最後覆寫原檔

## 回傳格式

成功：

```json
{
  "ok": true,
  "file_path": "C:\\...\\本段未處理故障通報0409.xlsx",
  "total_rows": 12,
  "success_count": 8,
  "fail_count": 4
}
```

找不到當日檔案：

```json
{
  "ok": false,
  "error": "找不到當日未處理故障通報檔案"
}
```

## Log

至少會輸出：

- `file loaded`
- `total rows`
- `processing row`
- `searching: 日期=... 車號=...`
- `found work order: ...`
- `no result`
- `completed`
- `success count`
- `fail count`
