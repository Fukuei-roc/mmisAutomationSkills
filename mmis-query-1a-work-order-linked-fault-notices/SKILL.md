---
name: mmis-query-1a-work-order-linked-fault-notices
description: 使用 Playwright 查詢 MMIS「動力車日檢(1A)」工單所勾稽的故障通報資料。當使用者提供 1A 工單號、要求查詢工單勾稽的故障通報清單、或需要截圖與明細頁驗證時使用。輸入為工作單號，輸出為故障通報清單，支援 0 筆、1 筆與多筆。
---

# MMIS Query 1A Work Order Linked Fault Notices

此 skill 使用 Playwright 查詢 MMIS「動力車日檢(1A)」工單明細，擷取其勾稽的故障通報清單。

執行方式：

```powershell
python C:\Users\NMMIS\.codex\skills\mmis-query-1a-work-order-linked-fault-notices\scripts\run_linked_fault_notices_query.py --work-order-no 115-1A-23391
```

核心腳本：

- `C:\Users\NMMIS\.codex\skills\mmis-query-1a-work-order-linked-fault-notices\scripts\playwright_linked_fault_notices_query.py`

共用登入核心：

- `C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts\mmisClient.py`

## 功能

- 沿用既有 MMIS login/session 機制
- 進入 `動力車日檢(1A)`
- 依工作單號查詢
- 進入第一筆工單明細
- 擷取所有勾稽故障通報 `span[title]`
- 截圖存到 `C:\Users\NMMIS\Downloads`

## 參數

- `work_order_no`
  - 必填
  - 例如 `115-1A-23391`

## 執行原則

- 優先使用 request-driven login + browser cookie bootstrap
- Playwright 只負責後續 UI 查詢與明細頁擷取
- 不使用固定 sleep，改用 `wait_for_selector`、`networkidle`、`locator`
- selector 優先用文字、href 結構、`contains(@id, ...)`，不要完全依賴單一動態 id

## 成功判斷

以下條件成立即視為成功：

- 已成功登入 MMIS
- 已成功進入 `動力車日檢(1A)`
- 已成功查到工作單並進入明細
- 明細頁工作單號與輸入一致
- 已成功儲存截圖
- 已成功回傳 `fault_notices`

## 回傳格式

成功：

```json
{
  "work_order": "115-1A-23391",
  "fault_notices": ["1150331-39", "1150331-49"]
}
```

查無勾稽：

```json
{
  "work_order": "115-1A-23391",
  "fault_notices": []
}
```

找不到工單：

```json
{
  "error": "找不到工作單"
}
```

## 失敗回報

回覆時至少包含：

- 工作單號
- 是否登入成功
- 是否找到工單
- 是否成功進入明細
- 故障通報筆數
- 截圖路徑
- `log_file`
- 若失敗，具體原因
