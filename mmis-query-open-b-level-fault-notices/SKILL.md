---
name: mmis-query-open-b-level-fault-notices
description: 以 HTTP/session 為主流程查詢 MMIS 的未結案 A/B/C 級故障通報並下載報表；支援 level 單一或組合值與 depot 參數，若 HTTP 失敗才回退到 Playwright fallback。
---

# MMIS Query Open B-Level Fault Notices

此 skill 目前已進入 Phase 2，主流程改為呼叫 `mmisClient.py` 的 HTTP/session 流程；只有在 HTTP 失敗時才回退到既有 Playwright 腳本。

執行方式：

```powershell
python C:\Users\NMMIS\.codex\skills\mmis-query-open-b-level-fault-notices\scripts\run_open_b_level_fault_notice_download.py
```

可選參數：

```powershell
python C:\Users\NMMIS\.codex\skills\mmis-query-open-b-level-fault-notices\scripts\run_open_b_level_fault_notice_download.py --level C --depot 七堵機務段
```

核心腳本：

- `C:\Users\NMMIS\.codex\skills\mmis-query-unprocessed-fault-notices\scripts\mmisClient.py`

Playwright fallback：

- `C:\Users\NMMIS\.codex\skills\mmis-query-open-b-level-fault-notices\scripts\playwright_open_b_level_fault_notice_download.py`

## 功能

- 共用既有 MMIS 登入邏輯與 session
- 直接送出 `maximo.jsp` event request：
  - `changeapp`
  - 開啟 query menu
  - 套用 `故障通報未結案清單`
  - `setvalue` 配屬段
  - `setvalue + filterrows` 等級與查詢
- 直接下載報表
- 依 `level` 參數儲存為 `{單位簡稱}{level}級未結案故障通報MMDD.xlsx`

## 參數

- `level`
  - 可選 `A` / `B` / `C`
  - 或組合 `AB` / `AC` / `BC` / `ABC`
  - 預設 `B`
- `depot`
  - 字串
  - 預設 `新竹機務段`

`level` 解析規則：

- `A` -> 查詢值 `A`
- `B` -> 查詢值 `B`
- `C` -> 查詢值 `C`
- `AB` -> 查詢值 `A,B`
- `AC` -> 查詢值 `A,C`
- `BC` -> 查詢值 `B,C`
- `ABC` -> 查詢值 `A,B,C`

若不傳參數，等同：

- `level=B`
- `depot=新竹機務段`

## 執行原則

- 預設只走 HTTP/session，不開瀏覽器。
- 優先沿用 `MMISClient.login()` 建立的 session 與 cookies。
- 只有 HTTP 失敗時才回退到 Playwright fallback。
- log 必須同時輸出到 console 與 `logs/mmis.log`。

## 成功判斷

以下條件成立即視為成功：

- 已成功登入 MMIS
- 已成功進入 `故障通報管理`
- 已成功套用 `故障通報未結案清單`
- 篩選條件 `level/depot` 已送出
- 已成功觸發下載
- 已將檔案覆寫存成 `{單位簡稱}{level}級未結案故障通報MMDD.xlsx`

## 失敗回報

回覆時至少包含：

- 是否沿用既有 session
- 是否成功登入
- 查詢條件
- 若可解析，查詢結果筆數
- 最終檔案完整路徑
- `log_file`
- 失敗步驟與原因

## 參數驗證

- `level` 僅允許 `A` / `B` / `C`
- `level` 可用 `A` / `B` / `C` 的任意不重複組合
- `depot` 若未提供則使用預設值
- `depot` 若為空字串則立即失敗

## 後續規劃

若 MMIS event/request pattern 改版，再修正 `mmisClient.py` 的 HTTP 流程；Playwright 僅保留作 contingency fallback。
