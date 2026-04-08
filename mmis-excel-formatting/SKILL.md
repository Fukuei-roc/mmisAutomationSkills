---
name: mmis-excel-formatting
description: 依檔名自動判斷並格式化 MMIS 匯出的 Excel 檔案。當使用者要求整理最近下載的 MMIS Excel 檔、或已知要套用 `故障通報管理` 類型格式時使用。支援自動偵測檔案類型與透過參數指定檔案類型。
---

# MMIS Excel Formatting

此 skill 只負責處理 MMIS 匯出的 Excel 檔案，不負責登入、查詢、下載。

## 何時使用

- 使用者要求整理最近下載的 MMIS Excel 檔。
- 使用者要求整理 `故障通報管理` 類型 Excel。
- 使用者要求將工作表 `故障通報管理 的清單` 的字型大小統一為 12。
- 使用者要求將工作表 `故障通報管理 的清單` 套用整體版面格式。
- 使用者要求依 `發生日期` 及 `車組/車號` 排序後直接覆寫原檔。
- 使用者要求在 B 與 C 欄之間新增 `車號` 欄位，並將計算結果填滿到最後一列。

## 何時不要使用

- 使用者要做欄位清理、公式、圖表、彙總、或另存新檔。
- 使用者尚未先下載 MMIS Excel 檔案。

## 執行規則

- 先執行腳本 [format_mmis_excel.py](C:\Users\NMMIS\.codex\skills\mmis-excel-formatting\scripts\format_mmis_excel.py)。
- 路徑中的空格與特殊字元要原樣處理，不可自行簡化或截斷。
- 預設優先接續處理最近產生的 `.xlsx` 檔案。
- 若 `--file-type auto`，先根據檔名判斷 formatter。
- 若檔名包含 `故障通報管理` 或 `未處理故障通報`，套用 `fault_notice` formatter。
- 也可用 `--file-type fault_notice` 強制指定 formatter。
- 若找不到目標檔案，停止並列出目標資料夾實際存在的檔案名稱。
- 若工作表 `故障通報管理 的清單` 不存在，停止並列出所有工作表名稱。
- 若欄位 `發生日期` 或 `車組/車號` 不存在，停止並列出實際欄位名稱。
- 排序時必須以整個資料列為單位重排，不可只排序單一欄位。
- 必須在 B 欄右側插入新欄位，將第 1 列設為 `車號`。
- 必須將整個工作表所有儲存格套用字型 `新細明體`、大小 `12`。
- 必須先將整個工作表所有儲存格設為水平靠左、垂直靠上。
- 必須再將第 1 列標題列覆蓋為水平置中、垂直置中。
- 必須將 `C1` 的 `車號` 標題設為粗體。
- 必須依欄位名稱刪除以下整欄，且由右到左刪除：
  - `發生時間`
  - `故障地點`
  - `立案人員`
  - `通報人員`
  - `通報單位`
  - `狀態`
  - `配屬段別`
  - `配屬段別名稱`
- 必須從 B 欄值解析尾端數字，依 Python 邏輯直接寫入最終結果到 `C2` 到 `C{最後列}`。
- 必須依內容自動調整所有欄位欄寬，不可寫死固定寬度。
- 不可在 `車號` 欄留下任何 Excel 公式。
- 儲存時直接覆寫原檔，不另存新檔。

## 流程

1. 找出最近下載的 `.xlsx` 檔，或使用明確指定的檔案路徑。
2. 根據檔名或 `--file-type` 判斷 formatter。
3. 若檔名包含 `故障通報管理` 或 `未處理故障通報`，套用 `formatFaultNoticeExcel(file_path)`。
4. 開啟該檔案。
5. 確認工作表 `故障通報管理 的清單` 存在。
6. 將整個工作表使用中的所有儲存格套用字型 `新細明體`、大小 `12`。
7. 確認欄位 `發生日期` 與 `車組/車號` 存在。
8. 以整個資料表為範圍排序：
   - 第一排序鍵：`發生日期`，由舊到新
   - 第二排序鍵：`車組/車號`，由小到大
9. 在 B 與 C 欄之間插入新欄位，將 `C1` 設為 `車號`。
10. 將整個工作表使用中的所有儲存格套用水平靠左、垂直靠上。
11. 將 `C1` 設為粗體。
12. 依欄位名稱找出指定欄位，並由右到左刪除。
13. 對每一列資料：
   - 從 B 欄字串解析尾端連續數字
   - 若解析出 `n` 且 `1000 <= n < 10000`，寫入 `INT(n / 10)`
   - 否則直接寫入 `n`
   - 若無法解析尾端數字，寫入空白
14. 將第 1 列標題列覆蓋為水平置中、垂直置中。
15. 依內容自動調整所有欄位欄寬。
16. 直接覆寫原檔。

## 成功判斷

以下條件全部成立才算成功：

- 成功找到目標檔案
- 成功判斷 formatter 類型
- 成功開啟工作表 `故障通報管理 的清單`
- 成功套用字型格式
- 成功套用整體對齊格式
- 成功完成排序
- 成功插入 `車號` 欄位
- 成功將 `C1` 套用粗體
- 成功刪除指定欄位
- 成功將 `車號` 最終值寫入整個資料列範圍
- 成功自動調整欄寬
- 成功覆寫儲存原檔

## 失敗回報

優先使用下列原因分類：

- 找不到目標檔案
- 無法判斷檔案類型
- 尚未支援的檔案類型
- 工作表不存在
- 欄位名稱不一致
- 格式套用失敗
- 排序失敗
- 找不到 B 欄
- 資料列為 0
- 車號寫入失敗
- 儲存失敗

若失敗，回覆時要包含：

- 是否成功找到檔案
- 使用的檔名
- `file_type`
- `detected_type`
- `selected_by`
- `deleted_headers`
- `header_bold_applied`
- `layout_applied`
- `autofit_applied`
- 是否成功套用格式
- 是否成功完成排序
- 是否成功插入欄位
- 欄位名稱是否正確
- 值寫入範圍
- 無法解析尾端數字的筆數
- 是否成功儲存
- 具體原因
- 若適用，附上實際檔案清單、工作表名稱、或欄位名稱

## 執行方式

優先直接執行：

```powershell
python C:\Users\NMMIS\.codex\skills\mmis-excel-formatting\scripts\format_mmis_excel.py
```

若要指定檔案：

```powershell
python C:\Users\NMMIS\.codex\skills\mmis-excel-formatting\scripts\format_mmis_excel.py --file "C:\Users\NMMIS\...\本段B級故障通報管理0408.xlsx"
```

若要強制指定類型：

```powershell
python C:\Users\NMMIS\.codex\skills\mmis-excel-formatting\scripts\format_mmis_excel.py --file-type fault_notice
```

腳本會輸出 JSON，包含：

- `file_found`
- `filename`
- `file_type`
- `detected_type`
- `selected_by`
- `format_applied`
- `sorted`
- `saved`
- `reason`
- `existing_files`
- `sheet_names`
- `headers`
- `path`
- `column_inserted`
- `header_set`
- `value_range`
- `value_verification`
- `unparsed_count`
