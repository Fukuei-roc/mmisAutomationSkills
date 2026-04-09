# MMIS 網頁資料定位指南

這份指南的目的，是讓後續開發 MMIS Playwright、自動化查詢、頁面分析、資料擷取時，可以更快找到真正的資料位置，並選出更穩定的 selector。

## 1. 基本觀念

- MMIS 建立在 IBM Maximo 架構上。
- Maximo 頁面常帶有大量動態生成的 `id`，不能把單一完整 `id` 視為永遠穩定。
- 畫面上看得到的資料，不一定就直接放在最外層元素。
- 真正資料常出現在：
  - `input.value`
  - `td` 內部的 `span`
  - `span.title`
  - `innerText`
- 做自動化時，必須以 DOM 結構為準，不要只憑畫面外觀判斷。

## 2. 使用 DevTools 定位元素

### 基本流程

1. 在 Chrome 中打開 MMIS 頁面。
2. 右鍵目標資料，選 `檢查`。
3. 使用 DevTools 左上角的元素選取工具，直接點畫面上的欄位或資料。
4. 在 Elements 面板中確認實際 HTML 結構。

### 觀察重點

定位元素時，至少確認以下資訊：

- tag：
  - `input`
  - `td`
  - `tr`
  - `span`
  - `a`
- 屬性：
  - `id`
  - `name`
  - `title`
  - `href`
  - `aria-label`
- 文字來源：
  - `innerText`
  - `textContent`
  - `value`
- 結構關係：
  - 父層是什麼
  - 是否位於 table row / cell 中
  - 是否被另一層 `span` 包住

### 實務判斷

- 若是輸入欄位，先看 `input.value`，不要只看畫面文字。
- 若是表格資料，先看 `td` 裡面是否有 `span[title]`。
- 若畫面顯示有內容，但 `td` 本身是空的，通常資料在子節點。

## 3. selector 選擇策略

### 優先順序

1. 穩定文字
2. 結構關係
3. 穩定屬性
4. 部分可辨識的 `id`

### 建議用法

#### 1. 穩定文字

適合：

- 功能入口
- 選單
- 頁面標題
- 固定按鈕文字

例子：

```xpath
//a[contains(normalize-space(.), '動力車日檢(1A)')]
```

#### 2. 結構關係

適合：

- table 內資料
- label 與輸入欄位的關聯
- 父子層級穩定，但 id 不穩定的情況

例子：

```xpath
//td[.//span[@title]]//span[@title]
```

#### 3. 穩定屬性

適合：

- `title`
- `href`
- `aria-label`
- 某些固定欄位代碼

例子：

```xpath
//a[contains(@href, 'ZZ_PMWO1A')]
```

#### 4. 部分可辨識的 id

Maximo 的完整 `id` 常會變，但其中某些片段代表欄位位置，仍可利用。

例子：

```xpath
//input[contains(@id, 'tfrow_') and contains(@id, '[C:5]_txt-tb')]
```

這種寫法比直接寫死完整 id 更耐變動。

### 避免的做法

- 完全依賴完整動態 `id`
- 極長、極脆弱的絕對 XPath
- 只依賴 class name
- 只抓第一層節點卻不確認資料是否在子節點

## 4. 表格資料定位技巧

### 先辨識表格結構

MMIS / Maximo 常見表格資料結構：

- `tr` 代表一列
- `td` 代表一個儲存格
- 真正資料可能在 `td > span`

### Maximo 常見 row / column 規則

常見 id 片段：

- `C:n`
  - 欄位 index
- `R:n`
  - 列 index

例如：

```text
m6a7dfd2f_tdrow_[C:5]-c[R:0]
```

可解讀為：

- 第 `5` 欄
- 第 `0` 列

### 多筆資料抓取方式

不要只抓 `R:0`。

建議做法：

1. 先找出所有符合條件的列或儲存格
2. 逐列迭代
3. 從 `td` 進一步抓 `span`
4. 取 `title` 或 `innerText`

例子：

```xpath
//td[contains(@id,'_tdrow_') and contains(@id,'[C:5]-c[R:')]
```

或：

```xpath
//td[.//span[@title]]//span[@title]
```

### 實務原則

- 若要支援多筆，永遠先取集合，再逐筆處理。
- 不要把資料定位綁死在第一列。
- 若欄位值常在 `title`，就優先讀 `title`，再退回 `innerText`。

## 5. 判斷資料狀態

自動化程式必須能分辨以下情況：

### 1. 查無資料

特徵：

- 查詢後結果列不存在
- 或頁面顯示 `查無資料`、`沒有資料`
- 或 table 存在，但結果 row count 為 `0`

### 2. 欄位存在但無值

特徵：

- element 存在
- `value` 為空字串
- `title` 為空
- `innerText` 為空白

這代表欄位存在，但資料本身沒有值，不能誤判成 selector 失敗。

### 3. 單筆資料

特徵：

- 結果集合長度為 `1`

### 4. 多筆資料

特徵：

- 結果集合長度大於 `1`

### 實務建議

- 先判斷 element 是否存在。
- 再判斷資料值是否為空。
- 最後再判斷資料筆數。

## 6. Playwright 實務建議

- 優先使用 `locator()`，不要過度依賴一次性 query。
- 優先使用：
  - `locator.wait_for()`
  - `page.wait_for_url()`
  - `page.wait_for_load_state("networkidle")`
- 不要使用固定 `sleep`。
- 操作前先確認元素存在且可見。
- 對於容易波動的步驟加 retry。
- 查詢結果不要只判斷 click 成功，要再驗證：
  - 是否真的進入目標頁
  - 欄位值是否與輸入一致

### 建議模式

```python
locator = page.locator(selector)
locator.wait_for(state="visible", timeout=30000)
locator.fill(value)
locator.press("Enter")
```

### 容錯原則

- selector 找不到時，要回傳清楚錯誤
- 查無資料時，要和 selector 失敗分開處理
- 欄位存在但空值時，不可直接當作異常

## 7. MMIS 常見模式（實戰經驗）

### 1. Maximo table id 有規律，但完整 id 不穩定

- 可利用 `C:n`、`R:n` 片段判斷欄位與列
- 不建議把整個 id 視為永久穩定值

### 2. 資料常放在 span.title 或 span 文字

常見情況：

- 畫面中一個表格欄位，真正資料在：

```html
<td ...>
  <span title="1150331-39">1150331-39</span>
</td>
```

這時應優先抓：

- `title`
- 若 `title` 空，再讀 `innerText`

### 3. 查詢常由 Enter 觸發

在許多 MMIS filter row 中：

- `fill()`
- `press("Enter")`

就會觸發查詢或表格刷新。

### 4. DOM 更新常晚於點擊

即使 click 完成，欄位或結果列不一定立刻可用。

因此：

- 先等 URL
- 再等目標欄位
- 再做下一步

這比只等 `networkidle` 更穩定。

### 5. 查詢欄位與結果欄位常是不同結構

- 查詢欄位通常是 `input`
- 結果資料通常是 `td` 或 `span`

不要把查詢欄 selector 的策略直接套到結果資料。

## 8. 開發時的推薦流程

1. 在畫面上找到目標資料
2. 用 DevTools 檢查真正 DOM
3. 確認資料實際存放位置：
   - `value`
   - `innerText`
   - `title`
4. 找出最穩定的 selector 組合
5. 驗證 4 種情況：
   - 查無資料
   - 空值
   - 單筆
   - 多筆
6. 再寫入 Playwright 自動化程式

## 9. 建議結論

對 MMIS / Maximo 開發來說，最重要的不是「先寫 selector」，
而是先確認：

- 資料真正在哪個節點
- 哪些屬性穩定
- 查詢後 DOM 怎麼變
- 無資料時畫面會怎麼呈現

只有先把這些看清楚，後續 Playwright 與 HTTP 分析才會穩定。
