# AI 資料格式整理工具 — 開發紀錄

## 專案基本資訊

| 項目 | 內容 |
|------|------|
| GitHub Repo | https://github.com/carolkao-fin/data-formatter |
| Streamlit Cloud | 部署後網址顯示於 share.streamlit.io |
| 主程式 | `app.py` |
| AI 模型 | Groq — `llama-3.3-70b-versatile`（免費） |
| API Key | Groq（`gsk_...`），免費申請：console.groq.com |

---

## 功能說明

### 單一整理模式
使用者上傳：
1. **原始資料**（CSV / xlsx）— 支援**多檔同時上傳，自動合併**
2. **目標格式範例**（CSV / xlsx / docx）— 想整理成的樣子

AI 自動分析兩份檔案的結構，決定欄位映射與轉換方式，輸出整理後的檔案。

### 批次整理模式
設定多組「原始資料（可多檔合併）＋目標格式」，一次處理所有群組，打包成 ZIP 下載。

### 支援的轉換類型

| 類型 | 說明 | 範例 |
|------|------|------|
| `direct` | 直接複製或改名 | `name` → `姓名` |
| `merge` | 合併多個欄位 | `first_name` + `last_name` → `全名` |
| `value_map` | 值對應翻譯 | `M/F` → `男/女` |
| `date_format` | 日期格式轉換 | `2024-01-15` → `2024/01/15` |
| `number_fmt` | 數字格式化 | `1234567` → `1,234,567` |
| `empty` | 無對應欄位，留空 | — |

### 輸出格式

目標格式的副檔名決定輸出格式：
- `.xlsx` / `.csv` → 輸出 Excel
- `.docx` → 輸出 Word（保留原始樣式，填入表格資料）

---

## 檔案結構

```
data_formatter/
├── app.py              # 主程式
├── requirements.txt    # 套件清單
├── .gitignore          # 排除 format_history.json、secrets
└── 開發紀錄.md         # 本文件
```

`format_history.json` — 操作歷史記錄，執行時自動產生，不推送至 GitHub。

---

## 套件清單（requirements.txt）

```
streamlit>=1.31.0
groq>=0.9.0
pandas>=2.0.0
openpyxl>=3.1.0
xlrd>=2.0.0
python-docx>=1.1.0
```

---

## Streamlit Cloud 部署設定

### Secrets（必填）

在 Streamlit Cloud → App → Settings → Secrets 填入：

```toml
GROQ_API_KEY = "gsk_你的key"
```

### 部署參數

| 欄位 | 值 |
|------|-----|
| Repository | `carolkao-fin/data-formatter` |
| Branch | `master` |
| Main file path | `app.py` |

---

## 更新流程

修改 `app.py` 後，在 `data_formatter/` 目錄執行：

```bash
git add app.py
git commit -m "說明改了什麼"
git push
```

Streamlit Cloud 偵測到 push 後會自動重新部署（約 1 分鐘）。

---

## 核心程式架構

### 1. `read_raw_file(uploaded)` — 讀取原始資料
- 支援 CSV、xlsx、xls
- 回傳 `pd.DataFrame`

### 2. `read_target_file(uploaded)` — 讀取目標格式
- Excel/CSV → `{"type": "excel", "df": DataFrame}`
- Word → `{"type": "word", "tables": [...], "paragraphs": [...]}`
- 同時保留原始 bytes 供 Word 輸出使用

### 3. `describe_raw_df(df)` — 產生原始資料的結構描述
- 欄位名稱、資料型態、範例值
- 唯一值 ≤ 12 種時列出全部（幫助 AI 理解值的語意）

### 4. `describe_target(target)` — 產生目標格式的結構描述
- Excel：逐欄描述型態與範例
- Word：描述表格標題列與段落結構

### 5. `get_mapping(client, raw_df, target)` — 呼叫 Groq AI
- 將兩份結構描述送給 Llama 3.3 70B
- 回傳 JSON：`structure_analysis` + `mappings` 陣列

### 6. `apply_mapping_to_df(raw_df, target_cols, mapping)` — 套用映射
- 依照 `transform` 類型執行對應的轉換邏輯
- 回傳整理後的 `pd.DataFrame`

### 7. `generate_excel(result_df)` — 輸出 Excel
- 使用 openpyxl

### 8. `generate_word(result_df, target, mapping)` — 輸出 Word
- 複製原始 Word 模板結構
- 保留標題列樣式，清除資料列後逐列填入
- 若模板無表格，自動建立 Table Grid 樣式的表格

---

## Word 格式自動偵測機制

`_fill_word_table` 以**欄位名稱文字比對**（非視覺格式）進行填入：

| 情況 | 是否自動處理 | 說明 |
|------|-------------|------|
| 欄位名稱順序不同 | ✅ 自動 | `col_map` 按名稱對應，不依位置 |
| 垂直合併格（vMerge） | ✅ 自動 | `_is_vmerge_continuation()` 偵測並跳過延續格 |
| 多列標題（row 0 非欄位名） | ❌ 不支援 | 程式只讀 `rows[0]`，會誤把合併標題當欄位 |
| 水平合併格（hSpan，一格跨多欄） | ❌ 不支援 | 同一 `_tc` 可能被重複寫入 |
| AI 映射欄名與 Word 標題不一致 | ❌ 不支援 | `col_map` 對應不到，整欄空白 |

> 若遇到填入異常（空白或標籤被覆蓋），優先確認 Word 模板的第一列是否即為欄位名稱列。

---

## 已知限制與未來可改進方向

| 問題 | 說明 | 可能的解法 |
|------|------|-----------|
| Groq TPM 限制 | 免費方案每分鐘 token 有上限，大量欄位時可能超限 | 縮短 prompt，或改用 Groq 付費方案 |
| Word 複雜版型 | 合併儲存格、巢狀表格等複雜結構可能填入異常 | 針對複雜模板加入預處理邏輯 |
| 操作歷史不持久 | Streamlit Cloud 重啟後 `format_history.json` 消失 | 改用 Google Sheets 或 Supabase 儲存歷史 |
| 大型檔案 | pandas 讀取超大 Excel 可能慢或 OOM | 加入分批讀取或限制最大行數提示 |
| AI 映射準確率 | 欄位語意不明確時可能映射錯誤 | 讓使用者在結果頁手動調整映射後重新生成 |

---

## 開發時間軸

| 日期 | 內容 |
|------|------|
| 2026-05-05 | 初版建立：Anthropic API + Excel 輸出 |
| 2026-05-05 | 新增 Word (.docx) 輸出支援 |
| 2026-05-05 | UI 改版：仿照 vitality_compare 風格（側邊欄 + disabled 按鈕） |
| 2026-05-05 | 改用 Groq（免費）取代 Anthropic，推送至 GitHub |
| 2026-05-05 | 修正殘留的 Anthropic 警告文字 |
| 2026-05-06 | 新增多檔合併：原始資料 uploader 支援同時上傳多個 CSV/xlsx，自動 `pd.concat` |
| 2026-05-06 | 新增批次整理 Tab：多組（原始資料＋目標格式）一次處理，多組輸出打包成 ZIP 下載 |
| 2026-05-06 | 重構：抽出 `run_conversion()`、`show_mapping_table()`、`merge_raw_files()` 共用邏輯 |
| 2026-05-06 | 批次模式表格選擇器改顯示 Word 文件真實標題（如 `indsum表3_jp`）；修正索引解析改用 `.index()` 取代字串切割 |
| 2026-05-06 | 新增 Excel 工作表選擇器：原始資料與目標格式上傳 xlsx 有多個 sheet 時，單一模式與批次模式均顯示下拉選單，修正預設讀取第一個 sheet 導致選錯工作表的問題 |
| 2026-05-06 | ③ 邏輯翻轉：改為「列出 Word 所有表格 → 使用者為每個表格指定資料來源」，並加入依表格標題自動預填來源（`_auto_match`） |
| 2026-05-06 | 允許編輯表格標題（`st.text_input`）、✕ 移除表格、＋ 加回已移除表格（`active_idxs` / `hidden_idxs` 機制） |
| 2026-05-06 | 在 ③ 每個表格加入「欄位名稱（可新增／刪除）」展開器（`st.data_editor`，預設收合），修正 delta 重複套用 bug（引入 `_hkey_orig` 作為唯一基底） |
| 2026-05-06 | 修正 Word 合併儲存格導致 DataFrame 欄數不符的 ValueError（padding / truncate 每列至 header 長度） |
| 2026-05-06 | 同步批次模式：批次 ③ 區段完整對應單一模式所有功能（表格管理、欄位編輯、資料來源選擇） |
| 2026-05-06 | 新增「＋ 新增自訂表格」：使用者可自訂名稱＋欄位清單＋資料來源，輸出時以 `doc_obj.add_table()` 附加至 Word 尾端；單一模式與批次模式均支援 |
| 2026-05-06 | 新增表格順序調整：③ 每個表格卡片加入 ↑ / ↓ 按鈕，點擊即與相鄰表格交換位置（首尾自動 disabled），輸出 Word 時依調整後順序填入；單一模式與批次模式均支援 |
| 2026-05-07 | 修正 Groq 413 Token 超限：新增 `_filter_raw_by_keywords()`，在 `get_mapping` 送 AI 前先以目標欄位名稱拆出關鍵字（≥2字元），只保留名稱包含任一關鍵字的原始資料欄位；`apply_mapping_to_df` 仍使用完整 raw_df 套用結果 |
| 2026-05-07 | 補強 413 防護：`get_mapping` 生成 `raw_desc` 後若字元數 > 2500 則退回「僅列欄位名稱」模式，`raw_sample` 超過 800 字元截斷；修正目標有 HS01~HS97 欄時關鍵字過濾無效的問題；UI 加入過濾欄數 debug caption |
| 2026-05-07 | 新增跨國家欄位匹配：`_strip_country()` 以 regex `^(.+?年).+?(自.+)$` 擷取「去國家名稱」版本（如「2015年法國自世界平均進口金額」→「2015年自世界平均進口金額」），過濾時對目標與原始欄位均比對去國家版，解決目標為他國、原始為本國的跨國欄位無法匹配問題 |
| 2026-05-07 | 新增自動國家名稱替換：`_infer_country()` 從 raw_df 欄位推斷國家，`_replace_country_in_cols()` 將 result_df 欄位與 Word 標題列的國家名稱自動替換為原始資料的國家（如目標欄位「日本」→ 原始資料「法國」→ 輸出欄位自動改為「法國」）；單一與批次模式均支援 |
| 2026-05-07 | 重寫 `_fill_word_table`：改為「原地填入」模式，保留 Word 模板的合併儲存格結構（汽車產業/上游/下游）與標籤列；新增 `_is_vmerge_continuation()` 跳過垂直合併延續格、`_set_cell_text()` 不破壞 XML 格式地更新文字、`_fmt_val()` 格式化數字（整數去 .0、小數最多 4 位）、`_replace_text_in_cell()` 安全更新標題列國家名稱 |
| 2026-05-07 | 修正 `_fmt_val()` 小數精度：改用 `str(f)` 保留原始位數，不再限制 4 位（如 `11.49507` 原樣輸出，不再四捨五入為 `11.4951`） |
