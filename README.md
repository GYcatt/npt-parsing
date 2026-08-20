# NPT Parsing — Windows 使用說明

把 NPT 量測 CSV 自動填進 `data analysis.xlsx` 的三張 4G 表（**LTE MAX** / **LTE AVG** / **LTE Mid AVG**）。程式與範例檔都在 `app\`，平常不用進去。

## 第一次使用（只需做一次）

1. **雙擊 `setup.bat`** — 自動安裝執行環境（約 1～5 分鐘，需能連網）
2. 看到「環境建置完成」後關閉視窗

新電腦**不必先裝 Python**。`setup.bat` 會優先用電腦上已有的 Python；沒有就下載免安裝版到 `app\runtime\`。

## 日常使用

1. 把 CSV 放到 `NPT result\<專案名稱>\`（必須在該資料夾**第一層**）
2. **關閉** `data analysis.xlsx`（開著會寫不進去）
3. **雙擊 `NPT Parsing.bat`**
4. 打開 Excel 查看 `LTE MAX`、`LTE AVG`、`LTE Mid AVG`

結果檔不見了也沒關係：程式會自動從 `app\data analysis_temp.xlsx` 複製一份新的再填。已存在的結果檔不會被蓋掉。

---

## 常用操作

| 想做什麼 | 執行哪個檔案 |
|----------|----------------|
| 第一次安裝環境 / 新電腦 | `setup.bat` |
| 平常匯入（新專案加欄、同名則更新數字） | `NPT Parsing.bat` |
| 全量重排（清空所有專案欄後重填） | `NPT Parsing_rebuild.bat` |
| 從零重建結果檔 | 刪掉 `data analysis.xlsx`，再跑 `NPT Parsing.bat` |
| 改預設表格（Band / Spec / 頻寬） | 編輯 `app\data analysis_temp.xlsx` 的三張 4G 表 |
| 套件壞了、重裝環境 | 刪除 `app\venv` 與 `app\runtime`，再執行 `setup.bat` |

更完整的步驟、檔名規則、顏色與 FAQ 請看 **`使用說明.txt`**（可直接用記事本打開）。

---

## 資料夾結構

```
NTP toolfor nick/
├── setup.bat                   環境建置
├── NPT Parsing.bat             平常匯入
├── NPT Parsing_rebuild.bat     全量重排
├── data analysis.xlsx          結果表（缺檔會自動從範本重建）
├── NPT result/<專案>/*.csv     量測資料
├── 使用說明.txt / README.md
└── app/
    ├── NPT Parsing.py
    ├── setup.ps1
    ├── requirements.txt
    ├── data analysis_temp.xlsx 空白範本
    └── venv/ 或 runtime/       setup 自動產生
```

---

## 常見問題

| 狀況 | 處理方式 |
|------|----------|
| 提示「找不到執行環境」 | 先雙擊 `setup.bat` |
| 提示「未偵測到 Python」或下載失敗 | 檢查網路；或自行安裝 [Python 3](https://www.python.org/downloads/)，勾選 **Add python.exe to PATH** 後重跑 `setup.bat` |
| 無法寫入 data analysis.xlsx | 先關閉 Excel 再跑一次 |
| 跳過（無符合 CSV） | CSV 放在專案資料夾第一層，檔名需含 `Band1_10MHz` 這種 Band 與頻寬 |
| 安裝套件失敗 | 檢查網路；刪除 `app\venv`、`app\runtime` 後重跑 `setup.bat` |

---

## 需要協助時

請將 `setup.bat` 或 `NPT Parsing.bat` 視窗中的**完整錯誤文字**截圖或複製給技術人員。
