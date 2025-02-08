# Excel Data Cleaner

## 簡介
本專案是一個簡單的 GUI 工具，讓使用者可以選擇資料夾並清空所有 Excel (`.xlsx` 或 `.xls`) 檔案內的數據，只保留標題列。此工具適合需要大量清理 Excel 資料的使用者。

## 前置作業
### 1. 安裝 Python
請確保已安裝 Python 3.7 以上版本，若尚未安裝，請至 [Python 官方網站](https://www.python.org/downloads/) 下載並安裝。

### 2. 安裝必要套件
請在終端機或命令提示字元中執行以下指令，以安裝所需的 Python 套件：

```bash
pip install pandas openpyxl tk
```

1. 點擊「選擇資料夾並清空Excel」，選擇要處理的 Excel 檔案資料夾。
2. 系統會彈出確認訊息，請再次確認是否要清空數據。
3. 確認後，所有 Excel 檔案的數據將被清空，僅保留標題列。
4. 若資料夾內無 Excel 檔案，系統會提示警告。
5. 完成後會顯示提示訊息。

## 目錄結構
```
excel-data-cleaner/
│── clear_excel_data.py  # 主程式
│── README.md            # 使用說明文件
```