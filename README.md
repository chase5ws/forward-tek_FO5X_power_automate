![version](https://img.shields.io/badge/version-1.0-green)
![license](https://img.shields.io/badge/license-MIT%20%2B%20-blue)
![python](https://img.shields.io/badge/Python-3.8+-orange)
![pandas](https://img.shields.io/badge/Pandas-1.5.0-lightgrey)
![openpyxl](https://img.shields.io/badge/openpyxl-3.0.10-yellow)
![powerautomate](https://img.shields.io/badge/Power%20Automate-2506-blue)
![icon](asset/icon.png)

---

### 專案說明

本專案結合 **Microsoft Power Automate** 和 **Python**，實現了一個自動化的 Excel 數據處理解決方案。該系統旨在處理需求單的數據，提供高效的數據分析與報表生成功能，並兼顧穿透力與數據彈性。

- **Power Automate**：負責模擬用戶行為，穿透系統限制，進行數據提取與初步處理。
- **Python**：負責進一步的數據解析、分析與報表生成，確保數據處理的靈活性。

---

### 功能特色

#### **Power Automate**
- **穿透力強**：模擬瀏覽器操作，避免系統的反爬蟲限制。
- **圖形化操作**：易於設置與維護，適合快速部署。
- **初步處理**：提取需求單數據並生成處理單，提升效率。

#### **Python**
- **數據彈性處理**：支持自定義的數據分析與報表生成。
- **高效解析**：使用正則表達式匹配數據，轉換為結構化的 DataFrame。
- **統計分析**：按部門與需求類別生成統計報表。

---

### 相對工具穿刺比較

| 工具          | 穿透力度 | 適配性  | 易用度 | 維護難度 |
|---------------|----------|---------|--------|----------|
| **autoid**    | 高       | 低      | 高     | 最高     |
| **automate**  | 高       | 最高    | 中     | 中       |
| **puppeteer.js** | 中       | 高      | 低     | 低       |
| **手動.js**   | 中       | 高      | 高     | 低       |
| **python**    | 低       | 高      | 低     | 低       |

（資料來源於內部測試結果）

---

### 實作範疇與流程

![實作範疇](https://d41chssnpqdne.cloudfront.net/user_upload_by_module/chat_bot/files/7320191/txVCiVDJe9p3KfOa.png)

- **Power Automate 流程：**
  1. 使用瀏覽器模擬用戶行為登入系統。
  2. 提取需求單數據（填表人、部門、內容等）。
  3. 生成處理單並填寫相關資訊。

- **Python 流程：**
  1. 解析需求單數據，生成結構化表格。
  2. 進行數據統計與分析。
  3. 將結果保存為 Excel 報表。

---

### 安裝與使用指南

#### **Power Automate**
1. 確保已安裝最新版本的 **Microsoft Power Automate**。
2. 按照流程圖設置自動化操作，並配置瀏覽器模擬行為。

#### **Python**
1. 安裝必要套件：
   ```bash
   pip install pandas openpyxl
   ```
2. 修改腳本參數：
   - 更新 `file_path` 和 `output_file` 為您的檔案路徑。
   - 根據需求調整篩選天數變數 `x`。
3. 執行腳本：
   ```bash
   python execel_data_processing.py
   ```

---

### 範例數據與結果

#### **輸入範例：**
| 表單代號       | 申請日期   | 填表人 | 填表部門 | 需求類別 | 結案日期             | 簽核狀態 | 簽核結果 |
|----------------|------------|--------|----------|----------|----------------------|----------|----------|
| FO5X1234567890 | 2025/06/10 | 王小明 | 行政部   | A類需求 | 2025/06/15 10:00:00 | 已簽核   | 同意     |

#### **輸出範例：**
- **篩選後數據**：包含最近 7 天的數據，按日期降序排列。
- **統計結果**：按部門和需求類別生成統計報表。

---

### 注意事項

- 確保輸入檔案的數據格式符合正則表達式的定義。
- 若遇到解析錯誤，請檢查 Excel 文件的內容或格式。
- 使用 Power Automate 時，需考慮系統的反爬蟲限制與資安風險。

---

### 解說影片

**待續**

---

### 授權

本專案採用 [MIT License](https://opensource.org/licenses/MIT) 並結合 **Microsoft Power Automate 最新版本** 授權，歡迎自由使用與修改。