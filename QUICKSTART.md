# 🚀 快速入門指南

這是BIRT AI產生器的快速入門指南，幫助您在5分鐘內開始使用。

## 📋 前置需求

- Python 3.8 或更高版本
- 穩定的網路連線（用於存取Gemini API）
- Google帳號（用於取得Gemini API金鑰）
- Quicksilver資料庫存取權限（用於報表資料來源）

## ⚡ 30秒快速開始

### 1. 安裝相依套件

```bash
pip install -r requirements.txt
```

### 2. 配置環境設定

複製環境設定範本並編輯：

```bash
cp .env.example .env
```

編輯 `.env` 檔案，設定必要配置：

```env
# Google Gemini API配置
GEMINI_API_KEY=your_gemini_api_key_here

# Quicksilver資料庫配置
QS_DB_URL=jdbc:jeedsql:jtds:sqlserver://your-server:1433;DatabaseName=QS;
QS_DB_USER=your_database_username
QS_DB_PASSWORD=your_database_password
```

**取得API金鑰步驟：**
- 造訪 https://makersuite.google.com/app/apikey
- 登入您的Google帳號
- 點擊「Create API Key」
- 複製產生的API金鑰到 `.env` 檔案中

### 3. 驗證配置

```bash
python src/main.py --validate
```

配置成功！🎉

## 📊 處理您的Excel檔案

### 步驟1: 準備Excel檔案

將您的Excel檔案放入 `input/` 目錄：

```bash
# 建立範例檔案結構
mkdir -p input
# 將您的Excel檔案複製到input目錄
```

### 步驟2: 執行轉換

```bash
# 處理單個檔案
python src/main.py -i input/your_file.xlsx

# 批量處理所有檔案
python src/main.py -i input/
```

### 步驟3: 檢視結果

產生的BIRT報表檔案位於 `output/` 目錄：
- `*.rptdesign` - BIRT報表設計檔案
- `batch_report_*.json` - 處理報告

## 🎯 典型使用場景

### 場景1: 單元管理報表

Excel檔案包含欄位：編碼、名稱、類型
→ 自動產生SQL: `SELECT "FCode", "FName", "FTable" FROM "TsUnit"`

### 場景2: 政府服務數據報表

Excel檔案包含欄位：案號、銀行代碼、銀行名稱、交易類別
→ 自動產生帶參數的報表，支援條件篩選

### 場景3: 批量轉換

有100個Excel報表需要轉換
→ 一鍵批量處理，自動產生所有BIRT檔案

## 🔧 常用指令

```bash
# 驗證環境配置
python src/main.py --validate

# 列出要處理的檔案
python src/main.py -i input/ --list

# 處理特定模式的檔案
python src/main.py -i input/ -p "*銷售*.xlsx"

# 測試資料庫連線
python -c "from src.utils.config import Config; print('DB連線測試:', Config.get_db_url())"
```

## ❓ 遇到問題？

### 問题1: 缺少API金鑰
編輯 `.env` 檔案，設定 `GEMINI_API_KEY`

### 問题2: 模組導入錯誤
```bash
pip install -r requirements.txt
```

### 問题3: Excel檔案無法解析
檢查檔案格式是否為 `.xlsx` 或 `.xls`

### 問题4: 網路連線問題
確保能夠存取 `generativelanguage.googleapis.com`

### 問题5: 資料庫連線失敗
檢查 `.env` 檔案中的資料庫連線設定：
- `QS_DB_URL`: 資料庫伺服器位址
- `QS_DB_USER`: 使用者名稱
- `QS_DB_PASSWORD`: 密碼

## 🎉 下一步

現在您已經成功執行了BIRT AI產生器！

**建議操作：**
1. 嘗試處理您自己的Excel檔案
2. 檢視產生的`.rptdesign`檔案
3. 在Eclipse BIRT Designer中開啟報表
4. 部署報表到BIRT伺服器
5. 閱讀完整的 [README.md](README.md) 了解進階功能

**需要協助？**
- 查看 [README.md](README.md) 詳細文件
- 執行 `python src/main.py --help` 查看所有選項
- 檢查 `logs/` 目錄中的日誌檔案
- 確認 `.env` 檔案中的資料庫配置正確

享受AI驅動的報表開發體驗！✨