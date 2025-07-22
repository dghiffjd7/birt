# BIRT AI報表產生器

基於AI技術的BIRT報表自動產生工具，能夠分析Excel檔案並自動產生對應的BIRT報表檔案。

## 功能特性

- **Excel智能分析**: 自動解析Excel檔案結構、資料類型、圖表等
- **AI驅動生成**: 使用Google Gemini 2.5 pro模型智能生成BIRT配置
- **批量處理**: 支援批量處理多個Excel檔案
- **Quicksilver整合**: 針對Quicksilver系統最佳化，支援JeedSQL驅動
- **模板化生成**: 基於模板系統，支援多種報表類型
- **互動式配置**: 執行時提示使用者輸入API金鑰，無需預先配置

## 快速開始

### 1. 安裝相依套件

```bash
pip install -r requirements.txt
```

### 2. 配置Gemini API與資料庫連線

執行配置助手：

```bash
python setup_gemini.py
```

或者手動配置，複製 `.env.example` 為 `.env` 並編輯：

```bash
cp .env.example .env
```

編輯 `.env` 檔案，設定必要的配置項目：

```env
# Google Gemini配置
GEMINI_API_KEY=your_gemini_api_key_here

# Quicksilver資料庫配置  
QS_DB_DRIVER=Your_Driver
QS_DB_URL=Your_sqlserver://localhost:;DatabaseName=QS;
QS_DB_USER=your_username
QS_DB_PASSWORD=your_password

# 額外資料庫設定（可選）
QS_CONNECTION_TIMEOUT=30
QS_QUERY_TIMEOUT=60
```

獲取Gemini API金鑰請造訪: https://makersuite.google.com/app/apikey

### 3. 準備Excel檔案

將需要轉換的Excel檔案放入 `input/` 目錄。

### 4. 執行產生器

```bash
# 首次執行會提示輸入Gemini API金鑰
python src/main.py -i input/sample.xlsx

# 批量處理整個目錄
python src/main.py -i input/ -o output/

# 驗證環境配置
python src/main.py --validate

# 執行互動式範例
python run_example.py
```

## 專案結構

```
birt/
├── src/                          # 原始碼目錄
│   ├── analyzers/               # Excel分析器
│   │   └── excel_analyzer.py    # Excel檔案分析
│   ├── generators/              # BIRT產生器
│   │   └── birt_generator.py    # BIRT報表產生
│   ├── utils/                   # 工具模組
│   │   ├── gemini_analyzer.py   # Gemini AI分析器
│   │   ├── config.py           # 配置管理
│   │   └── logger.py           # 日誌配置
│   └── main.py                 # 主程式入口
├── templates/                   # 模板目錄
│   └── birt/                   # BIRT模板
│       └── base_template.xml   # 基礎模板
├── input/                      # 輸入目錄（Excel檔案）
├── output/                     # 輸出目錄（BIRT檔案）
├── logs/                       # 日誌目錄
├── requirements.txt            # Python相依套件
├── .env.example               # 環境配置範例
└── README.md                  # 說明文件
```

## 使用方法

### 命令列參數

```bash
python src/main.py [選項]

選項:
  -i, --input     輸入Excel檔案或目錄 (必需)
  -o, --output    輸出目錄 (預設: output)
  -p, --pattern   檔案匹配模式 (預設: *.xlsx)
  -c, --config    配置檔案路徑 (預設: .env)
  --validate      驗證環境配置
  --list          只列出匹配的檔案，不處理
```

### 範例

```bash
# 處理單個Excel檔案
python src/main.py -i input/sales_report.xlsx

# 批量處理所有xlsx檔案
python src/main.py -i input/ -o output/

# 處理特定模式的檔案
python src/main.py -i input/ -p "*銷售*.xlsx"

# 驗證配置是否正確
python src/main.py --validate

# 列出要處理的檔案
python src/main.py -i input/ --list
```

## 支援的Excel格式

- **.xlsx** - Excel 2007及以上格式
- **.xls** - Excel 97-2003格式

## 產生的報表類型

根據Excel內容自動選擇合適的報表類型：

- **simple_listing**: 簡單清單報表
- **detailed_listing**: 詳細清單報表  
- **summary_report**: 彙總報表
- **dashboard**: 儀表板報表（包含圖表）

## 配置說明

### 資料庫配置

針對Quicksilver系統的資料庫連線配置，請編輯 `.env` 檔案：

```env
# Quicksilver資料庫連線設定
QS_DB_DRIVER=com.jeedsoft.jeedsql.jdbc.Driver
QS_DB_URL=jdbc:jeedsql:jtds:sqlserver://your-server:1433;DatabaseName=QS;
QS_DB_USER=your_database_username
QS_DB_PASSWORD=your_database_password

# 連線逾時設定（可選）
QS_CONNECTION_TIMEOUT=30
QS_QUERY_TIMEOUT=60

# 連線池設定（可選）
QS_MAX_POOL_SIZE=10
QS_MIN_POOL_SIZE=2
```

**資料庫配置說明：**
- `QS_DB_URL`: 替換 `your-server` 為實際的伺服器位址
- `QS_DB_USER`: 資料庫使用者名稱
- `QS_DB_PASSWORD`: 資料庫密碼
- 確保資料庫使用者具有適當的讀取權限

### AI配置

Google Gemini API配置，用於智慧分析：

```env
GEMINI_API_KEY=your_gemini_api_key
```

**取得API金鑰：**
1. 造訪 https://makersuite.google.com/app/apikey
2. 登入Google帳號
3. 點擊 "Create API Key"
4. 選擇專案或建立新專案
5. 複製產生的API金鑰

## 輸出檔案

產生的檔案包括：

- **\*.rptdesign**: BIRT報表設計檔案
- **batch_report_\*.json**: 批量處理報告
- **logs/birt_generator_\*.log**: 處理日誌

## 故障排除

### 常見問題

1. **ImportError**: 確保已安裝所有相依套件 `pip install -r requirements.txt`
2. **API Key錯誤**: 執行 `python setup_gemini.py` 重新配置API金鑰
3. **檔案權限錯誤**: 確保對輸出目錄有寫入權限
4. **Excel解析失敗**: 檢查Excel檔案是否損壞或格式不支援
5. **網路連線問題**: 確保能存取Google API服務
6. **資料庫連線失敗**: 檢查`.env`中的資料庫配置設定

### 環境驗證

執行環境驗證命令：

```bash
python src/main.py --validate
```

### 日誌查看

查看詳細的處理日誌：

```bash
tail -f logs/birt_generator_$(date +%Y%m%d).log
```

## 開發指南

### 新增報表模板

1. 在 `templates/birt/` 目錄下建立新模板檔案
2. 在 `BIRTGenerator._select_template()` 方法中新增模板對應
3. 在 `ExcelAnalyzer._suggest_report_type()` 中新增類型判斷邏輯

### 擴展AI分析功能

1. 修改 `gemini_analyzer.py` 中的提示詞模板
2. 在 `_parse_response()` 方法中處理新的AI輸出格式
3. 更新 `AIAnalysisResult` 資料結構
4. 調整Gemini模型參數以優化產生品質

## 許可證

本專案僅供內部使用，請勿外傳。

## 聯絡支援

如有問題請聯絡開發人員。