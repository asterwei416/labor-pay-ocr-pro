# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 專案簡介

勞作金名冊 OCR 轉換系統。使用者提供 Google Drive 連結（單一 PDF 或整個資料夾），系統透過 Gemini API 辨識手寫掃描件中的「編號」與「姓名」，輸出多分頁 Excel。

## 執行方式

```bash
# 安裝依賴（建議在 venv 內）
pip install -r requirements.txt

# 啟動 Streamlit 應用
streamlit run app.py

# 執行診斷工具（檢查 API Key 與套件是否就緒）
python diag.py
```

API Key 設定方式（擇一）：
- 本地：專案根目錄建立 `.env` 檔，寫入 `GEMINI_API_KEY=你的金鑰`
- 雲端部署：設定 Streamlit Secrets `GEMINI_API_KEY`

## 架構

```
app.py          — Streamlit UI；處理 Google Drive 下載、呼叫 processor
processor.py    — 核心邏輯；API 初始化、模型選擇、OCR、CSV 解析、重複偵測、Excel 輸出
diag.py         — 本地診斷腳本，不影響生產流程
```

**資料流：**
`app.py` 下載 PDF → `process_labor_pay_pdf()` 上傳至 Gemini Files API → Dry-run 等待 → 取得 OCR 文字 → 解析 CSV → 重複編號偵測 → 寫入 Excel 各分頁

## 重要設計決策

### 模型選擇
`processor.py` 在執行時動態從 `genai.list_models()` 挑選優先度最高的可用模型（priority_list 排序），不寫死版本。若需更換優先模型，修改 `priority_list`。

### OCR Prompt 設計原則
Prompt 包含兩個關鍵區塊，修改時需一起維護：
1. **數字字形混淆對照表**：列出草寫手稿中高頻混淆字形（`9↔7`、`8↔1` 等）
2. **編號位數強制規則**：所有編號**必為四位數**，AI 輸出不足四位時須補位並標記疑慮

### 重複偵測邏輯（processor.py）
- **完全重複**：`duplicated(keep=False)` 標記所有相同編號
- **近似重複**：編號等長且只差一個字元 + 姓名有共同漢字 → 標記為混淆疑慮
  （用於抓出草寫導致的單字母錯認，如 `9007` vs `7007`）

### Excel Sheet 命名
Sheet name 最長 31 字元，特殊字元 `\ / * ? [ ] :` 會被 `re.sub` 清除。

### 對帳監控
Prompt 要求 AI 在 CSV 區塊後輸出 `TOTAL_ROWS: N`，系統會比對解析筆數與 AI 自報數字，不符時在 audit log 標記差異筆數。
