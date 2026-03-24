import os
import time
import pandas as pd
import re
import glob
import streamlit as st

try:
    import google.generativeai as genai
except ImportError:
    st.error("❌ 請先安裝套件: pip install google-generativeai pandas openpyxl")

# 嘗試讀取同目錄下的 .env 檔案以確保向下相容
def load_env_file():
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                if "=" in line and not line.startswith("#"):
                    k, v = line.strip().split("=", 1)
                    clean_v = v.strip().strip("'").strip('"').strip()
                    os.environ[k.strip()] = clean_v

def get_api_key():
    # 優先嘗試從 Streamlit secrets 取得 (雲端部署用)
    try:
        if st.secrets and "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass
    
    # 若無，則嘗試從環境變數取得 (包括 .env)
    return os.getenv("GEMINI_API_KEY")

def process_labor_pay_pdf(target_path: str, output_excel_path: str, status_container=None) -> tuple[bool, str]:
    """
    處理單一 PDF 或整個資料夾的 PDF，並輸出為多個分頁的 Excel。
    """
    load_env_file() # 自動載入本地 API KEY
    
    api_key = get_api_key()
    if not api_key or not api_key.strip():
        return False, "未偵測到 API Key。如果在本地執行，請確認專案資料夾有 .env 檔案或設定好 Streamlit Secrets。"

    genai.configure(api_key=api_key.strip())
    
    # 動態優選最強模型
    try:
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    except Exception as e:
        return False, f"API Key 驗證失敗: {e}"

    priority_list = [
        'models/gemini-2.5-pro', 
        'models/gemini-2.0-pro', 
        'models/gemini-1.5-pro-latest',
        'models/gemini-1.5-pro'
    ]
    
    model_name = 'gemini-1.5-pro'
    for p in priority_list:
        if p in available_models:
            model_name = p.replace('models/', '')
            break

    model = genai.GenerativeModel(model_name)
    
    prompt = """
    這是一份「勞作金名冊」手寫掃描件。請執行以下任務：
    1. 【🧐 嚴謹思維鏈分析】在輸出資料前，請仔細觀察圖片中表格的每一列。特別注意：
       - 編號多為數字，這是系統的唯一識別碼，**極度重要，辨識錯誤是非常嚴重的問題**！手寫連筆極易誤判（例如：數字 8 常常被寫得像 1 加上一個圈，或者像 5、0；數字 5 容易寫得像 1 ；編號最後一碼經常帶有勾筆或變形）。
       - 當你看到一根豎線旁邊帶有封閉或半封閉的圈套，請高度懷疑它是 `8` 而不是 `1` 或 `5`。
       - 姓名為中文，請根據筆畫結構、前後文與常見百家姓推敲草寫字。
       - 請先在 `<thinking>` 標籤內，針對「編號」進行特別的交叉比對與推論，若有任何一絲疑慮，必須記錄下來。
    2. 【📋 輸出 CSV 資料】標題行必須為：「編號,姓名,AI_辨識疑慮」。
       - 請務必將 CSV 資料段落獨立包裝在 ```csv 與 ``` 區塊內。
    3. 【🚨 嚴格命令：禁止漏行】即便字跡模糊也必須輸出一行，完全看不懂的字元請填 `?`，切勿整行跳過。
    4. 【⚠️ AI_辨識疑慮填寫規則】
       - **最高警報：只要「編號」的任何一個數字不是 100% 清晰，有連筆、塗改或是長得像 A 也像 B 的情況，必須強制在該欄位填寫「是(數字疑慮：可能為 X 或 Y)」。**
       - 姓名若有連筆難辨，請填寫「是(姓名連筆：說明原因)」。
       - 只有當字跡完全清晰確切，且 100% 肯定沒有連筆混淆可能時，才可填寫「否」。
       - 完全解譯失敗填「?」。
    5. 【📊 總計標註】在 CSV 區塊之後，請務必加上獨立一行標註：`TOTAL_ROWS: [數字]`，代表你辨識到的資料總筆數(不含標題)。
    """

    # 判斷 target_path 是檔案還是目錄，準備要處理的檔案清單
    pdf_files = []
    if os.path.isdir(target_path):
        pdf_files = glob.glob(os.path.join(target_path, "**/*.pdf"), recursive=True)
    elif os.path.isfile(target_path) and target_path.lower().endswith(".pdf"):
        pdf_files = [target_path]
        
    if not pdf_files:
        return False, "在下載的內容中找不到任何 PDF 檔案可以處理。"

    all_dfs = {}
    audit_logs = []
    
    total_files = len(pdf_files)

    if status_container:
        progress_bar = status_container.progress(0)
    else:
        progress_text = st.empty()
        progress_bar = st.empty()

    def log_status(msg):
        if status_container:
            status_container.write(msg)
        else:
            progress_text.info(msg)

    for idx, pdf_file in enumerate(pdf_files):
        # sheet_name 最長31字元，且不能包含特殊字元
        base_name = os.path.basename(pdf_file).replace('.pdf', '')
        sheet_name = re.sub(r'[\\/\*\?\[\]:]', '', base_name)[:31]
        
        log_status(f"⏳ 正在處理第 {idx + 1}/{total_files} 個檔案：`{base_name}` ...")
        if total_files > 0:
            progress_bar.progress(idx / total_files)
            
        uploaded_file = None
        try:
            uploaded_file = genai.upload_file(pdf_file)
            log_status(f"↖️ `{base_name}` 已上傳至 AI 模型，正在進行高精準度 OCR 辨識，此步驟需要較長時間...")
            time.sleep(5)  # 等待文檔處理完畢
            
            response = model.generate_content([uploaded_file, prompt])
            log_status(f"🧠 `{base_name}` AI 辨識完成，正在解析並轉換資料結構...")
            full_text = response.text.strip()
            
            # 分離 CSV 與 對帳統計
            csv_part = full_text
            expected_count = 0
            
            # Regex 抓取 TOTAL_ROWS: [數字]
            match = re.search(r"TOTAL_ROWS:\s*(\d+)", full_text, re.IGNORECASE)
            if match:
                expected_count = int(match.group(1))
            
            if "```" in full_text:
                csv_blocks = re.findall(r"```(?:csv)?(.*?)```", full_text, re.DOTALL)
                if csv_blocks: csv_part = csv_blocks[-1].strip()  # 抓最後一個區塊，防止 thinking block 也包裝在 markdown block 中

            # 解析 CSV (手動逐行解析確保防呆)
            parsed_rows = []
            raw_lines = csv_part.strip().split("\n")
            header = None
            for line in raw_lines:
                fields = [f.strip() for f in line.split(",")]
                if header is None:
                    header = fields[:3]
                    continue
                padded = (fields[:3] + ['', '', ''])[:3]
                parsed_rows.append(padded)
                
            df = pd.DataFrame(parsed_rows, columns=header if header else ['編號', '姓名', 'AI_辨識疑慮'])
            actual_count = len(df)
            
            # 對帳監控
            if expected_count > 0:
                is_match = (actual_count == expected_count)
                log_icon = "✅" if is_match else "⚠️"
                diff_count = abs(expected_count - actual_count)
                log_detail = "100%吻合" if is_match else f"筆數不符，差 {diff_count} 筆"
                audit_logs.append(f"【{sheet_name}】: 預期 {expected_count} / 實得 {actual_count} ({log_icon} {log_detail})")
            else:
                audit_logs.append(f"【{sheet_name}】: 解析筆數 {actual_count} 筆 (未取得 AI 預期統計)")

            # 檢核邏輯
            def check_reliability(row):
                warnings = []
                target_cols = [c for c in df.columns if c != 'AI_辨識疑慮']
                id_col = next((c for c in target_cols if '編號' in str(c) or '號' in str(c)), None)
                name_col = next((c for c in target_cols if '姓名' in str(c) or '名' in str(c)), None)
                
                if not id_col or pd.isna(row.get(id_col)) or str(row.get(id_col)).strip() in ['?', 'nan', '']:
                    warnings.append("🚨嚴重:編號遺失")
                elif not str(row[id_col]).strip().isdigit():
                    warnings.append("⚠️異狀:編號非數字")
                        
                if not name_col or pd.isna(row.get(name_col)) or str(row.get(name_col)).strip() in ['?', 'nan', '']:
                    warnings.append("🚨嚴重:姓名遺失")
                
                ai_note = str(row.get('AI_辨識疑慮', '')).strip()
                if ai_note and ai_note not in ['nan', '否', '正常', '?', 'None', 'False', 'True', '0', '1']:
                    warnings.append(f"🤖AI註記:{ai_note}")
                
                return " | ".join(warnings) if warnings else "正常"
                
            df['AI_辨識疑慮'] = df.apply(check_reliability, axis=1)
            
            # 重複檢測
            id_col = next((c for c in df.columns if '編號' in str(c) or '號' in str(c)), None)
            if id_col:
                valid_mask = df[id_col].notna() & (~df[id_col].astype(str).str.strip().isin(['?', '??', 'nan', '']))
                dup_mask = pd.Series(False, index=df.index)
                dup_mask[valid_mask] = df.loc[valid_mask, id_col].duplicated(keep=False)
                if dup_mask.any():
                    df.loc[dup_mask, 'AI_辨識疑慮'] = df.loc[dup_mask, 'AI_辨識疑慮'].astype(str) + " | ⚠️重複編號"
                    
            all_dfs[sheet_name] = df

        except Exception as e:
            audit_logs.append(f"【{sheet_name}】: ❌ 處理發生錯誤 - {e}")
        finally:
            if uploaded_file: 
                try:
                    genai.delete_file(uploaded_file.name)
                except Exception:
                    pass
                    
    if total_files > 0:
        progress_bar.progress(1.0)
        
    log_status("✅ 所有檔案皆已通過 AI 辨識與資料解析！準備產出報表...")
    time.sleep(1)
    
    if not status_container:
        progress_text.empty()
        progress_bar.empty()

    # 所有檔案處理完畢，合併打包成單一 Excel (多個 Sheet 分頁)
    if all_dfs:
        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                for sn, d in all_dfs.items():
                    d.to_excel(writer, sheet_name=sn, index=False)
                    
            global_audit_msg = "\n".join(audit_logs)
            return True, global_audit_msg
        except Exception as e:
            return False, f"Excel 打包失敗: {e}"
    else:
        return False, "所有 PDF 皆處理失敗，未生成 Excel。"
