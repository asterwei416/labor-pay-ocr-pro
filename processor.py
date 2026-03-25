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
        'models/gemini-2.5-pro-preview-03-25',
        'models/gemini-2.0-flash',
        'models/gemini-2.0-pro-exp-02-05',
        'models/gemini-2.0-flash-exp',
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
    這是一份「勞作金名冊」手寫掃描件。請只擷取「編號」與「姓名」兩欄。

    【數字字形混淆對照表——辨識前必讀】
    - `9`：草寫收尾若下鉤不清晰，極易被誤判為 `7`；若圈部封閉則可能被誤判為 `0`。
      ⚠️ 遇到開頭為 `7` 的編號，必須回頭確認字形上方是否有封閉圓弧（即實為 `9`）。
    - `8`：常寫成一豎線旁帶封閉或半封閉圈套，極易被誤判為 `1`、`5` 或 `0`
    - `5`：草寫時上半部可能退化成一橫，形似 `1` 或 `7`
    - `0`：可能偏扁形似 `6` 或 `9`
    - `1`：若末端帶鉤，注意是否實為 `7`
    - 編號末位常有書寫者自帶「勾筆」，此為書寫習慣，**不代表多一個數字**
    - 看到豎線＋圈套結構，優先假設為 `8`，再逐步排除

    【姓名辨識原則】
    - 依筆畫結構＋常見百家姓交叉推敲草寫字
    - 連筆字請拆解部首後比對常見漢字

    【編號位數強制規則——最高優先】
    - 本名冊所有編號**必為四位數**（例如 5007、9844、7210）。
    - 若辨識結果不足四位（如只讀到 500、981），代表有數字被遺漏，必須重新審視該格，補上缺失數字。
    - 補上後若仍不確定，在 AI_辨識疑慮填寫「是(數字疑慮：編號疑似缺位，補填為 XXXX)」。
    - 禁止輸出三位或五位以上的編號；若真的完全無法確認第四位，填 `?` 補位（如 500?）。

    【任務步驟】
    1. 先在 `<thinking>` 標籤內，逐列針對「編號」數字做交叉比對推論，記錄任何疑慮。
       特別確認每個編號是否為四位數，不足者重新檢查。
    2. 輸出 CSV，標題行必須為：「編號,姓名,AI_辨識疑慮」
       - CSV 必須包在 ```csv ``` 區塊內。
    3. 禁止漏行：字跡模糊也必須輸出，完全無法辨識填 `?`。
    4. AI_辨識疑慮填寫規則：
       - 編號任一數字非 100% 清晰 → 填「是(數字疑慮：說明)」
       - 編號補位不確定 → 填「是(數字疑慮：編號疑似缺位，補填為 XXXX)」
       - 姓名有連筆難辨 → 填「是(姓名連筆：說明)」
       - 全部清晰確定 → 填「正常」
       - 完全失敗 → 填「?」
    5. CSV 區塊後輸出：`TOTAL_ROWS: [數字]`
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

    def log_status(msg):
        if status_container:
            try:
                # 兼容 streamlit status 物件與一般文字容器
                status_container.write(msg)
            except Exception:
                st.write(msg)
        else:
            print(msg)

    for idx, pdf_file in enumerate(pdf_files):
        # sheet_name 最長31字元，且不能包含特殊字元
        base_name = os.path.basename(pdf_file).replace('.pdf', '')
        sheet_name = re.sub(r'[\\/\*\?\[\]:]', '', base_name)[:31]
        
        log_status(f"⏳ 正在處理第 {idx + 1}/{total_files} 個檔案：`{base_name}` ...")
            
        uploaded_file = None
        try:
            uploaded_file = genai.upload_file(pdf_file)
            log_status(f"↖️ `{base_name}` 已上傳至 AI 模型，正在進行高精準度 OCR 辨識...")

            while uploaded_file.state.name == "PROCESSING":
                time.sleep(2)
                uploaded_file = genai.get_file(uploaded_file.name)

            if uploaded_file.state.name == "FAILED":
                raise Exception("AI 模型處理檔案失敗")

            response = model.generate_content([uploaded_file, prompt])
            log_status(f"🧠 `{base_name}` AI 辨識完成，正在解析數據...")
            full_text = response.text.strip()

            # 解析 CSV
            csv_part = full_text
            expected_count = 0

            m = re.search(r"TOTAL_ROWS:\s*(\d+)", full_text, re.IGNORECASE)
            if m:
                expected_count = int(m.group(1))

            if "```" in full_text:
                csv_blocks = re.findall(r"```(?:csv)?(.*?)```", full_text, re.DOTALL)
                if csv_blocks:
                    csv_part = csv_blocks[-1].strip()

            parsed_rows = []
            raw_lines = csv_part.strip().split("\n")
            header = None
            for line in raw_lines:
                if not line.strip():
                    continue
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

            # 欄位名稱
            id_col = next((c for c in df.columns if '編號' in str(c) or '號' in str(c)), None)
            name_col = next((c for c in df.columns if '姓名' in str(c) or '名' in str(c)), None)

            # 重置為正常，只標記真正需要人工確認的情況
            df['AI_辨識疑慮'] = '正常'

            if id_col:
                valid_mask = df[id_col].notna() & (~df[id_col].astype(str).str.strip().isin(['?', '??', 'nan', '']))

                # 完全重複編號
                dup_mask = pd.Series(False, index=df.index)
                dup_mask[valid_mask] = df.loc[valid_mask, id_col].duplicated(keep=False)
                if dup_mask.any():
                    df.loc[dup_mask, 'AI_辨識疑慮'] = '⚠️確認重複編號'

                # 近似重複：編號差 1 碼 + 姓名有共同字
                valid_ids = df.loc[valid_mask, id_col].astype(str).str.strip().tolist()
                valid_names = df.loc[valid_mask, name_col].astype(str).str.strip().tolist() if name_col else [''] * len(valid_ids)
                valid_indices = df.loc[valid_mask].index.tolist()

                near_dup_flags = {}
                for i in range(len(valid_ids)):
                    for j in range(i + 1, len(valid_ids)):
                        a, b = valid_ids[i], valid_ids[j]
                        if len(a) != len(b) or sum(x != y for x, y in zip(a, b)) != 1:
                            continue
                        name_a, name_b = valid_names[i], valid_names[j]
                        if not (set(name_a) & set(name_b) - {'?', ''}):
                            continue
                        near_dup_flags.setdefault(valid_indices[i], []).append(f"{b}({name_b})")
                        near_dup_flags.setdefault(valid_indices[j], []).append(f"{a}({name_a})")

                for idx_val, similar_list in near_dup_flags.items():
                    df.at[idx_val, 'AI_辨識疑慮'] = '⚠️近似重複:疑與' + '、'.join(similar_list) + '混淆'

            all_dfs[sheet_name] = df

        except Exception as e:
            audit_logs.append(f"【{sheet_name}】: ❌ 處理發生錯誤 - {e}")
        finally:
            if uploaded_file:
                try:
                    genai.delete_file(uploaded_file.name)
                except Exception:
                    pass

    log_status("✅ 所有檔案皆已通過 AI 辨識與資料解析！準備產出報表...")
    time.sleep(1)

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
