import os
import time
import tempfile
import shutil
import pandas as pd
import re
import glob
import streamlit as st

try:
    import google.generativeai as genai
except ImportError:
    st.error("❌ 請先安裝套件: pip install google-generativeai pandas openpyxl")

try:
    import fitz  # pymupdf
    from PIL import Image, ImageEnhance
    _ENHANCE_AVAILABLE = True
except ImportError:
    _ENHANCE_AVAILABLE = False

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

def pdf_to_enhanced_images(pdf_path: str, out_dir: str) -> list[str]:
    """將 PDF 每頁轉為增強對比的 PNG，回傳圖片路徑清單。"""
    if not _ENHANCE_AVAILABLE:
        return []
    doc = fitz.open(pdf_path)
    paths = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img = img.convert("L")
        img = ImageEnhance.Contrast(img).enhance(2.0)
        img = ImageEnhance.Sharpness(img).enhance(2.0)
        out_path = os.path.join(out_dir, f"page_{i}.png")
        img.save(out_path, "PNG")
        paths.append(out_path)
    doc.close()
    return paths


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
    - `8`：常寫成一豎線旁帶封閉或半封閉圈套，極易被誤判為 `1`、`5` 或 `0`
    - `5`：草寫時上半部可能退化成一橫，形似 `1` 或 `7`
    - `0`：可能偏扁形似 `6` 或 `9`
    - `1`：若末端帶鉤，注意是否實為 `7`
    - 編號末位常有書寫者自帶「勾筆」，此為書寫習慣，**不代表多一個數字**
    - 看到豎線＋圈套結構，優先假設為 `8`，再逐步排除

    【姓名辨識原則】
    - 依筆畫結構＋常見百家姓交叉推敲草寫字
    - 連筆字請拆解部首後比對常見漢字

    【任務步驟】
    1. 先在 `<thinking>` 標籤內，逐列針對「編號」數字做交叉比對推論，記錄任何疑慮。
    2. 輸出 CSV，標題行必須為：「編號,姓名,AI_辨識疑慮」
       - CSV 必須包在 ```csv ``` 區塊內。
    3. 禁止漏行：字跡模糊也必須輸出，完全無法辨識填 `?`。
    4. AI_辨識疑慮填寫規則：
       - 編號任一數字非 100% 清晰 → 填「是(數字疑慮：說明)」
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
            except:
                st.write(msg)
        else:
            print(msg)

    for idx, pdf_file in enumerate(pdf_files):
        # sheet_name 最長31字元，且不能包含特殊字元
        base_name = os.path.basename(pdf_file).replace('.pdf', '')
        sheet_name = re.sub(r'[\\/\*\?\[\]:]', '', base_name)[:31]
        
        log_status(f"⏳ 正在處理第 {idx + 1}/{total_files} 個檔案：`{base_name}` ...")
            
        uploaded_files = []
        img_tmp_dir = None
        try:
            # ── 影像增強 ──────────────────────────────────────────
            if _ENHANCE_AVAILABLE:
                img_tmp_dir = tempfile.mkdtemp()
                image_paths = pdf_to_enhanced_images(pdf_file, img_tmp_dir)
                log_status(f"🖼️ `{base_name}` 影像增強完成（{len(image_paths)} 頁），正在上傳...")
                for img_path in image_paths:
                    uf = genai.upload_file(img_path)
                    while uf.state.name == "PROCESSING":
                        time.sleep(2)
                        uf = genai.get_file(uf.name)
                    if uf.state.name == "FAILED":
                        raise Exception("AI 模型處理影像失敗")
                    uploaded_files.append(uf)
            else:
                # 無增強套件，fallback 直接上傳原始 PDF
                uf = genai.upload_file(pdf_file)
                while uf.state.name == "PROCESSING":
                    time.sleep(2)
                    uf = genai.get_file(uf.name)
                if uf.state.name == "FAILED":
                    raise Exception("AI 模型處理檔案失敗")
                uploaded_files.append(uf)
                log_status(f"↖️ `{base_name}` 已上傳（未套用影像增強），正在辨識...")

            # ── 第一階段 OCR ──────────────────────────────────────
            log_status(f"🧠 `{base_name}` 第一階段 OCR 辨識中...")
            response = model.generate_content([*uploaded_files, prompt])
            full_text = response.text.strip()

            # 解析 CSV
            def parse_csv_from_text(text):
                csv_part = text
                if "```" in text:
                    blocks = re.findall(r"```(?:csv)?(.*?)```", text, re.DOTALL)
                    if blocks:
                        csv_part = blocks[-1].strip()
                rows, header = [], None
                for line in csv_part.strip().split("\n"):
                    if not line.strip():
                        continue
                    fields = [f.strip() for f in line.split(",")]
                    if header is None:
                        header = fields[:3]
                        continue
                    rows.append((fields[:3] + ['', '', ''])[:3])
                return header, rows

            header, parsed_rows = parse_csv_from_text(full_text)
            df = pd.DataFrame(parsed_rows, columns=header if header else ['編號', '姓名', 'AI_辨識疑慮'])
            actual_count = len(df)

            expected_count = 0
            m = re.search(r"TOTAL_ROWS:\s*(\d+)", full_text, re.IGNORECASE)
            if m:
                expected_count = int(m.group(1))

            # ── 第二階段驗證（疑慮行重送）────────────────────────
            suspicious_mask = df['AI_辨識疑慮'].str.contains('是', na=False)
            if suspicious_mask.any():
                suspicious_count = suspicious_mask.sum()
                log_status(f"🔍 `{base_name}` 發現 {suspicious_count} 行有疑慮，進行第二階段驗證...")

                lines_desc = []
                for pos, (_, row) in enumerate(df[suspicious_mask].iterrows(), 1):
                    lines_desc.append(f"行{pos}: 編號={row['編號']}, 姓名={row['姓名']}, 疑慮={row['AI_辨識疑慮']}")
                suspicious_str = "\n".join(lines_desc)

                stage2_prompt = f"""
這是同一份「勞作金名冊」掃描件。第一次辨識後，以下行有疑慮，請重新放大仔細辨識：

{suspicious_str}

【重點：字形混淆對照表】
- `8`：豎線旁帶封閉圈套，極易被誤判為 `1`、`5` 或 `0`
- `5`：草寫形似 `1` 或 `7`
- `0`：偏扁形似 `6` 或 `9`
- `1`：末端帶鉤注意是否為 `7`

只輸出這些行的修正結果，標題行：「編號,姓名,AI_辨識疑慮」
請將 CSV 包在 ```csv ``` 區塊內。
"""
                s2_response = model.generate_content([*uploaded_files, stage2_prompt])
                _, s2_rows = parse_csv_from_text(s2_response.text.strip())

                if s2_rows:
                    s2_df = pd.DataFrame(s2_rows, columns=['編號', '姓名', 'AI_辨識疑慮'])
                    for _, s2_row in s2_df.iterrows():
                        # 先嘗試以編號比對，找不到則以姓名比對
                        idx_match = df.index[df['編號'] == s2_row['編號']].tolist()
                        if not idx_match:
                            idx_match = df.index[df['姓名'] == s2_row['姓名']].tolist()
                        if idx_match:
                            i = idx_match[0]
                            df.at[i, '編號'] = s2_row['編號']
                            df.at[i, '姓名'] = s2_row['姓名']
                            df.at[i, 'AI_辨識疑慮'] = s2_row['AI_辨識疑慮'] + " [二次驗證]"

                log_status(f"✅ `{base_name}` 二次驗證完成。")

            # ── 對帳監控 ─────────────────────────────────────────
            if expected_count > 0:
                is_match = (actual_count == expected_count)
                log_icon = "✅" if is_match else "⚠️"
                diff_count = abs(expected_count - actual_count)
                log_detail = "100%吻合" if is_match else f"筆數不符，差 {diff_count} 筆"
                audit_logs.append(f"【{sheet_name}】: 預期 {expected_count} / 實得 {actual_count} ({log_icon} {log_detail})")
            else:
                audit_logs.append(f"【{sheet_name}】: 解析筆數 {actual_count} 筆 (未取得 AI 預期統計)")

            # ── 檢核邏輯 ─────────────────────────────────────────
            def check_reliability(row):
                warnings = []
                target_cols = [c for c in df.columns if c != 'AI_辨識疑慮']
                id_col = next((c for c in target_cols if '編號' in str(c) or '號' in str(c)), None)
                name_col = next((c for c in target_cols if '姓名' in str(c) or '名' in str(c)), None)

                if not id_col or pd.isna(row.get(id_col)) or str(row.get(id_col)).strip() in ['?', 'nan', '']:
                    warnings.append("🚨嚴重:編號遺失")
                else:
                    id_val = str(row[id_col]).strip()
                    if not id_val.isdigit():
                        warnings.append(f"⚠️異狀:編號非純數字({id_val})")

                if not name_col or pd.isna(row.get(name_col)) or str(row.get(name_col)).strip() in ['?', 'nan', '']:
                    warnings.append("🚨嚴重:姓名遺失")

                ai_note = str(row.get('AI_辨識疑慮', '')).strip()
                if ai_note and ai_note.lower() not in ['nan', '否', '正常', '?', 'none', 'false', 'true', '0', '1']:
                    warnings.append(f"🤖AI原始註記:{ai_note}")

                return " | ".join(warnings) if warnings else "正常"

            df['AI_辨識疑慮'] = df.apply(check_reliability, axis=1)

            # ── 重複檢測 ─────────────────────────────────────────
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
            for uf in uploaded_files:
                try:
                    genai.delete_file(uf.name)
                except Exception:
                    pass
            if img_tmp_dir:
                shutil.rmtree(img_tmp_dir, ignore_errors=True)
                    
    # 移除未定義的 progress_bar 呼叫，改用 log_status 或檢查是否存在
    if 'progress_bar' in locals() or 'progress_bar' in globals():
        try:
            progress_bar.progress(1.0)
        except:
            pass
        
    log_status("✅ 所有檔案皆已通過 AI 辨識與資料解析！準備產出報表...")
    time.sleep(1)
    
    if not status_container:
        # 確保這些 Streamlit 元件存在才嘗試清空
        if 'progress_text' in locals() or 'progress_text' in globals():
            try: progress_text.empty()
            except: pass
        if 'progress_bar' in locals() or 'progress_bar' in globals():
            try: progress_bar.empty()
            except: pass

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
