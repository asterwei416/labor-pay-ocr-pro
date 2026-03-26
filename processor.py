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

    【有效列判斷——最優先執行】
    - 只擷取「編號欄有數字、姓名欄有中文姓名」的人員記錄列。
    - 以下內容**不是人員記錄，直接跳過，不輸出**：
      * 計算公式列（如「16天×35人=560天」、「4.5天×2人=9天」等含「×」「=」「天」的算式）
      * 空白列
      * 頁首標題列

    【編號位數強制規則】
    - 所有有效人員編號**必為四位數**（例如 5007、9844、7210）。
    - 若讀到三位數（如 500），代表有一位被遺漏，必須重新審視後補上；補後仍不確定，填 `?` 補位（如 500?）。
    - ⚠️ 無論如何都**不可因編號不確定而漏行**——寧可輸出 `500?` 也不可省略該列。

    【任務步驟】
    1. 直接輸出 CSV，標題行必須為：「編號,姓名,AI_辨識疑慮」
       - CSV 必須包在 ```csv ``` 區塊內。
       - **不需要輸出任何前置思考或推論文字**，直接進入 CSV 區塊。
    2. 禁止漏行：所有有效人員列字跡模糊也必須輸出，完全無法辨識填 `?`。
    3. AI_辨識疑慮填寫規則：
       - 編號任一數字非 100% 清晰 → 填「是(數字疑慮：說明)」
       - 編號補位不確定 → 填「是(數字疑慮：編號疑似缺位，補填為 XXXX)」
       - 姓名有連筆難辨 → 填「是(姓名連筆：說明)」
       - 全部清晰確定 → 填「正常」
       - 完全失敗 → 填「?」
    4. CSV 區塊後輸出：`TOTAL_ROWS: [數字]`（只計算人員記錄列數，不含跳過的計算公式列）
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
            # 上傳最多重試 2 次（應對網路抖動）
            for attempt in range(3):
                try:
                    uploaded_file = genai.upload_file(pdf_file)
                    break
                except Exception as upload_err:
                    if attempt == 2:
                        raise
                    log_status(f"⚠️ `{base_name}` 上傳失敗（第 {attempt + 1} 次），3 秒後重試... ({upload_err})")
                    time.sleep(3)

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

            # 保留 AI 原始疑慮，補上「正常」預設值
            if 'AI_辨識疑慮' not in df.columns:
                df['AI_辨識疑慮'] = '正常'
            else:
                df['AI_辨識疑慮'] = df['AI_辨識疑慮'].fillna('正常').replace('', '正常')

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

    # ── 跨頁重複偵測 ──────────────────────────────────────────
    # 合併所有頁的資料，統一比對編號與姓名
    if all_dfs:
        # 建立含來源頁名的合併表
        combined_rows = []
        for sn, d in all_dfs.items():
            id_col   = next((c for c in d.columns if '編號' in str(c) or '號' in str(c)), None)
            name_col = next((c for c in d.columns if '姓名' in str(c) or '名' in str(c)), None)
            if not id_col or not name_col:
                continue
            for ridx, row in d.iterrows():
                raw_id   = str(row.get(id_col,   '')).strip()
                raw_name = str(row.get(name_col, '')).strip()
                if raw_id in ('', 'nan', '?') or raw_name in ('', 'nan', '?'):
                    continue
                combined_rows.append({
                    'sheet': sn, 'ridx': ridx,
                    'id': raw_id, 'name': raw_name,
                    'id_col': id_col, 'name_col': name_col,
                })

        flag_map = {}        # (sheet, ridx) -> list of flag strings
        ai_candidates = []   # 近似編號候選對，待 AI 二次判斷

        def add_flag(sheet, ridx, msg):
            flag_map.setdefault((sheet, ridx), []).append(msg)

        n = len(combined_rows)
        for i in range(n):
            for j in range(i + 1, n):
                a, b = combined_rows[i], combined_rows[j]
                # 同一頁同一列不比
                if a['sheet'] == b['sheet'] and a['ridx'] == b['ridx']:
                    continue

                id_a, id_b     = a['id'], b['id']
                name_a, name_b = a['name'], b['name']
                loc_b = f"{b['sheet']}({name_b}/{id_b})"
                loc_a = f"{a['sheet']}({name_a}/{id_a})"

                # 1. 編號相同 且 姓名相同 → 確定重複，直接標記
                if id_a == id_b and name_a == name_b and len(name_a) >= 2:
                    add_flag(a['sheet'], a['ridx'], f"🔴確定重複:{loc_b}")
                    add_flag(b['sheet'], b['ridx'], f"🔴確定重複:{loc_a}")
                    continue

                # 1b. 編號相同 但 姓名不同 → 直接標記（人工核查哪筆編號打錯）
                if id_a == id_b and name_a != name_b and len(name_a) >= 2 and len(name_b) >= 2:
                    add_flag(a['sheet'], a['ridx'], f"🟠同編號異名:{loc_b}")
                    add_flag(b['sheet'], b['ridx'], f"🟠同編號異名:{loc_a}")
                    continue

                # 2. 近似編號候選 → 交給 AI 二次判斷
                #    條件：長度相同、diff ≤2、名字有共同字
                #    過濾：若共同字只有姓氏（第一字）且 diff=2 → 噪音太多，略過
                if len(id_a) == len(id_b):
                    diff = sum(x != y for x, y in zip(id_a, id_b))
                    common_any = set(name_a) & set(name_b) - {'?', '', ' '}
                    if 1 <= diff <= 2 and len(common_any) >= 1:
                        # 過濾：共同字只剩姓氏且 diff=2 → 不進候選
                        only_surname = (common_any == {name_a[0]}
                                        and name_a[0] == name_b[0]
                                        and diff == 2)
                        if not only_surname:
                            ai_candidates.append({
                                'a': a, 'b': b,
                                'loc_a': loc_a, 'loc_b': loc_b,
                            })

        # AI 二次判斷近似編號候選對
        if ai_candidates:
            import json as _json
            pair_list = [
                {
                    "index": idx,
                    "a": {"id": c['a']['id'], "name": c['a']['name'], "sheet": c['a']['sheet']},
                    "b": {"id": c['b']['id'], "name": c['b']['name'], "sheet": c['b']['sheet']},
                }
                for idx, c in enumerate(ai_candidates)
            ]
            judge_prompt = f"""你是一位手寫工資表審核員。以下是從手寫文件 OCR 出的人員資料配對，編號長度相同且姓氏一致，但編號有 1-2 碼差異。

請逐一判斷：這兩筆在手寫環境下是否可能因為筆跡難以辨認而其實是同一個人？

判斷原則：
- 名字完全不同（如「家明」vs「家輝」，字型差異大）→ 不同人
- 名字高度相似（如「家明」vs「家朋」，手寫易混淆）→ 可疑
- 只有姓氏相同、名字毫無關聯 → 不同人

請以 JSON 陣列回傳，每項格式：{{"index": <數字>, "verdict": "可疑" 或 "不同人", "reason": "<一句話>"}}

配對清單：
{_json.dumps(pair_list, ensure_ascii=False)}"""

            try:
                judge_model = genai.GenerativeModel(model_name)
                judge_resp = judge_model.generate_content(judge_prompt)
                raw = judge_resp.text.strip()
                # 取出 JSON 陣列
                json_start = raw.find('[')
                json_end   = raw.rfind(']') + 1
                if json_start != -1 and json_end > json_start:
                    verdicts = _json.loads(raw[json_start:json_end])
                    for v in verdicts:
                        if v.get('verdict') == '可疑':
                            c = ai_candidates[v['index']]
                            reason = v.get('reason', '')
                            add_flag(c['a']['sheet'], c['a']['ridx'],
                                     f"🟡近似編號:{c['loc_b']}（{reason}）")
                            add_flag(c['b']['sheet'], c['b']['ridx'],
                                     f"🟡近似編號:{c['loc_a']}（{reason}）")
            except Exception as e:
                st.warning(f"⚠️ AI 近似編號判斷失敗，略過此步驟：{e}")

        # 將旗標寫回各頁 DataFrame
        for (sn, ridx), flags in flag_map.items():
            if sn not in all_dfs:
                continue
            d = all_dfs[sn]
            existing = str(d.at[ridx, 'AI_辨識疑慮']) if 'AI_辨識疑慮' in d.columns else '正常'
            if existing in ('正常', 'nan', ''):
                existing = ''
            else:
                existing = existing + ' / '
            d.at[ridx, 'AI_辨識疑慮'] = existing + ' / '.join(flags)

        # 建立「重複疑慮彙整」總表
        summary_rows = []
        for (sn, ridx), flags in flag_map.items():
            if sn not in all_dfs:
                continue
            d = all_dfs[sn]
            id_col   = next((c for c in d.columns if '編號' in str(c) or '號' in str(c)), None)
            name_col = next((c for c in d.columns if '姓名' in str(c) or '名' in str(c)), None)
            summary_rows.append({
                '來源頁':   sn,
                '編號':     d.at[ridx, id_col]   if id_col   else '',
                '姓名':     d.at[ridx, name_col] if name_col else '',
                '疑慮說明': ' / '.join(flags),
            })

        if summary_rows:
            summary_df = pd.DataFrame(summary_rows).sort_values(['疑慮說明', '編號']).reset_index(drop=True)
        else:
            summary_df = pd.DataFrame(columns=['來源頁', '編號', '姓名', '疑慮說明'])
            summary_df.loc[0] = ['（無疑慮）', '', '', '']

    # ── 寫出 Excel ───────────────────────────────────────────
    if all_dfs:
        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                # 第一頁：重複疑慮彙整總表
                summary_df.to_excel(writer, sheet_name='⚠️重複疑慮彙整', index=False)
                # 其餘各頁原始資料
                for sn, d in all_dfs.items():
                    d.to_excel(writer, sheet_name=sn, index=False)

            global_audit_msg = "\n".join(audit_logs)
            if summary_rows:
                global_audit_msg += f"\n\n🔍 跨頁重複疑慮共 {len(summary_rows)} 筆，請查看「⚠️重複疑慮彙整」頁。"
            return True, global_audit_msg
        except Exception as e:
            return False, f"Excel 打包失敗: {e}"
    else:
        return False, "所有 PDF 皆處理失敗，未生成 Excel。"
