# app.py
import streamlit as st
import re
import gdown
import os
import tempfile
from processor import process_labor_pay_pdf

def extract_gdrive_id(url):
    """從 Google Drive 網址中萃取出 File ID"""
    pattern = r'(?:https?:\/\/)?(?:drive\.google\.com\/(?:file\/d\/|open\?id=)|docs\.google\.com\/(?:document|presentation|spreadsheets)\/d\/)([-\w]{25,})'
    match = re.search(pattern, url)
    if match:
        return match.group(1)
    return None

def main():
    st.set_page_config(page_title="勞作金自動化轉換系統", page_icon="📝", layout="centered")
    st.title("📄 勞作金名冊 OCR 轉換系統")
    st.markdown("請在下方輸入公開分享的 **Google Drive PDF 連結**，系統會自動辨識並轉換為 Excel 檔案。")
    st.info("💡 提示：請確定 Google Drive 檔案共用狀態已設定為「知道連結的人皆可檢視」。")

    gdrive_url = st.text_input("🔗 輸入 Google Drive 網址", placeholder="https://drive.google.com/file/d/1a2b3c4d5e/view?usp=sharing")

    if st.button("開始轉換 Excel", type="primary", use_container_width=True):
        if not gdrive_url:
            st.warning("⚠️ 請先貼上連結！")
            return
        
        file_id = extract_gdrive_id(gdrive_url)
        if not file_id:
            st.error("❌ 無法解析連結！請確定這是 Google Drive 的完整分享網址。")
            return

        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_path = os.path.join(temp_dir, "input_roster.pdf")
            excel_path = os.path.join(temp_dir, "output_result.xlsx")
            
            with st.spinner("⏳ 正在從雲端硬碟下載名冊..."):
                download_url = f'https://drive.google.com/uc?export=download&id={file_id}'
                try:
                    # 使用 gdown 下載大檔案
                    gdown.download(download_url, pdf_path, quiet=False)
                except Exception as e:
                    st.error(f"下載失敗！可能是因為檔案權限未公開。\n系統錯誤訊息: {e}")
                    return

            if not os.path.exists(pdf_path):
                 st.error("❌ 下載失敗：未能成功下載 PDF，請確認連結使否有效且公開分享。")
                 return

            with st.spinner("🧠 檔案下載成功！正在透過 AI 執行高精準度 OCR 轉換，請稍候（約需數分鐘）..."):
                try:
                    # 呼叫已經封裝好的轉換邏輯
                    success, error_msg = process_labor_pay_pdf(pdf_path, excel_path)
                    
                    if success and os.path.exists(excel_path):
                        st.success("🎉 辨識與轉換成功！")
                        st.info(f"📊 {error_msg}")
                        
                        with open(excel_path, "rb") as f:
                            excel_data = f.read()

                        st.download_button(
                            label="📥 點我下載 Excel 報表",
                            data=excel_data,
                            file_name="自動化勞作金名冊.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.error(f"❌ 轉換失敗：{error_msg}")

                except Exception as e:
                    st.error(f"❌ OCR 處理過程中發生未預期的錯誤：{e}")

if __name__ == "__main__":
    main()
