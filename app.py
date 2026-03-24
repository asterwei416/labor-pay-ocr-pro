import streamlit as st
import re
import gdown
import os
import tempfile
import glob
from processor import process_labor_pay_pdf

def extract_gdrive_id(url):
    """從 Google Drive 網址中萃取出 ID 與類型 (檔案或資料夾)"""
    # 判斷是否為資料夾 (包含 /drive/folders/ 或 /open?id= 且非文件特定 url)
    folder_pattern = r'(?:https?:\/\/)?(?:drive\.google\.com\/drive\/(?:u\/\d+\/)?folders\/|drive\.google\.com\/open\?id=)([-\w]{25,})'
    folder_match = re.search(folder_pattern, url)
    if folder_match:
        return 'folder', folder_match.group(1)
        
    # 判斷是否為單一檔案
    file_pattern = r'(?:https?:\/\/)?(?:drive\.google\.com\/(?:file\/d\/|open\?id=)|docs\.google\.com\/(?:document|presentation|spreadsheets)\/d\/)([-\w]{25,})'
    file_match = re.search(file_pattern, url)
    if file_match:
        return 'file', file_match.group(1)
        
    return None, None

def main():
    st.set_page_config(page_title="勞作金自動化轉換系統", page_icon="📝", layout="centered")
    st.title("📄 勞作金名冊 OCR 轉換系統")
    st.markdown("支援單一 PDF 檔案或是 **整個裝滿 PDF 的資料夾**。\n請在下方輸入公開分享的 Google Drive 連結，系統會自動辨識並轉換為 Excel 檔案。")
    st.info("💡 提示：請確定 Google Drive 檔案或資料夾的共用狀態已設定為「知道連結的人皆可檢視」。")

    gdrive_url = st.text_input("🔗 輸入 Google Drive 網址", placeholder="https://drive.google.com/drive/folders/1k1xOmDu... 或 /file/d/...")

    if st.button("開始批量轉換 Excel", type="primary", use_container_width=True):
        if not gdrive_url:
            st.warning("⚠️ 請先貼上連結！")
            return
        
        link_type, file_id = extract_gdrive_id(gdrive_url)
        if not link_type or not file_id:
            st.error("❌ 無法解析連結！請確定這是 Google Drive 的完整分享網址 (支援資料夾或單一檔案)。")
            return

        with tempfile.TemporaryDirectory() as temp_dir:
            excel_path = os.path.join(temp_dir, "output_result.xlsx")
            target_path = ""
            
            with st.spinner(f"⏳ 正在從雲端硬碟下載{'資料夾' if link_type == 'folder' else '檔案'}..."):
                try:
                    if link_type == 'folder':
                        target_path = os.path.join(temp_dir, "downloaded_folder")
                        # 備註: gdown.download_folder 具有遞迴下載功能
                        gdown.download_folder(id=file_id, output=target_path, quiet=True, use_cookies=False)
                        
                        # 檢查下載下來的目錄有沒有 PDF
                        pdf_files = glob.glob(os.path.join(target_path, "**/*.pdf"), recursive=True)
                        if not pdf_files:
                            st.error(f"❌ 錯誤：在下載的資料夾中沒有找到任何 PDF 檔案！請確認資料夾內有副檔名為 .pdf 的檔案。")
                            return
                    else:
                        target_path = os.path.join(temp_dir, "input_roster.pdf")
                        download_url = f'https://drive.google.com/uc?export=download&id={file_id}'
                        gdown.download(download_url, target_path, quiet=True)
                        
                        if not os.path.exists(target_path):
                            st.error("❌ 下載失敗：未能成功下載 PDF，請確認連結是否有效且公開分享。")
                            return
                except Exception as e:
                    st.error(f"下載失敗！可能是因為檔案或資料夾權限未公開。\n系統錯誤訊息: {e}")
                    return

            st.success(f"✅ 成功下載 {'整個目錄的 PDF 檔案' if link_type == 'folder' else '單一 PDF 檔案'}。")

            with st.status("🧠 正在透過 AI 執行高精準度批量 OCR 轉換，請稍候（可能需要數分鐘）...", expanded=True) as status:
                try:
                    # 呼叫已經封裝好的轉換邏輯，現在 target_path 可能是目錄或檔案
                    success, audit_msg = process_labor_pay_pdf(target_path, excel_path, status)
                    
                    if success and os.path.exists(excel_path):
                        status.update(label="✅ 所有檔案辨識與轉換成功！", state="complete", expanded=False)
                        st.success("🎉 所有檔案辨識與轉換成功！")
                        st.info(f"📊 **系統監控日誌：**\n\n{audit_msg}")
                        
                        with open(excel_path, "rb") as f:
                            excel_data = f.read()

                        st.download_button(
                            label="📥 點我下載合併 Excel 報表 (包含多分頁)",
                            data=excel_data,
                            file_name="自動化勞作金名冊_包含多分頁.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        status.update(label="❌ 處理失敗", state="error", expanded=True)
                        st.error(f"❌ 轉換失敗：{audit_msg}")

                except Exception as e:
                    status.update(label="❌ 發生未預期的錯誤", state="error", expanded=True)
                    st.error(f"❌ OCR 處理過程中發生未預期的錯誤：{e}")

if __name__ == "__main__":
    main()
