import os
import sys
import pandas as pd
from processor import process_labor_pay_pdf

# 模擬環境
def run_diagnostic():
    print("--- 勞作金 OCR 診斷工具 ---")
    
    # 1. 檢查 API Key
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        print("❌ 錯誤: 未偵測到環境變數 GEMINI_API_KEY。")
    else:
        print(f"✅ 偵測到 API Key (前4位: {api_key[:4]}...)")

    # 2. 測試匯入
    try:
        import google.generativeai as genai
        print("✅ google-generativeai 套件已安裝。")
    except ImportError:
        print("❌ 錯誤: 請執行 pip install google-generativeai")

    print("\n--- 程式碼熱檢查 ---")
    # 這裡檢查 processor.py 的邏輯是否與使用者描述的「辨識疑慮」功能相符
    with open("processor.py", "r", encoding="utf-8") as f:
        content = f.read()
        if "AI_辨識疑慮" in content:
            print("✅ 發現 'AI_辨識疑慮' 相關邏輯。")
        else:
            print("❌ 未發現 'AI_辨識疑慮' 相關邏輯。")

    print("\n--- 結束診斷 ---")

if __name__ == "__main__":
    run_diagnostic()
