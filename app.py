import streamlit as st
import pandas as pd
import plotly.express as px
import easyocr
import numpy as np
import cv2
import re
from PIL import Image
from io import BytesIO
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils.dataframe import dataframe_to_rows

# --- 1. 網頁初始配置 ---
st.set_page_config(page_title="汽車保養工單系統", layout="centered")

# 初始化 EasyOCR 辨識器 (繁體中文與英文)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['ch_tra', 'en'], gpu=False)

reader = load_ocr()

st.title("🚗 汽車保養工單系統 (馳加單據高辨識率版)")
st.write("已強化 AI 影像處理演算法，提升對點矩陣列印字體與表格工單的辨識能力。")

# --- 2. 初始化資料庫 (在記憶體中模擬) ---
if "order_list" not in st.session_state:
    st.session_state.order_list = [
        {"工單單號": "D260513-375", "車牌號碼": "ANV-1055", "保養項目": "輪胎更換與定位", "保養類別": "輪胎定位", "金額": 24300},
    ]

# --- 3. AI 影像自動擷取區塊 ---
st.header("📸 AI 工單自動擷取")

extracted_wo = ""
extracted_plate = ""
extracted_price = 0

input_method = st.radio("請選擇輸入方式：", ["從手機相簿上傳照片/檔案", "使用手機相機現場拍照"])
img_file = st.file_uploader("請選擇工單照片", type=["jpg", "jpeg", "png"]) if input_method == "從手機相簿上傳照片/檔案" else st.camera_input("請對準工單拍照")

if img_file is not None:
    # 讀取圖片並轉為 OpenCV 格式
    image = Image.open(img_file)
    img_np = np.array(image)
    
    # 影像預處理：轉灰階並強化對比度，讓點矩陣淡字變清晰
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    enhanced_img = cv2.equalizeHist(gray) # 直方圖均衡化，強化文字邊緣
    
    with st.spinner("AI 正在使用高階影像演算法解析中..."):
        # 進行全文辨識
        results = reader.readtext(enhanced_img, detail=1, paragraph=False)
        
        # 整理辨識結果，保留文字、信心度與位置
        all_texts = []
        raw_lines = []
        for res in results:
            text_curr = res[1].strip().replace(" ", "")
            raw_lines.append(text_curr)
            all_texts.append(text_curr.upper())
            
        full_text_block = "||".join(all_texts)
        
        # 開啟偵錯工具，讓您看見 AI 到底讀到了什麼字
        with st.expander("🔍 檢視 AI 影像除錯紀錄 (若辨識不到請點開此處)"):
            st.write("AI 辨識出的文字片段：", raw_lines)

        # ---- 演算法升級：模糊匹配與前後文搜尋 ----
        
        # 1. 擷取車牌號碼 (包含台灣新舊式車牌正規表示法)
        plate_pattern = r'[A-Z0-9]{2,4}[-─][A-Z0-9]{2,4}'
        plates_found = re.findall(plate_pattern, full_text_block)
        if plates_found:
            extracted_plate = plates_found[0]
        else:
            # 模糊搜尋：尋找「車牌」或「車號」關鍵字週邊的英文數字
            for i, text in enumerate(all_texts):
                if "車牌" in text or "車號" in text or "車輛" in text:
                    for offset in range(0, 3):
                        if i + offset < len(all_texts):
                            candidate = all_texts[i + offset]
                            clean_candidate = re.sub(r'[^\w-]', '', candidate)
                            if len(clean_candidate) >= 6 and any(char.isdigit() for char in clean_candidate):
                                extracted_plate = clean_candidate
                                break
                if extracted_plate: break

        # 2.
