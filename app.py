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

st.title("🚗 汽車保養工單系統 (多欄位 AI 擷取版)")
st.write("上傳工單照片，AI 將自動擷取：車牌號碼、總金額、行駛里程、保養項目明細。")

# --- 2. 初始化資料庫 (在記憶體中模擬) ---
if "order_list" not in st.session_state:
    st.session_state.order_list = [
        {"工單單號": "D260513-375", "車牌號碼": "ANV-1055", "行駛里程": "140,029", "保養項目": "215/55R17 Michelin PRIMACY 5, 輪胎拆裝工資, 輪胎平衡校正, 氮氣填充, 四輪定位-電腦3D, 煞車來令片(前/WTC), 力魔LM3318-快速清潔噴劑, 底盤拆裝工資, 煞車來令片(後/WTC), 煞車盤(後)", "保養類別": "輪胎定位", "金額": 24300},
    ]

# --- 3. AI 影像自動擷取區塊 ---
st.header("📸 AI 工單自動擷取")

extracted_wo = ""
extracted_plate = ""
extracted_mileage = ""
extracted_price = 0
extracted_items = []

input_method = st.radio("請選擇輸入方式：", ["從手機相簿上傳照片/檔案", "使用手機相機現場拍照"])
img_file = st.file_uploader("請選擇工單照片", type=["jpg", "jpeg", "png"]) if input_method == "從手機相簿上傳照片/檔案" else st.camera_input("請對準工單拍照")

if img_file is not None:
    image = Image.open(img_file)
    img_np = np.array(image)
    
    # 影像預處理：強化對比度
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    enhanced_img = cv2.equalizeHist(gray)
    
    with st.spinner("AI 正在分析工單資訊中..."):
        results = reader.readtext(enhanced_img, detail=1, paragraph=False)
        
        all_texts = []
        raw_lines = []
        for res in results:
            text_curr = res[1].strip().replace(" ", "")
            raw_lines.append(text_curr)
            all_texts.append(text_curr.upper())
            
        full_text_block = "||".join(all_texts)
        
        # 顯示除錯紀錄，方便查看辨識字體
        with st.expander("🔍 檢視 AI 影像除錯紀錄"):
            st.write(raw_lines)

        # ---- 核心關鍵字擷取演算法 ----
        
        # 1. 擷取車牌號碼 (尋找「車牌」或「車號」關鍵字)
        plate_pattern = r'[A-Z0-9]{2,4}[-─][A-Z0-9]{2,4}'
        plates_found = re.findall(plate_pattern, full_text_block)
        if plates_found:
            extracted_plate = plates_found[0]
        else:
            for i, text in enumerate(all_texts):
                if any(k in text for k in ["車牌", "車號", "車輛"]):
                    for offset in range(0, 3):
                        if i + offset < len(all_texts):
                            candidate = re.sub(r'[^\w-]', '', all_texts[i + offset])
                            if len(candidate) >= 6 and any(c.isdigit() for c in candidate):
                                extracted_plate = candidate
                                break
                if extracted_plate: break

        # 2. 擷取行駛里程 (尋找「里程」關鍵字)
        for i, text in enumerate(all_texts):
            if "里程" in text or "公里" in text:
                for offset in range(0, 3):
                    if i + offset < len(all_texts):
                        candidate = all_texts[i + offset]
                        mileage_match = re.search(r'[\d,]+', candidate)
                        if mileage_match and len(mileage_match.group(0)) >= 3:
                            extracted_mileage = mileage_match.group(0)
                            break
            if extracted_mileage: break

        # 3. 擷取總金額 (尋找「總金額」或「合計」關鍵字)
        for i, text in enumerate(all_texts):
            if any(k in text for k in ["總金額", "金額", "總計", "合計"]):
                combined_look = "".join(all_texts[i:i+4])
                digits = re.findall(r'[\d,.]+', combined_look)
                # 【此處已修正修正】：確保 for 迴圈完整閉合
                for d_str in digits:
                    clean_d = d_str.replace(",", "").split(".")[0]
                    if clean_d.isdigit() and 100 < int(clean_d) < 1000000:
                        extracted_price = int(clean_d)
                        break
            if extracted_price > 0: break

        # 4. 擷取保養項目明細
        item_keywords = ["輪胎", "工資", "平衡", "定位", "氮氣", "煞車", "來令片", "清潔", "噴劑", "機油", "濾網", "電瓶", "火星塞", "皮帶", "保養"]
        for text in raw_lines:
            if any(k in text for k in item_keywords) and "總金額" not in text and "馳加" not in text:
                clean_item = re.sub(r'^[A-Z0-9]{5,15}', '', text) 
                if len(clean_item) > 2 and clean_item not in extracted_items:
                    extracted_items.append(clean_item)
                    
        extracted_items_str = ", ".join(extracted_items)

        # 順便抓工單單號當作識別 (馳加專用)
        wo_found = re.search(r'D[0-9]{5,6}[-─][0-9]+', full_text_block)
        if wo_found: extracted_wo = wo_found.group(0)

        # ---- 畫面上秀出 AI 擷取結果 ----
        st.subheader("🎯 AI 欄位擷取結果")
        col_res1, col_res2, col_res3 = st.columns(3)
        col_res1.metric("車牌號碼", extracted_plate if extracted_plate else "未偵測到")
        col_res2.metric("行駛里程", f"{extracted_mileage} KM" if extracted_mileage else "未偵測到")
        col_res3.metric("辨識總金額", f"${extracted_price:,}" if extracted_price else "未偵測到")
        st.text_area("自動擷取保養項目清單", value=extracted_items_str, height=70)

# --- 4. 確認與填寫工單資料表單 ---
st.header("📝 確認與填寫工單資料")
with st.form("order_form", clear_on_submit=False):
    col1, col2 = st.columns(2)
    with col1:
        wo_number = st.text_input("工單單號", value=extracted_wo if extracted_wo else f"WO-2026{len(st.session_state.order_list)+1:03d}")
        car
