import streamlit as st
import pandas as pd
import plotly.express as px
import easyocr
import numpy as np
import re
from PIL import Image
from io import BytesIO
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils.dataframe import dataframe_to_rows

# 設定網頁標題與 RWD 手機版配置
st.set_page_config(page_title="汽車保養工單系統", layout="centered")

# 初始化 EasyOCR 辨識器 (載入繁體中文 'ch_tra' 與英文 'en')
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['ch_tra', 'en'])

reader = load_ocr()

st.title("🚗 汽車保養工單系統 (馳加單據優化版)")
st.write("請拍攝或上傳清晰的維修結帳單，AI 將自動擷取工單單號、車牌與總金額。")

# --- 1. 初始化資料庫 (在記憶體中模擬) ---
if "order_list" not in st.session_state:
    st.session_state.order_list = [
        {"工單單號": "WO-2026001", "車牌號碼": "ABC-1234", "保養項目": "機油更換", "保養類別": "定期保養", "金額": 2500},
        {"工單單號": "WO-2026002", "車牌號碼": "XYZ-5678", "保養項目": "前煞車皮更換", "保養類別": "底盤系統", "金額": 1800},
    ]

# --- 2. AI 影像自動擷取區塊 (工單專用邏輯) ---
st.header("📸 AI 工單自動擷取")

# 預設辨識欄位為空
extracted_wo = ""
extracted_plate = ""
extracted_price = 0

input_method = st.radio("請選擇輸入方式：", ["從手機相簿上傳照片/檔案", "使用手機相機現場拍照"])
img_file = st.file_uploader("請選擇工單照片", type=["jpg", "jpeg", "png"]) if input_method == "從手機相簿上傳照片/檔案" else st.camera_input("請對準工單拍照")

if img_file is not None:
    image = Image.open(img_file)
    img_array = np.array(image)
    
    with st.spinner("AI 正在深度解析工單內容，請稍候..."):
        # 進行 OCR 文字辨識
        results = reader.readtext(img_array)
        
        # 將辨識到的每一行文字純文字化，並去除空白以便比對
        lines = [res[1].strip().replace(" ", "") for res in results]
        full_text_block = "||".join(lines) # 用雙豎線串接方便正規表示法搜尋
        
        # 偵錯用：在網頁上展開 AI 實際看見的文字（若不需要可註解掉）
        with st.expander("🔍 查看 AI 原始辨識文字紀錄"):
            st.write(lines)

        # ---- 關鍵字匹配演算法 ----
        
        # 1. 擷取車牌號碼：尋找「車牌號碼」或「車號」後方的標準台灣車牌格式
        for line in lines:
            if "車牌" in line or "車號" in line:
                plate_match = re.search(r'[A-Z0-9]{2,4}[-─]?[A-Z0-9]{2,4}', line)
                if plate_match:
                    extracted_plate = plate_match.group(0)
                    break
        # 備援機制：如果沒在同一行找到，直接在整篇文字找最像車牌的字
        if not extracted_plate:
            all_plates = re.findall(r'[A-Z]{2,3}[-─][0-9]{4}|[0-9]{4}[-─][A-Z]{2,3}', full_text_block.upper())
            if all_plates:
                extracted_plate = all_plates[0]

        # 2. 擷取工單單號：尋找「單據編號」或「編號」後方的英數組合
        wo_match = re.search(r'單據編號[：:]?([A-Z0-9-]+)', full_text_block)
        if wo_match:
            extracted_wo = wo_match.group(1)
        else:
            # 馳加專用備援：直接抓 D260513-375 這種格式
            specific_wo = re.search(r'D[0-9]{6}-[0-9]+', full_text_block)
            if specific_wo:
                extracted_wo = specific_wo.group(0)

        # 3. 擷取總金額：尋找「總金額」後方的數字
        for i, line in enumerate(lines):
            if "總金額" in line:
                # 試圖從同行或下一行撈取包含逗號的數字
                price_digits = re.findall(r'[\d,]+', "".join(lines[i:i+2]))
                for digit_str in price_digits:
                    clean_num = int(digit_str.replace(",", ""))
                    if clean_num > 100: # 過濾掉頁碼或小雜訊
                        extracted_price = clean_num
                        break
                if extracted_price > 0:
                    break

        # ---- 辨識結果反饋 ----
        st.subheader("🎯 AI 欄位擷取結果")
        col_res1, col_res2, col_res3 = st.columns(3)
        col_res1.metric("自動擷取車牌", extracted_plate if extracted_plate else "未偵測到")
        col_res2.metric("自動工單單號", extracted_wo if extracted_wo else "未偵測到")
        col_res3.metric("自動辨識金額", f"${extracted_price:,}" if extracted_price else "未偵測到")

# --- 3. 新增工單表單區塊 ---
st.header("📝 確認與填寫工單資料")
with st.form("order_form", clear_on_submit=False):
    col1, col2 = st.columns(2)
    with col1:
        # 自動代入 AI 擷取到的單號與車牌，若沒抓到則手動填寫
        wo_number = st.text_input("工單單號", value=extracted_wo if extracted_wo else f"WO-2026{len(st.session_state.order_list)+1:03d}")
        car_number = st.text_input("車牌號碼", value=extracted_plate if extracted_plate else "", placeholder="例如: ANV-1055")
    with col2:
        # 馳加這張單主要為輪胎、四輪定位與底盤煞車，預設下拉選單供選擇
        category = st.selectbox("保養類別", ["輪胎定位", "底盤系統", "定期保養", "引擎系統", "冷氣系統", "其他"])
        price = st.number_input("金額 (元)", min_value=0, value=int(extracted_price), step=100)
        
    item_name = st.text_input("保養項目明細描述", placeholder="例如: 米其林輪胎更換、四輪定位、前後煞車片更換")
    
    submit_btn = st.form_submit_button("確認新增此筆工單至系統")
    
    if submit_btn:
        if car_number and price > 0:
            new_order = {
                "工單單號": wo_number,
                "車牌號碼": car_number,
                "保養項目": item_name if item_name else "未填寫明細",
                "保養類別": category,
                "金額": price
            }
            st.session_state.order_list.append(new_order)
            st.success(f"🎉 工單 {wo_number} 已成功寫入系統資料庫！")
            st.rerun()
        else:
            st.error("❌ 請確認車牌號碼與金額是否正確填寫！")

# --- 4. 圖表與 Excel 匯出邏輯 (保持穩定運作) ---
df_detail = pd.DataFrame(st.session_state.order_list)

st.header("📊 汽車保養類別金額統計圓餅圖")
if not df_detail.empty:
    df_summary = df_detail.groupby("保養類別")["金額"].sum().reset_index()
    fig = px.pie(df_summary, values='金額', names='保養類別', title='各類別累積消費比例', hole=0.3)
    fig.update_traces(textinfo='value+percent')
    st.plotly_chart(fig, use_container_width=True)

def generate_excel(df_d, df_s):
    output = BytesIO()
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "類別統計摘要"
    for r in dataframe_to_rows(df_s, index=False, header=True):
        ws_summary.append(r)
    pie = PieChart()
    pie.title = "各類別保養金額比例圖"
    data = Reference(ws_summary, min_col=2, min_row=1, max_row=len(df_s) + 1)
    labels = Reference(ws_summary, min_col=1, min_row=2, max_row=len(df_s) + 1)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showVal = True
    ws_summary.add_chart(pie, "D2")
    
    ws_detail = wb.create_sheet(title="工單明細")
    for r in dataframe_to_rows(df_d, index=False, header=True):
        ws_detail.append(r)
    wb.save(output)
    return output.getvalue()

st.header("📋 系統歷史工單總表")
st.dataframe(df_detail, use_container_width=True)

excel_data = generate_excel(df_detail, df_detail.groupby("保養類別")["金額"].sum().reset_index())
st.download_button(
    label="📥 匯出並下載保養工單 Excel 檔",
    data=excel_data,
    file_name="汽車保養工單統計表.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
