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

# 設定網頁標題與 RWD 手機版配置
st.set_page_config(page_title="汽車保養工單系統", layout="centered")

# 初始化 EasyOCR 辨識器
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'])

reader = load_ocr()

st.title("🚗 汽車保養工單系統 (AI 辨識升級版)")
st.write("提示：若手機未彈出相機畫面，請檢查瀏覽器的相機權限，或直接使用「上傳照片」功能。")

# --- 1. 初始化資料庫 (在記憶體中模擬) ---
if "order_list" not in st.session_state:
    st.session_state.order_list = [
        {"工單單號": "WO-2026001", "車牌號碼": "ABC-1234", "保養項目": "機油更換", "保養類別": "定期保養", "金額": 2500},
        {"工單單號": "WO-2026002", "車牌號碼": "XYZ-5678", "保養項目": "前煞車皮更換", "保養類別": "底盤系統", "金額": 1800},
    ]

# --- 2. AI 影像自動擷取區塊 (雙軌制) ---
st.header("📸 AI 車牌自動擷取")
captured_car_number = ""

# 讓使用者自由選擇要「用相機現場拍」還是「從相簿選照片」
input_method = st.radio("請選擇輸入方式：", ["使用手機相機現場拍照", "從手機相簿上傳照片/檔案"])

img_file = None

if input_method == "使用手機相機現場拍照":
    # 若瀏覽器阻擋相機，此區塊可能顯示空白，故提供第二方案
    img_file = st.camera_input("請對準車牌拍照")
else:
    img_file = st.file_uploader("請選擇車牌照片", type=["jpg", "jpeg", "png"])

# 開始進行 AI 辨識
if img_file is not None:
    image = Image.open(img_file)
    img_array = np.array(image)
    
    with st.spinner("AI 偵測車牌中..."):
        results = reader.readtext(img_array)
        detected_texts = [text_res[1] for text_res in results]
        combined_text = "".join(detected_texts).upper().replace(" ", "")
        
        # 台灣車牌正規格式篩選
        plate_match = re.search(r'[A-Z0-9]{2,4}[-─]?[A-Z0-9]{2,4}', combined_text)
        
        if plate_match:
            captured_car_number = plate_match.group(0)
            st.success(f"🎯 成功擷取車牌號碼：{captured_car_number}")
        else:
            all_words = re.findall(r'[A-Z0-9]+', combined_text)
            if all_words:
                captured_car_number = max(all_words, key=len)
                st.warning(f"🤔 擷取到最接近文字：{captured_car_number}")
            else:
                st.error("❌ 無法辨識文字，請確保照片清晰且光線充足。")

# --- 3. 新增工單表單區塊 ---
st.header("📝 填寫工單資料")
with st.form("order_form", clear_on_submit=False):
    col1, col2 = st.columns(2)
    with col1:
        wo_number = st.text_input("工單單號", value=f"WO-2026{len(st.session_state.order_list)+1:03d}")
        car_number = st.text_input(
            "車牌號碼", 
            value=captured_car_number if captured_car_number else "", 
            placeholder="例如: ABC-1234"
        )
    with col2:
        category = st.selectbox("保養類別", ["定期保養", "引擎系統", "底盤系統", "冷氣系統", "電機電控", "輪胎定位", "其他"])
        price = st.number_input("金額 (元)", min_value=0, step=100)
        
    item_name = st.text_input("保養項目明細", placeholder="例如: 換機油、電瓶更換")
    
    submit_btn = st.form_submit_button("確認新增工單")
    
    if submit_btn:
        if car_number and item_name and price > 0:
            new_order = {
                "工單單號": wo_number,
                "車牌號碼": car_number,
                "保養項目": item_name,
                "保養類別": category,
                "金額": price
            }
            st.session_state.order_list.append(new_order)
            st.success(f"🎉 工單 {wo_number} 新增成功！")
            st.rerun()
        else:
            st.error("❌ 請填寫完整的車牌、項目與金額！")

# 轉換為 DataFrame 與畫圖、匯出 Excel 邏輯 (保持不變)
df_detail = pd.DataFrame(st.session_state.order_list)

st.header("📊 保養類別金額統計")
if not df_detail.empty:
    df_summary = df_detail.groupby("保養類別")["金額"].sum().reset_index()
    fig = px.pie(df_summary, values='金額', names='保養類別', title='各類別保養金額比例', hole=0.3)
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

st.header("📋 歷史工單明細")
st.dataframe(df_detail, use_container_width=True)

excel_data = generate_excel(df_detail, df_detail.groupby("保養類別")["金額"].sum().reset_index())
st.download_button(
    label="📥 匯出並下載保養工單 Excel 檔",
    data=excel_data,
    file_name="汽車保養工單統計表.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
