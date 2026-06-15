import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils.dataframe import dataframe_to_rows

# 設定網頁標題與 RWD 手機版配置
st.set_page_config(page_title="汽車保養工單系統", layout="centered")

st.title("🚗 汽車保養工單系統")
st.write("填寫下方表單以記錄工單，並可即時查看統計與下載 Excel。")

# --- 1. 初始化資料庫 (在記憶體中模擬) ---
if "order_list" not in st.session_state:
    st.session_state.order_list = [
        {"工單單號": "WO-2026001", "車牌號碼": "ABC-1234", "保養項目": "機油更換", "保養類別": "定期保養", "金額": 2500},
        {"工單單號": "WO-2026002", "車牌號碼": "XYZ-5678", "保養項目": "前煞車皮更換", "保養類別": "底盤系統", "金額": 1800},
        {"工單單號": "WO-2026003", "車牌號碼": "ABC-1234", "保養項目": "冷氣濾網更換", "保養類別": "冷氣系統", "金額": 1200},
    ]

# --- 2. 新增工單表單區塊 ---
st.header("📝 新增保養工單")
with st.form("order_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        wo_number = st.text_input("工單單號", value=f"WO-2026{len(st.session_state.order_list)+1:03d}")
        car_number = st.text_input("車牌號碼", placeholder="例如: ABC-1234")
    with col2:
        category = st.selectbox("保養類別", ["定期保養", "引擎系統", "底盤系統", "冷氣系統", "電機電控", "輪胎定位", "其他"])
        price = st.number_input("金額 (元)", min_value=0, step=100)
        
    item_name = st.text_input("保養項目明細", placeholder="例如: 換冷氣濾網、煞車油")
    
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
        else:
            st.error("❌ 請填寫完整的車牌、項目與金額！")

# 轉換為 DataFrame 方便後續處理
df_detail = pd.DataFrame(st.session_state.order_list)

# --- 3. 圖表統計區塊 ---
st.header("📊 保養類別金額統計")
if not df_detail.empty:
    df_summary = df_detail.groupby("保養類別")["金額"].sum().reset_index()
    
    # 使用 Plotly 繪製網頁互動式圓餅圖
    fig = px.pie(df_summary, values='金額', names='保養類別', title='各類別保養金額比例', hole=0.3)
    fig.update_traces(textinfo='value+percent')
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("目前尚無資料可顯示圖表。")

# --- 4. Excel 自動匯出功能邏輯 ---
def generate_excel(df_d, df_s):
    output = BytesIO()
    wb = Workbook()
    
    # 工作表一：統計摘要與圓餅圖
    ws_summary = wb.active
    ws_summary.title = "類別統計摘要"
    for r in dataframe_to_rows(df_s, index=False, header=True):
        ws_summary.append(r)
        
    # 建立 Excel 內建圓餅圖
    pie = PieChart()
    pie.title = "各類別保養金額比例圖"
    data = Reference(ws_summary, min_col=2, min_row=1, max_row=len(df_s) + 1)
    labels = Reference(ws_summary, min_col=1, min_row=2, max_row=len(df_s) + 1)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    
    # 【修正地方】：使用安全相容的新版語法建立資料標籤
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showVal = True  # 在 Excel 圓餅圖上顯示數值
    
    ws_summary.add_chart(pie, "D2")
    
    # 工作表二：明細
    ws_detail = wb.create_sheet(title="工單明細")
    for r in dataframe_to_rows(df_d, index=False, header=True):
        ws_detail.append(r)
        
    wb.save(output)
    return output.getvalue()

# --- 5. 明細顯示與下載按鈕 ---
st.header("📋 歷史工單明細")
st.dataframe(df_detail, use_container_width=True)

# 產生 Excel 檔案並提供下載
excel_data = generate_excel(df_detail, df_detail.groupby("保養類別")["金額"].sum().reset_index())
st.download_button(
    label="📥 匯出並下載保養工單 Excel 檔",
    data=excel_data,
    file_name="汽車保養工單統計表.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
