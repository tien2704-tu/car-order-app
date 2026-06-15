# --- 3. 新增工單表單區塊 ---
st.header("📝 確認與填寫工單資料")
with st.form("order_form", clear_on_submit=False):
    col1, col2 = st.columns(2)
    with col1:
        wo_number = st.text_input("工單單號", value=extracted_wo if extracted_wo else f"WO-2026{len(st.session_state.order_list)+1:03d}")
        car_number = st.text_input("車牌號碼", value=extracted_plate if extracted_plate else "", placeholder="例如: ANV-1055")
    with col2:
        category = st.selectbox("保養類別", ["輪胎定位", "底盤系統", "定期保養", "引擎系統", "冷氣系統", "其他"])
        price = st.number_input("金額 (元)", min_value=0, value=int(extracted_price), step=100)
        
    item_name = st.text_input("保養項目明細描述", placeholder="例如: 米其林輪胎更換、四輪定位")
    
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
