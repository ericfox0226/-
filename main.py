import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Mm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from io import BytesIO
from datetime import datetime

# --- 頁面配置 ---
st.set_page_config(page_title="公司零用金系統", layout="centered")

# 自定義 CSS
st.markdown("""
    <style>
    div.stButton > button:first-child { width: 100%; height: 3em; font-size: 18px; }
    .total-preview { 
        background-color: #f8f9fa; 
        padding: 20px; 
        border-radius: 10px; 
        text-align: center; 
        border: 2px solid #343a40;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("📂 雜支明細表自動化系統")

if 'data_list' not in st.session_state:
    st.session_state.data_list = []

# --- 頂部總金額預覽 ---
if st.session_state.data_list:
    total_amt = sum(d['金額'] for d in st.session_state.data_list)
    # 根據金額正負顯示顏色：負數（支出）用紅色，正數用藍色
    text_color = "#d32f2f" if total_amt < 0 else "#01579b"
    st.markdown(f"""
        <div class="total-preview">
            <span style="font-size: 16px; color: #666;">目前累計總餘額</span><br>
            <span style="font-size: 32px; font-weight: bold; color: {text_color};">NT$ {total_amt:,}</span>
        </div>
    """, unsafe_allow_html=True)

# --- 輸入區塊 ---
with st.expander("🖋️ 新增資料 (金額預設為支出)", expanded=True):
    today_str = datetime.now().strftime("%m/%d")
    date_val = st.text_input("日期", value=today_str)
    content_val = st.text_input("花費內容")
    
    col_a, col_b = st.columns(2)
    with col_a:
        # 修改點：讓預設步長為 -1，並在說明中提醒
        # 如果用戶輸入 100，我們在邏輯中把它轉為 -100 (除非他手動輸入 +100)
        raw_amount = st.number_input("金額 (直接輸入數字即為支出)", step=1, value=0)
    with col_b:
        location_val = st.text_input("工地全名")

    if st.button("➕ 新增至清單"):
        if date_val and content_val and location_val:
            # 邏輯調整：如果用戶輸入的是正數且不為0，自動轉為負數 (支出)
            # 如果用戶刻意要輸入收入，他們可以輸入負數的負數，但這不直觀
            # 更好的做法是：我們假設輸入的金額就是「支出金額」
            actual_amount = -abs(raw_amount) if raw_amount > 0 else raw_amount
            
            st.session_state.data_list.append({
                "日期": date_val, 
                "內容": content_val, 
                "金額": actual_amount, 
                "工地": location_val
            })
            st.rerun()

# --- 排序與邏輯函式 ---
def get_sorted_data(data):
    def sort_key(item):
        try:
            return datetime.strptime(f"{datetime.now().year}/{item['日期']}", "%Y/%m/%d")
        except: return datetime.max
    return sorted(data, key=sort_key)

def get_location_mapping(sorted_data):
    unique_locations = []
    for d in sorted_data:
        if d["工地"] not in unique_locations:
            unique_locations.append(d["工地"])
    return {loc: chr(65 + i) for i, loc in enumerate(unique_locations)}

# --- Word 生成邏輯 (垂直排列 + 置中) ---
def export_word(data, mapping):
    doc = Document()
    # A4 窄邊距設定
    section = doc.sections[0]
    section.top_margin = Mm(15)
    section.bottom_margin = Mm(15)
    section.left_margin = Mm(15)
    section.right_margin = Mm(15)

    title = doc.add_paragraph("雜支明細表")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.runs[0]
    run.font.size = Pt(18)
    run.bold = True
    
    doc.add_paragraph(f"報告日期：{datetime.now().strftime('%Y/%m/%d')}")
    doc.add_paragraph(f"經手人：_________________")

    # 分配左右兩側資料 (垂直排列邏輯)
    rows_per_page = 28 
    left_side = data[:rows_per_page]
    right_side = data[rows_per_page:rows_per_page*2]

    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    
    # 設定標題欄位
    headers = ["日期", "內容", "金額", "工地代號"] * 2
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.paragraphs[0].runs[0].font.bold = True

    last_d_l = None
    last_d_r = None
    
    for i in range(len(left_side)):
        row_cells = table.add_row().cells
        
        # 左側資料處理
        d_l = left_side[i]
        show_date_l = "" if d_l["日期"] == last_d_l else d_l["日期"]
        last_d_l = d_l["日期"]
        l_vals = [show_date_l, d_l["內容"], f"{d_l['金額']:,}", mapping[d_l["工地"]]]
        
        # 右側資料處理
        r_vals = ["", "", "", ""]
        if i < len(right_side):
            d_r = right_side[i]
            show_date_r = "" if d_r["日期"] == last_d_r else d_r["日期"]
            last_d_r = d_r["日期"]
            r_vals = [show_date_r, d_r["內容"], f"{d_r['金額']:,}", mapping[d_r["工地"]]]

        # 填入儲存格並套用置中格式
        for idx, val in enumerate(l_vals + r_vals):
            cell = row_cells[idx]
            cell.text = str(val)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # 底部總結
    total = sum(d['金額'] for d in data)
    p_total = doc.add_paragraph(f"\n總計金額：NT$ {total:,} 元")
    p_total.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # 工地索引
    doc.add_paragraph("-" * 20)
    doc.add_paragraph("【工地代號索引】").bold = True
    for name, code in mapping.items():
        doc.add_paragraph(f"{code} : {name}")

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- 下載與表格顯示 ---
if st.session_state.data_list:
    sorted_list = get_sorted_data(st.session_state.data_list)
    loc_mapping = get_location_mapping(sorted_list)
    
    st.subheader("📊 本月明細預覽")
    st.table(pd.DataFrame([{
        "日期": d["日期"], "內容": d["內容"], "金額": d["金額"], "工地": d["工地"]
    } for d in sorted_list]))

    col1, col2 = st.columns(2)
    with col1:
        if st.button("⏪ 刪除最後一筆"):
            st.session_state.data_list.pop()
            st.rerun()
    with col2:
        if st.button("🗑️ 全部清空"):
            st.session_state.data_list = []
            st.rerun()

    word_file = export_word(sorted_list, loc_mapping)
    st.download_button(
        label="📥 下載 A4 垂直排列報表",
        data=word_file,
        file_name=f"雜支明細表_{datetime.now().strftime('%m%d')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
