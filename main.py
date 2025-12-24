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
    .stMetric { background-color: #f0f2f6; padding: 15px; border-radius: 10px; }
    </style>
    """, unsafe_allow_html=True)

st.title("📂 雜支明細表自動化系統")

# 初始化 session_state
if 'data_list' not in st.session_state:
    st.session_state.data_list = []

# --- 輸入區塊 ---
with st.expander("🖋️ 新增資料", expanded=True):
    today_str = datetime.now().strftime("%m/%d")
    date_val = st.text_input("日期", value=today_str)
    content_val = st.text_input("花費內容", placeholder="例如: 購買五金材料")
    
    col_a, col_b = st.columns(2)
    with col_a:
        # 移除 min_value=0，允許輸入負數
        amount_val = st.number_input("金額 (退款請輸負數)", step=1, value=0)
    with col_b:
        location_val = st.text_input("工地全名", placeholder="例如: 台北大巨蛋")

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("➕ 新增至清單"):
            if date_val and content_val and location_val:
                st.session_state.data_list.append({
                    "日期": date_val, "內容": content_val, "金額": amount_val, "工地": location_val
                })
                st.rerun()
            else:
                st.error("請確保日期、內容、工地皆已填寫")
    with col_btn2:
        if st.button("⏪ 刪除最後一筆"):
            if st.session_state.data_list:
                st.session_state.data_list.pop()
                st.rerun()

# --- 處理工地代號邏輯 ---
def get_location_mapping(data):
    unique_locations = []
    for d in data:
        if d["工地"] not in unique_locations:
            unique_locations.append(d["工地"])
    
    # 生成對照表 { "工地全名": "代號" }
    mapping = {loc: chr(65 + i) for i, loc in enumerate(unique_locations)} # 65 是 'A'
    return mapping

# --- 顯示與計算區塊 ---
if st.session_state.data_list:
    loc_mapping = get_location_mapping(st.session_state.data_list)
    
    st.subheader("📊 當月預覽")
    # 轉換預覽資料，顯示代號
    display_data = []
    for d in st.session_state.data_list:
        display_data.append({
            "日期": d["日期"],
            "內容": d["內容"],
            "金額": f"{d['金額']:,}",
            "工地代號": loc_mapping[d["工地"]]
        })
    st.table(pd.DataFrame(display_data))
    
    # 顯示代號索引參考
    with st.info("🏗️ 工地代號對照："):
        cols = st.columns(3)
        for i, (full_name, code) in enumerate(loc_mapping.items()):
            cols[i % 3].write(f"**{code}**: {full_name}")

    total = sum(d['金額'] for d in st.session_state.data_list)
    st.metric("目前累計總額", f"{total:,} 元")

    if st.button("🗑️ 全部清空"):
        st.session_state.data_list = []
        st.rerun()

# --- Word 生成邏輯 ---
def export_word(data, mapping):
    doc = Document()
    section = doc.sections[0]
    section.top_margin, section.bottom_margin = Mm(15), Mm(15)
    section.left_margin, section.right_margin = Mm(15), Mm(15)

    title = doc.add_paragraph("雜支明細表")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.runs[0]
    run.font.size = Pt(18)
    run.bold = True
    
    doc.add_paragraph(f"報告日期：{datetime.now().strftime('%Y/%m/%d')}")
    doc.add_paragraph(f"經手人：_________________")

    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    
    headers = ["日期", "內容", "金額", "工地"] * 2
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i in range(0, len(data), 2):
        row_cells = table.add_row().cells
        # 左半部
        d_l = data[i]
        row_cells[0].text = str(d_l["日期"])
        row_cells[1].text = str(d_l["內容"])
        row_cells[2].text = f"{d_l['金額']:,}"
        row_cells[3].text = mapping[d_l["工地"]] # 使用代號
        
        # 右半部
        if i + 1 < len(data):
            d_r = data[i+1]
            row_cells[4].text = str(d_r["日期"])
            row_cells[5].text = str(d_r["內容"])
            row_cells[6].text = f"{d_r['金額']:,}"
            row_cells[7].text = mapping[d_r["工地"]] # 使用代號

    # 總計
    total_amt = sum(d['金額'] for d in data)
    p_sum = doc.add_paragraph(f"\n總計金額：NT$ {total_amt:,} 元")
    p_sum.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # 新增：工地代號對照索引
    doc.add_paragraph("-" * 30)
    doc.add_paragraph("【工地代號對照索引】").bold = True
    for full_name, code in mapping.items():
        doc.add_paragraph(f"{code} ： {full_name}")

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- 下載按鈕 ---
if st.session_state.data_list:
    mapping = get_location_mapping(st.session_state.data_list)
    word_file = export_word(st.session_state.data_list, mapping)
    st.download_button(
        label="📥 下載 A4 報表 (含工地索引)",
        data=word_file,
        file_name=f"雜支明細表_{datetime.now().strftime('%m%d')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
