import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Mm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from io import BytesIO
from datetime import datetime

# --- 頁面配置與排序邏輯維持不變 ---
st.set_page_config(page_title="公司零用金系統", layout="centered")

if 'data_list' not in st.session_state:
    st.session_state.data_list = []

def get_sorted_data(data):
    def sort_key(item):
        date_str = item["日期"]
        try:
            if '/' in date_str:
                parts = date_str.split('/')
                if len(parts) == 2:
                    return datetime.strptime(f"{datetime.now().year}/{date_str}", "%Y/%m/%d")
                return datetime.strptime(date_str, "%Y/%m/%d")
            return date_str
        except:
            return date_str
    return sorted(data, key=sort_key)

def get_location_mapping(sorted_data):
    unique_locations = []
    for d in sorted_data:
        if d["工地"] not in unique_locations:
            unique_locations.append(d["工地"])
    return {loc: chr(65 + i) for i, loc in enumerate(unique_locations)}

# --- Word 生成邏輯（優化相同日期隱藏） ---
def export_word(data, mapping):
    doc = Document()
    section = doc.sections[0]
    section.top_margin, section.bottom_margin = Mm(15), Mm(15)
    section.left_margin, section.right_margin = Mm(15), Mm(15)

    title = doc.add_paragraph("雜支明細表")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.runs[0]
    run.font.size = Pt(18); run.bold = True
    
    doc.add_paragraph(f"報告日期：{datetime.now().strftime('%Y/%m/%d')}")
    doc.add_paragraph(f"經手人：_________________")

    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    
    # 表頭
    headers = ["日期", "內容", "金額", "工地"] * 2
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.paragraphs[0].runs[0].font.bold = True

    # 用於追蹤日期是否重複
    last_date_left = None
    last_date_right = None

    for i in range(0, len(data), 2):
        row_cells = table.add_row().cells
        
        # --- 左半部處理 ---
        d_l = data[i]
        display_date_l = "" if d_l["日期"] == last_date_left else d_l["日期"]
        last_date_left = d_l["日期"] # 更新最後出現的日期
        
        l_vals = [display_date_l, d_l["內容"], f"{d_l['金額']:,}", mapping[d_l["工地"]]]
        
        # --- 右半部處理 ---
        r_vals = ["", "", "", ""]
        if i + 1 < len(data):
            d_r = data[i+1]
            display_date_r = "" if d_r["日期"] == last_date_right else d_r["日期"]
            last_date_right = d_r["日期"]
            r_vals = [display_date_r, d_r["內容"], f"{d_r['金額']:,}", mapping[d_r["工地"]]]
        
        # 填入並置中
        for idx, val in enumerate(l_vals + r_vals):
            row_cells[idx].text = str(val)
            row_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row_cells[idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    total_amt = sum(d['金額'] for d in data)
    doc.add_paragraph(f"\n總計金額：NT$ {total_amt:,} 元").alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_paragraph("-" * 30)
    doc.add_paragraph("【工地代號對照索引】").bold = True
    for name, code in mapping.items():
        doc.add_paragraph(f"{code} ： {name}")

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- Streamlit UI 部分維持原有邏輯 ---
with st.expander("🖋️ 新增資料", expanded=True):
    today_str = datetime.now().strftime("%m/%d")
    date_val = st.text_input("日期", value=today_str)
    content_val = st.text_input("花費內容")
    col_a, col_b = st.columns(2)
    with col_a:
        amount_val = st.number_input("金額", step=1, value=0)
    with col_b:
        location_val = st.text_input("工地全名")

    if st.button("➕ 新增至清單"):
        if date_val and content_val and location_val:
            st.session_state.data_list.append({"日期": date_val, "內容": content_val, "金額": amount_val, "工地": location_val})
            st.rerun()

if st.session_state.data_list:
    sorted_list = get_sorted_data(st.session_state.data_list)
    loc_mapping = get_location_mapping(sorted_list)
    
    st.subheader("📊 當月預覽")
    st.table(pd.DataFrame([{
        "日期": d["日期"], "內容": d["內容"], "金額": f"{d['金額']:,}", "工地代碼": loc_mapping[d["工地"]]
    } for d in sorted_list]))

    word_file = export_word(sorted_list, loc_mapping)
    st.download_button(
        label="📥 下載 A4 簡潔版報表",
        data=word_file,
        file_name=f"雜支明細表_{datetime.now().strftime('%m%d')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
