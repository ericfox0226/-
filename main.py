import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# 頁面配置
st.set_page_config(page_title="公司零用金系統", layout="centered")

# 自定義 CSS 讓手機按鈕更大更好按
st.markdown("""
    <style>
    div.stButton > button:first-child {
        width: 100%;
        height: 3em;
        font-size: 18px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("📂 雜支明細表自動化系統")
st.write("輸入資料後，系統會自動生成符合 A4 排版的 Word 報表。")

# 初始化資料儲存
if 'data_list' not in st.session_state:
    st.session_state.data_list = []

# --- 輸入區塊 ---
with st.container():
    st.subheader("🖋️ 資料輸入")
    date_val = st.text_input("日期", placeholder="例如: 11月18日")
    content_val = st.text_input("花費內容", placeholder="例如: 午餐")
    
    col_a, col_b = st.columns(2)
    with col_a:
        amount_val = st.number_input("金額", min_value=0, step=1, value=0)
    with col_b:
        location_val = st.text_input("工地", placeholder="例如: H")

    if st.button("➕ 新增至清單"):
        if date_val and content_val:
            st.session_state.data_list.append({
                "日期": date_val,
                "內容": content_val,
                "金額": amount_val,
                "工地": location_val
            })
            st.success("已新增一筆！")
        else:
            st.warning("請填寫日期與內容")

# --- 顯示與計算 ---
if st.session_state.data_list:
    st.subheader("📊 當月預覽")
    df = pd.DataFrame(st.session_state.data_list)
    st.table(df)
    
    total = sum(d['金額'] for d in st.session_state.data_list)
    st.info(f"### 目前累計總額：**{total:,}** 元")

    if st.button("🗑️ 全部清空"):
        st.session_state.data_list = []
        st.rerun()

# --- 生成 Word 邏輯 ---
def export_word(data):
    doc = Document()
    
    # 設定 A4 邊距
    section = doc.sections[0]
    section.page_height = Mm(297)
    section.page_width = Mm(210)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)

    # 標題
    p = doc.add_paragraph("雜支明細表")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].font.size = Pt(20)
    p.runs[0].bold = True
    
    doc.add_paragraph("經手人：")

    # 建立 8 欄表格
    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    
    # 表頭設定
    headers = ["日期", "內容", "金額", "工地"] * 2
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(10)

    # 雙欄資料填入
    total_amt = 0
    for i in range(0, len(data), 2):
        row_cells = table.add_row().cells
        # 左側 (0-3欄)
        item_l = data[i]
        row_cells[0].text = item_l["日期"]
        row_cells[1].text = item_l["內容"]
        row_cells[2].text = str(item_l["金額"])
        row_cells[3].text = item_l["工地"]
        total_amt += item_l["金額"]
        
        # 右側 (4-7欄)
        if i + 1 < len(data):
            item_r = data[i+1]
            row_cells[4].text = item_r["日期"]
            row_cells[5].text = item_r["內容"]
            row_cells[6].text = str(item_r["金額"])
            row_cells[7].text = item_r["工地"]
            total_amt += item_r["金額"]

    doc.add_paragraph(f"\n總計金額：{total_amt:,} 元")
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- 下載按鈕 ---
if st.session_state.data_list:
    word_file = export_word(st.session_state.data_list)
    st.download_button(
        label="🚀 下載 Word 報表 (.docx)",
        data=word_file,
        file_name="公司雜支明細表.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
