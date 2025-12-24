import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Mm
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
        background-color: #f8f9fa; padding: 20px; border-radius: 10px; 
        text-align: center; border: 2px solid #343a40; margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 初始化 Session State ---
if 'data_list' not in st.session_state:
    st.session_state.data_list = []
if 'location_options' not in st.session_state:
    # 預設一些常用工地選項
    st.session_state.location_options = ["A工地", "B中心", "C住宅"]

# --- 側邊欄：選項管理 (新增/刪減) ---
with st.sidebar:
    st.header("⚙️ 選項管理")
    st.subheader("工地清單")
    
    # 新增選項
    new_loc = st.text_input("新增工地名稱", placeholder="例如：台北大巨蛋")
    if st.button("➕ 增加至選單"):
        if new_loc and new_loc not in st.session_state.location_options:
            st.session_state.location_options.append(new_loc)
            st.rerun()
            
    st.divider()
    
    # 刪除選項
    del_loc = st.selectbox("選擇要刪除的工地", options=st.session_state.location_options)
    if st.button("🗑️ 刪除該選項"):
        if del_loc in st.session_state.location_options:
            st.session_state.location_options.remove(del_loc)
            st.rerun()

# --- 主頁面：總金額預覽 ---
st.title("📂 雜支明細表自動化系統")

if st.session_state.data_list:
    total_amt = sum(d['金額'] for d in st.session_state.data_list)
    text_color = "#d32f2f" if total_amt < 0 else "#01579b"
    st.markdown(f"""
        <div class="total-preview">
            <span style="font-size: 16px; color: #666;">目前累計總餘額</span><br>
            <span style="font-size: 32px; font-weight: bold; color: {text_color};">NT$ {total_amt:,}</span>
        </div>
    """, unsafe_allow_html=True)

# --- 輸入區塊 ---
with st.expander("🖋️ 新增資料", expanded=True):
    today_str = datetime.now().strftime("%m/%d")
    date_val = st.text_input("日期", value=today_str)
    content_val = st.text_input("花費內容", placeholder="例如：五金零件、便當")
    
    col_a, col_b = st.columns(2)
    with col_a:
        raw_amount = st.number_input("支出金額 (自動轉負數)", step=1, value=0)
    with col_b:
        # 使用下拉選單選擇工地
        selected_loc = st.selectbox("選擇工地", options=st.session_state.location_options + ["+ 手動輸入"])
        
        # 如果選擇手動輸入，顯示輸入框
        if selected_loc == "+ 手動輸入":
            final_location = st.text_input("請輸入新工地名稱")
        else:
            final_location = selected_loc

    if st.button("➕ 新增至清單"):
        if date_val and content_val and final_location:
            # 支出預設轉負數邏輯
            actual_amount = -abs(raw_amount) if raw_amount > 0 else raw_amount
            st.session_state.data_list.append({
                "日期": date_val, "內容": content_val, "金額": actual_amount, "工地": final_location
            })
            st.rerun()
        else:
            st.warning("請填寫完整資訊")

# --- 排序與 Word 生成 (維持之前優化的垂直排列與置中邏輯) ---
def get_sorted_data(data):
    def sort_key(item):
        try: return datetime.strptime(f"{datetime.now().year}/{item['日期']}", "%Y/%m/%d")
        except: return datetime.max
    return sorted(data, key=sort_key)

def get_location_mapping(sorted_data):
    unique_locations = []
    for d in sorted_data:
        if d["工地"] not in unique_locations:
            unique_locations.append(d["工地"])
    return {loc: chr(65 + i) for i, loc in enumerate(unique_locations)}

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

    rows_per_page = 28 
    left_side = data[:rows_per_page]
    right_side = data[rows_per_page:rows_per_page*2]

    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    headers = ["日期", "內容", "金額", "工地代號"] * 2
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.paragraphs[0].runs[0].font.bold = True

    last_d_l, last_d_r = None, None
    for i in range(len(left_side)):
        row_cells = table.add_row().cells
        d_l = left_side[i]
        show_date_l = "" if d_l["日期"] == last_d_l else d_l["日期"]
        last_d_l = d_l["日期"]
        l_vals = [show_date_l, d_l["內容"], f"{d_l['金額']:,}", mapping[d_l["工地"]]]
        
        r_vals = ["", "", "", ""]
        if i < len(right_side):
            d_r = right_side[i]
            show_date_r = "" if d_r["日期"] == last_d_r else d_r["日期"]
            last_d_r = d_r["日期"]
            r_vals = [show_date_r, d_r["內容"], f"{d_r['金額']:,}", mapping[d_r["工地"]]]

        for idx, val in enumerate(l_vals + r_vals):
            cell = row_cells[idx]
            cell.text = str(val)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    total = sum(d['金額'] for d in data)
    doc.add_paragraph(f"\n總計金額：NT$ {total:,} 元").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph("-" * 20 + "\n【工地代號索引】").bold = True
    for name, code in mapping.items():
        doc.add_paragraph(f"{code} : {name}")

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- 下載與資料列表 ---
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
        file_name=f"雜支明細表_{datetime.now().strftime('%m%d')}.docx"
    )
