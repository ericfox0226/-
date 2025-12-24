import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Mm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from io import BytesIO
from datetime import datetime

# --- 1. 頁面配置與樣式 ---
st.set_page_config(page_title="工地雜支管理系統", layout="centered")

st.markdown("""
    <style>
    div.stButton > button:first-child { width: 100%; height: 3.5em; font-size: 18px; font-weight: bold; }
    .total-preview { 
        background-color: #f8f9fa; padding: 20px; border-radius: 15px; 
        text-align: center; border: 2px solid #343a40; margin-bottom: 25px;
    }
    </style>
    """, unsafe_allow_html=True)

if 'data_list' not in st.session_state:
    st.session_state.data_list = []

# --- 2. 頂部總金額預覽 (UI 加強) ---
st.title("📂 雜支明細表自動化系統")

if st.session_state.data_list:
    total_amt = sum(d['金額'] for d in st.session_state.data_list)
    text_color = "#d32f2f" if total_amt < 0 else "#01579b"
    st.markdown(f"""
        <div class="total-preview">
            <span style="font-size: 16px; color: #666;">目前累計總預算餘額</span><br>
            <span style="font-size: 32px; font-weight: bold; color: {text_color};">NT$ {total_amt:,}</span>
        </div>
    """, unsafe_allow_html=True)

# --- 3. 輸入區塊 (邏輯：預設支出為負) ---
with st.expander("🖋️ 快速新增資料", expanded=True):
    today_str = datetime.now().strftime("%m/%d")
    date_val = st.text_input("日期", value=today_str)
    content_val = st.text_input("花費內容", placeholder="如：水泥、午餐費")
    
    col_a, col_b = st.columns(2)
    with col_a:
        # 使用者輸入 500，系統存入 -500；若要存入正數，請手動輸入 -500
        raw_amount = st.number_input("金額 (輸入數字即為支出)", step=1, value=0)
    with col_b:
        location_val = st.text_input("工地全名", placeholder="如：台北大巨蛋")

    if st.button("➕ 新增至清單"):
        if date_val and content_val and location_val:
            # 自動轉負數邏輯
            actual_amount = -abs(raw_amount) if raw_amount > 0 else raw_amount
            st.session_state.data_list.append({
                "日期": date_val, "內容": content_val, "金額": actual_amount, "工地": location_val
            })
            st.rerun()

# --- 4. 排序與自動代號生成邏輯 ---
def process_data(data):
    # 依日期排序
    def sort_key(item):
        try: return datetime.strptime(f"{datetime.now().year}/{item['日期']}", "%Y/%m/%d")
        except: return datetime.max
    sorted_data = sorted(data, key=sort_key)
    
    # 生成 A-Z 代號 Mapping
    unique_locs = []
    for d in sorted_data:
        if d["工地"] not in unique_locs:
            unique_locs.append(d["工地"])
    mapping = {loc: chr(65 + i) for i, loc in enumerate(unique_locs)}
    
    return sorted_data, mapping

# --- 5. Word 生成邏輯 (垂直排列 + 置中 + 同日期隱藏) ---
def export_word(data, mapping):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Mm(15)
    section.left_margin = section.right_margin = Mm(15)

    title = doc.add_paragraph("雜支明細表")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.size = Pt(18)
    title.runs[0].bold = True
    
    doc.add_paragraph(f"報告日期：{datetime.now().strftime('%Y/%m/%d')}")
    doc.add_paragraph(f"經手人：_________________")

    # 分配資料：左側填滿 28 列後再填右側
    rows_limit = 28 
    left_part = data[:rows_limit]
    right_part = data[rows_limit:rows_limit*2]

    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    
    # 表頭
    headers = ["日期", "內容", "金額", "工地代號"] * 2
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.paragraphs[0].runs[0].font.bold = True

    last_l_date, last_r_date = None, None
    for i in range(len(left_part)):
        row = table.add_row().cells
        
        # 左側填值
        d_l = left_part[i]
        date_l = "" if d_l["日期"] == last_l_date else d_l["日期"]
        last_l_date = d_l["日期"]
        l_vals = [date_l, d_l["內容"], f"{d_l['金額']:,}", mapping[d_l["工地"]]]
        
        # 右側填值
        r_vals = [""] * 4
        if i < len(right_part):
            d_r = right_part[i]
            date_r = "" if d_r["日期"] == last_r_date else d_r["日期"]
            last_r_date = d_r["日期"]
            r_vals = [date_r, d_r["內容"], f"{d_r['金額']:,}", mapping[d_r["工地"]]]

        # 套用格式
        for idx, val in enumerate(l_vals + r_vals):
            row[idx].text = str(val)
            row[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row[idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # 結尾總計與索引
    doc.add_paragraph(f"\n總計金額：NT$ {sum(d['金額'] for d in data):,} 元").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph("-" * 20 + "\n【工地代號索引對照】").bold = True
    for name, code in mapping.items():
        doc.add_paragraph(f"{code} : {name}")

    out = BytesIO(); doc.save(out); out.seek(0)
    return out

# --- 6. 顯示預覽與下載 ---
if st.session_state.data_list:
    sorted_list, loc_map = process_data(st.session_state.data_list)
    
    st.subheader("📊 本月明細預覽 (已依日期排序)")
    st.table(pd.DataFrame([{
        "日期": d["日期"], "內容": d["內容"], "金額": d["金額"], "代號": loc_map[d["工地"]]
    } for d in sorted_list]))

    col1, col2 = st.columns(2)
    with col1:
        if st.button("⏪ 刪除最後一筆"):
            st.session_state.data_list.pop(); st.rerun()
    with col2:
        if st.button("🗑️ 全部清空"):
            st.session_state.data_list = []; st.rerun()

    word_file = export_word(sorted_list, loc_map)
    st.download_button(
        label="📥 下載 A4 專業報表",
        data=word_file,
        file_name=f"雜支明細_{datetime.now().strftime('%m%d')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
