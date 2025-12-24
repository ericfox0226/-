import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from io import BytesIO
from datetime import datetime

# --- 1. 頁面配置 ---
st.set_page_config(page_title="公司雲端零用金系統", layout="centered")

st.markdown("""
    <style>
    div.stButton > button:first-child { width: 100%; height: 3.5em; font-size: 18px; font-weight: bold; }
    .total-preview { 
        background-color: #f8f9fa; padding: 20px; border-radius: 15px; 
        text-align: center; border: 2px solid #343a40; margin-bottom: 25px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 建立 Google Sheets 連線 ---
# 注意：這需要在 Streamlit Cloud 後台設定 Secrets
conn = st.connection("gsheets", type=GSheetsConnection)

# 讀取現有資料
try:
    existing_data = conn.read(ttl="0s") # ttl=0s 確保每次都抓最新資料
except:
    # 如果是第一次運行或表格是空的，建立空 DataFrame
    existing_data = pd.DataFrame(columns=["日期", "內容", "金額", "工地"])

# --- 3. 頂部總金額預覽 ---
st.title("📂 雲端雜支明細系統")

if not existing_data.empty:
    total_amt = existing_data["金額"].astype(int).sum()
    text_color = "#d32f2f" if total_amt < 0 else "#01579b"
    st.markdown(f"""
        <div class="total-preview">
            <span style="font-size: 16px; color: #666;">雲端同步：目前累計總餘額</span><br>
            <span style="font-size: 32px; font-weight: bold; color: {text_color};">NT$ {total_amt:,}</span>
        </div>
    """, unsafe_allow_html=True)

# --- 4. 輸入區塊 ---
with st.expander("🖋️ 新增雲端帳目", expanded=True):
    date_val = st.text_input("日期", value=datetime.now().strftime("%m/%d"))
    content_val = st.text_input("花費內容")
    col_a, col_b = st.columns(2)
    with col_a:
        raw_amount = st.number_input("金額 (自動轉支出)", step=1, value=0)
    with col_b:
        location_val = st.text_input("工地全名")

    if st.button("🚀 同步至 Google Sheets"):
        if date_val and content_val and location_val:
            actual_amount = -abs(raw_amount) if raw_amount > 0 else raw_amount
            new_row = pd.DataFrame([{
                "日期": date_val, "內容": content_val, "金額": actual_amount, "工地": location_val
            }])
            # 合併新舊資料並寫回 Google Sheets
            updated_df = pd.concat([existing_data, new_row], ignore_index=True)
            conn.update(data=updated_df)
            st.success("資料已成功存入雲端！")
            st.rerun()

# --- 5. 排序、代號與 Word 生成邏輯 (維持您的專業排版) ---
def process_data(df):
    def sort_key(row):
        try: return datetime.strptime(f"{datetime.now().year}/{row['日期']}", "%Y/%m/%d")
        except: return datetime.max
    df['sort_key'] = df.apply(sort_key, axis=1)
    sorted_df = df.sort_values('sort_key').drop(columns=['sort_key'])
    
    unique_locs = sorted_df["工地"].unique().tolist()
    mapping = {loc: chr(65 + i) for i, loc in enumerate(unique_locs)}
    return sorted_df, mapping

if not existing_data.empty:
    sorted_df, loc_map = process_data(existing_data)
    data_list = sorted_df.to_dict('records')

    st.subheader("📊 雲端明細預覽")
    st.table(sorted_df)

    if st.button("🗑️ 清空雲端所有資料"):
        conn.update(data=pd.DataFrame(columns=["日期", "內容", "金額", "工地"]))
        st.rerun()

    # --- 此處省略 export_word 函式，內容與之前相同，僅需將 data 傳入即可 ---
    # (為節省長度，請沿用您前一版本的 export_word 函式內容)
