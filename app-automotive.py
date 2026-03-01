import streamlit as st
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe, get_as_dataframe
from google.oauth2.service_account import Credentials

# --- CONFIGURATION ---
SHEET_ID = "17Nq4MVLOKtdantiDayXwAgPRZKCvkI1FD4n7FJMZlJo"

# Mapping GID (Sheet ID ย่อย) ตามที่คุณระบุ
GID_MAP = {
    "1.1 ยอดผลิตรถยนต์": "446114989",
    "1.2 ยอดจำหน่ายรถยนต์": "0",
    "1.3 ยอดส่งออกรถยนต์": "1329648063",
    "2.1 ยอดผลิตจักรยานยนต์": "2053730588",
    "2.2 ยอดจำหน่ายจักรยานยนต์": "787611063",
    "2.3 ยอดส่งออกจักรยานยนต์": "77561054"
}

# รายชื่อภูมิภาคสำหรับยอดส่งออก (1.3) ตามรูปแนบที่ 1
REGIONS_LIST = [
    "Asia", "Australia, NZ & Other Oceania", "Middle East", 
    "Africa", "Europe", "North America", "Central & South America", "Others"
]

# --- CORE FUNCTIONS ---
def get_gspread_client():
    """เชื่อมต่อ Google Sheets โดยใช้ Key จาก Secrets"""
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds_info = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)
# บรรทัดที่ 34 โดยประมาณ
def update_google_sheet(gid, input_df, month, year, is_export=False):
    gc = get_gspread_client()
    sh = gc.open_by_key(SHEET_ID) # บรรทัดที่ 36
    # บรรทัดที่ 38: แก้ไขเป็นแบบนี้
    worksheet = sh.get_worksheet_by_id(int(gid)) 
    
    # 1. ดึงข้อมูลเดิมจาก Google Sheet
    existing_df = get_as_dataframe(worksheet).dropna(how='all').dropna(axis=1, how='all')
    # ... บรรทัดต่อๆ ไปเหมือนเดิม ...
    
    # ดึงข้อมูลเดิม
    existing_df = get_as_dataframe(worksheet).dropna(how='all').dropna(axis=1, how='all')
    
    # ลบข้อมูลเก่าที่มี Month/Year ตรงกันออกก่อน (Overwrite)
    if not existing_df.empty:
        existing_df['Month'] = existing_df['Month'].astype(str)
        existing_df['Year'] = existing_df['Year'].astype(str)
        mask = ~((existing_df['Month'] == month) & (existing_df['Year'] == year))
        existing_df = existing_df[mask]

    # รวมข้อมูล
    updated_df = pd.concat([existing_df, input_df], ignore_index=True)
    
    # จัดเรียงลำดับ (Sorting)
    month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    updated_df['Month'] = pd.Categorical(updated_df['Month'], categories=month_order, ordered=True)
    
    if is_export:
        updated_df['Region'] = pd.Categorical(updated_df['Region'], categories=REGIONS_LIST, ordered=True)
        updated_df = updated_df.sort_values(by=['Year', 'Region', 'Month'])
    else:
        updated_df = updated_df.sort_values(by=['Year', 'Month'])

    # บันทึกกลับลง Google Sheet
    set_with_dataframe(worksheet, updated_df)
    return True

# --- UI INTERFACE ---
st.set_page_config(page_title="Auto Data Entry", layout="wide")
st.title("🚀 ระบบจัดการข้อมูลยานยนต์ (Google Sheets)")

# ส่วนตั้งค่าด้านข้าง
with st.sidebar:
    st.header("📅 กำหนดข้อมูล")
    sel_month = st.selectbox("Month", ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    sel_year = st.text_input("Year (พ.ศ.)", value="2568")
    st.divider()
    category = st.radio("เลือกหัวข้อที่ต้องการกรอก:", list(GID_MAP.keys()))

# --- หมวดหมู่ 1.3 ยอดส่งออก (Special Logic) ---
if category == "1.3 ยอดส่งออกรถยนต์":
    st.subheader(f"📍 {category}")
    st.info("💡 ขั้นตอน: ก๊อปปี้คอลัมน์ตัวเลขจาก Excel (แนวตั้ง 1 คอลัมน์) ตามรูปแนบ 1 แล้ววางในช่องด้านล่าง")
    
    raw_paste = st.text_area("วางข้อมูลจาก Excel ที่นี่", height=300, placeholder="ตัวอย่าง:\n8695\n8766\n2483\n0\n...")
    
    if raw_paste:
        # ประมวลผลข้อความที่วางให้กลายเป็น DataFrame
        lines = raw_paste.strip().split('\n')
        vals = []
        for l in lines:
            v = l.replace(',', '').strip()
            vals.append(float(v) if v not in ['-', '', ' '] else 0)

        # Map ข้อมูลจากแนวตั้ง (รูป 1) ไปเป็นแนวนอน (รูป 2)
        rows = []
        v_idx = 0
        for reg in REGIONS_LIST:
            # ดึงข้อมูล 4 ประเภทต่อภูมิภาค (Pickup, Passenger, PPV, Truck)
            data_chunk = vals[v_idx : v_idx+4] if v_idx+4 <= len(vals) else [0,0,0,0]
            rows.append({
                "Month": sel_month, "Year": sel_year, "Region": reg,
                "Pickup": data_chunk[0], "Passenger": data_chunk[1],
                "PPV": data_chunk[2], "Truck": data_chunk[3],
                "Total_region": sum(data_chunk)
            })
            v_idx += 5 # ขยับข้ามไป 5 บรรทัด (เพราะใน Excel มีบรรทัด Sub Total คั่น)
            
        final_df = pd.DataFrame(rows)
        st.write("### 🔍 ตรวจสอบข้อมูลก่อนบันทึก (จัดรูปแบบตามรูปแนบ 2)")
        st.dataframe(final_df, use_container_width=True)
        
        if st.button("🚀 บันทึกข้อมูลส่งออก"):
            if update_google_sheet(GID_MAP[category], final_df, sel_month, sel_year, is_export=True):
                st.success("บันทึกและจัดเรียงข้อมูลส่งออกเรียบร้อย!")

# --- หมวดหมู่อื่นๆ (Manual Entry Table) ---
else:
    st.subheader(f"📊 {category}")
    
    # กำหนดคอลัมน์ตามประเภท
    col_config = {
        "1.1 ยอดผลิตรถยนต์": ["Passenger", "Pickup", "Commercial"],
        "1.2 ยอดจำหน่ายรถยนต์": ["Passenger", "Pickup", "Commercial", "PPV_SUV"],
        "2.1 ยอดผลิตจักรยานยนต์": ["Family", "Sport", "EV", "ICONIC"],
        "2.2 ยอดจำหน่ายจักรยานยนต์": ["< 50 CC", "51-110 CC", "111-125 CC", "126-250 CC", "251-399 CC", "< 400 CC"],
        "2.3 ยอดส่งออกจักรยานยนต์": ["CBU", "CKD", "Value"]
    }
    
    # สร้างตารางให้กรอกหรือวางข้อมูล
    current_cols = col_config[category]
    init_df = pd.DataFrame([{"Month": sel_month, "Year": sel_year, **{c: 0 for c in current_cols}}])
    
    edited_df = st.data_editor(init_df, num_rows="dynamic", hide_index=True)
    
    # คำนวณ Total
    numeric_cols = [c for c in edited_df.columns if c not in ["Month", "Year"]]
    edited_df["Total"] = edited_df[numeric_cols].sum(axis=1)

    if st.button(f"💾 บันทึก {category}"):
        if update_google_sheet(GID_MAP[category], edited_df, sel_month, sel_year):
            st.success("บันทึกข้อมูลเรียบร้อย!")



