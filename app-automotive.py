import streamlit as st
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe, get_as_dataframe
from google.oauth2.service_account import Credentials

# --- CONFIGURATION ---
SHEET_ID = "17Nq4MVLOKtdantiDayXwAgPRZKCvkI1FD4n7FJMZlJo"
GID_MAP = {
    "1.1 ยอดผลิตรถยนต์": "446114989",
    "1.2 ยอดจำหน่ายรถยนต์": "0",
    "1.3 ยอดส่งออกรถยนต์": "1329648063",
    "2.1 ยอดผลิตจักรยานยนต์": "2053730588",
    "2.2 ยอดจำหน่ายจักรยานยนต์": "787611063",
    "2.3 ยอดส่งออกจักรยานยนต์": "77561054"
}
REGIONS_LIST = ["Asia", "Australia, NZ & Other Oceania", "Middle East", "Africa", "Europe", "North America", "Central & South America", "Others"]
MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def get_gspread_client():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    # ตรวจสอบชื่อ Key ใน Secrets ให้ตรงกับภาพ image_ca721b.png
    if "gcp_service_account" in st.secrets:
        creds_info = st.secrets["gcp_service_account"]
    else:
        # กรณีเก็บแบบกระจายตัว (Flat)
        creds_info = dict(st.secrets)
    
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)

def update_google_sheet(gid, input_df, month, year, is_export=False):
    gc = get_gspread_client()
    sh = gc.open_by_key(SHEET_ID)
    worksheet = sh.get_worksheet_by_id(int(gid))
    
    # 1. ดึงข้อมูลปัจจุบันและทำความสะอาด
    existing_df = get_as_dataframe(worksheet).dropna(how='all').dropna(axis=1, how='all')
    
    # บังคับให้เป็น String และตัดช่องว่าง
    input_df['Month'] = input_df['Month'].astype(str).str.strip()
    input_df['Year'] = input_df['Year'].astype(str).str.strip()
    target_month = str(month).strip()
    target_year = str(year).strip()

    if not existing_df.empty:
        existing_df['Month'] = existing_df['Month'].astype(str).str.strip()
        existing_df['Year'] = existing_df['Year'].astype(str).str.strip()

        # 2. OVERWRITE LOGIC: กรองข้อมูลที่ "ไม่ใช่" เดือน/ปี ที่กำลังบันทึกเก็บไว้
        if is_export:
            # สำหรับ Export ต้องเช็ค Region ด้วย เพื่อไม่ให้ไปทับภูมิภาคอื่นในเดือนเดียวกัน
            mask = ~((existing_df['Month'] == target_month) & 
                     (existing_df['Year'] == target_year) & 
                     (existing_df['Region'].isin(input_df['Region'].unique())))
        else:
            mask = ~((existing_df['Month'] == target_month) & 
                     (existing_df['Year'] == target_year))
        
        existing_df = existing_df[mask]

    # 3. รวมข้อมูลใหม่ (Concat)
    updated_df = pd.concat([existing_df, input_df], ignore_index=True)

    # 4. จัดเรียงลำดับ (Safe Sort)
    # สร้างคอลัมน์ชั่วคราวสำหรับเรียงปี (ป้องกัน ValueError)
    updated_df['Year_tmp'] = pd.to_numeric(updated_df['Year'], errors='coerce').fillna(0).astype(int)
    updated_df['Month'] = pd.Categorical(updated_df['Month'], categories=MONTH_ORDER, ordered=True)
    
    if is_export:
        updated_df['Region'] = pd.Categorical(updated_df['Region'], categories=REGIONS_LIST, ordered=True)
        updated_df = updated_df.sort_values(by=['Year_tmp', 'Month', 'Region'])
    else:
        updated_df = updated_df.sort_values(by=['Year_tmp', 'Month'])

    # ลบคอลัมน์ชั่วคราว
    updated_df = updated_df.drop(columns=['Year_tmp'])

    # 5. เขียนทับลง Sheet (Clear & Write)
    worksheet.clear()
    set_with_dataframe(worksheet, updated_df)
    return True

# --- UI (Streamlit) ---
st.set_page_config(page_title="Data Entry Fix", layout="wide")
st.title("📌 บันทึกข้อมูลแบบเขียนทับ (Force Overwrite)")

with st.sidebar:
    sel_month = st.selectbox("เดือน", MONTH_ORDER)
    sel_year = st.text_input("ปี (พ.ศ.)", value="2568")
    category = st.radio("หัวข้อ:", list(GID_MAP.keys()))

# --- ประมวลผลแต่ละหมวด ---
if category == "1.3 ยอดส่งออกรถยนต์":
    raw_paste = st.text_area("วางข้อมูลจาก Excel", height=200)
    if raw_paste:
        lines = raw_paste.strip().split('\n')
        vals = [float(l.replace(',', '').strip()) if l.strip() not in ['-', '', ' '] else 0 for l in lines]
        rows = []
        idx = 0
        for reg in REGIONS_LIST:
            chunk = vals[idx : idx+5] if idx+5 <= len(vals) else [0,0,0,0,0]
            rows.append({"Month": sel_month, "Year": sel_year, "Region": reg, "Pickup": chunk[0], "Passenger": chunk[1], "PPV": chunk[2], "Truck": chunk[3], "Total_region": chunk[4]})
            idx += 5
        df = pd.DataFrame(rows)
        if st.button("🚀 บันทึกทับข้อมูลเดือน " + sel_month):
            if update_google_sheet(GID_MAP[category], df, sel_month, sel_year, is_export=True):
                st.success("สำเร็จ! ข้อมูลเก่าถูกแทนที่แล้ว")
else:
    # หมวดอื่นๆ 1.1, 1.2, 2.1...
    col_dict = {
        "1.1 ยอดผลิตรถยนต์": ["Passenger", "Pickup", "Commercial", "Total"],
        "1.2 ยอดจำหน่ายรถยนต์": ["Passenger", "Pickup", "Commercial", "PPV_SUV", "Total"]
    }
    cols = col_dict.get(category, ["Data"])
    input_data = pd.DataFrame([{"Month": sel_month, "Year": sel_year, **{c: 0 for c in cols}}])
    edited_df = st.data_editor(input_data, hide_index=True)
    if st.button("💾 บันทึกทับ"):
        if update_google_sheet(GID_MAP[category], edited_df, sel_month, sel_year):
            st.success("บันทึกทับข้อมูล " + sel_month + " เรียบร้อย")



