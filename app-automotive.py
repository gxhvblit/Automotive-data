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

REGIONS_LIST = [
    "Asia", "Australia, NZ & Other Oceania", "Middle East", 
    "Africa", "Europe", "North America", "Central & South America", "Others"
]

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# --- FUNCTIONS ---
def get_gspread_client():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    if "gcp_service_account" not in st.secrets:
        st.error("❌ ไม่พบข้อมูล Secrets กรุณาตั้งค่าให้ถูกต้อง")
        st.stop()
    creds_info = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)

def update_google_sheet(gid, input_df, month, year, is_export=False):
    gc = get_gspread_client()
    sh = gc.open_by_key(SHEET_ID)
    worksheet = sh.get_worksheet_by_id(int(gid))
    
    # 1. ดึงข้อมูลเดิมทั้งหมด
    existing_df = get_as_dataframe(worksheet).dropna(how='all').dropna(axis=1, how='all')
    
    # 2. เตรียมข้อมูลใหม่ให้เป็น String เพื่อป้องกันการเทียบ Type ผิดพลาด
    input_df['Month'] = input_df['Month'].astype(str)
    input_df['Year'] = input_df['Year'].astype(str)
    
    if not existing_df.empty:
        existing_df['Month'] = existing_df['Month'].astype(str)
        existing_df['Year'] = existing_df['Year'].astype(str)
        
        # 3. OVERWRITE LOGIC: ลบข้อมูลเก่าที่ "เดือน" และ "ปี" ตรงกับที่เราจะบันทึกออกให้หมด
        if is_export:
            # สำหรับหมวด 1.3/2.3 เช็ค Region ด้วย
            mask = ~((existing_df['Month'] == str(month)) & 
                     (existing_df['Year'] == str(year)) & 
                     (existing_df['Region'].isin(input_df['Region'])))
        else:
            # สำหรับหมวดปกติ ลบทุกแถวที่มี Month/Year ตรงกัน
            mask = ~((existing_df['Month'] == str(month)) & 
                     (existing_df['Year'] == str(year)))
        
        existing_df = existing_df[mask]

    # 4. รวมข้อมูลใหม่เข้ากับข้อมูลเก่าที่เหลือ
    updated_df = pd.concat([existing_df, input_df], ignore_index=True)
    
    # 5. STRICT SORTING: บังคับลำดับให้ถูกต้อง
    # แปลง Year เป็นตัวเลขชั่วคราวเพื่อให้เรียง 2567 ก่อน 2568
    updated_df['Year_int'] = updated_df['Year'].astype(int)
    # กำหนดลำดับเดือน Jan-Dec
    updated_df['Month'] = pd.Categorical(updated_df['Month'], categories=MONTH_ORDER, ordered=True)
    
    if is_export:
        updated_df['Region'] = pd.Categorical(updated_df['Region'], categories=REGIONS_LIST, ordered=True)
        updated_df = updated_df.sort_values(by=['Year_int', 'Month', 'Region'])
    else:
        updated_df = updated_df.sort_values(by=['Year_int', 'Month'])

    # ลบ Column ช่วยเรียงลำดับทิ้งก่อนบันทึก
    updated_df = updated_df.drop(columns=['Year_int'])

    # 6. บันทึกกลับลง Sheet (Clear & Write)
    worksheet.clear()
    set_with_dataframe(worksheet, updated_df)
    return True

# --- UI ---
st.set_page_config(page_title="Automotive Data Entry Pro", layout="wide")
st.title("🚀 ระบบบันทึกข้อมูลยานยนต์ (V3: Strict Overwrite)")

with st.sidebar:
    st.header("📅 เลือกช่วงเวลา")
    sel_month = st.selectbox("เดือน", MONTH_ORDER)
    sel_year = st.text_input("ปี (พ.ศ.)", value="2568")
    st.divider()
    category = st.radio("หัวข้อ:", list(GID_MAP.keys()))

# --- หมวด 1.3 (Paste แนวตั้ง) ---
if category == "1.3 ยอดส่งออกรถยนต์":
    st.subheader(f"📍 {category}")
    raw_paste = st.text_area("วางข้อมูลคอลัมน์จาก Excel", height=200)
    if raw_paste:
        lines = raw_paste.strip().split('\n')
        vals = [float(l.replace(',', '').strip()) if l.strip() not in ['-', '', ' '] else 0 for l in lines]
        rows = []
        idx = 0
        for reg in REGIONS_LIST:
            chunk = vals[idx : idx+5] if idx+5 <= len(vals) else [0,0,0,0,0]
            rows.append({
                "Month": sel_month, "Year": sel_year, "Region": reg,
                "Pickup": chunk[0], "Passenger": chunk[1], "PPV": chunk[2], "Truck": chunk[3],
                "Total_region": chunk[4]
            })
            idx += 5
        final_df = pd.DataFrame(rows)
        st.write("🔍 ตรวจสอบ (จะบันทึกทับข้อมูลเดิมทันที)")
        edited_final = st.data_editor(final_df, use_container_width=True)
        if st.button("🚀 บันทึกข้อมูลและจัดเรียงใหม่"):
            if update_google_sheet(GID_MAP[category], edited_final, sel_month, sel_year, is_export=True):
                st.success("บันทึกและจัดเรียงข้อมูลเรียบร้อย!")

# --- หมวดอื่นๆ ---
else:
    st.subheader(f"📊 {category}")
    col_config = {
        "1.1 ยอดผลิตรถยนต์": ["Passenger", "Pickup", "Commercial", "Total"],
        "1.2 ยอดจำหน่ายรถยนต์": ["Passenger", "Pickup", "Commercial", "PPV_SUV", "Total"],
        "2.1 ยอดผลิตจักรยานยนต์": ["Family", "Sport", "EV", "ICONIC", "Total"],
        "2.2 ยอดจำหน่ายจักรยานยนต์": ["< 50 CC", "51-110 CC", "111-125 CC", "126-250 CC", "251-399 CC", "< 400 CC", "Total"],
        "2.3 ยอดส่งออกจักรยานยนต์": ["CBU", "CKD", "Value", "Total"]
    }
    cols = col_config[category]
    template_df = pd.DataFrame([{"Month": sel_month, "Year": sel_year, **{c: 0 for c in cols}}])
    edited_df = st.data_editor(template_df, hide_index=True, use_container_width=True)
    if st.button(f"💾 บันทึกทับ {category}"):
        if update_google_sheet(GID_MAP[category], edited_df, sel_month, sel_year):
            st.success("บันทึกสำเร็จ!")



