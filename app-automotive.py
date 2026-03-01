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
        st.error("❌ ไม่พบข้อมูล Secrets: กรุณาเพิ่ม [gcp_service_account] ใน Streamlit Cloud")
        st.stop()
    creds_info = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)

def update_google_sheet(gid, input_df, month, year, is_export=False):
    gc = get_gspread_client()
    sh = gc.open_by_key(SHEET_ID)
    worksheet = sh.get_worksheet_by_id(int(gid))
    
    # 1. ดึงข้อมูลเดิม และล้างค่าว่างทิ้งทันที
    existing_df = get_as_dataframe(worksheet).dropna(how='all').dropna(axis=1, how='all')
    
    # 2. ทำความสะอาดข้อมูล Input
    input_df['Month'] = input_df['Month'].astype(str).str.strip()
    input_df['Year'] = input_df['Year'].astype(str).str.strip()
    
    if not existing_df.empty:
        # ล้างช่องว่างในข้อมูลเก่าด้วย (ป้องกัน ValueError ตอนจัดเรียง)
        existing_df['Month'] = existing_df['Month'].astype(str).str.strip()
        existing_df['Year'] = existing_df['Year'].astype(str).str.strip()
        
        # 3. STRICT OVERWRITE: กรองข้อมูลเก่าทิ้งหากเดือนและปีตรงกัน
        if is_export:
            # สำหรับหมวด 1.3 เช็ค Region ด้วย
            mask = ~((existing_df['Month'] == str(month)) & 
                     (existing_df['Year'] == str(year)) & 
                     (existing_df['Region'].isin(input_df['Region'])))
        else:
            mask = ~((existing_df['Month'] == str(month)) & 
                     (existing_df['Year'] == str(year)))
        
        existing_df = existing_df[mask]

    # 4. รวมข้อมูล
    updated_df = pd.concat([existing_df, input_df], ignore_index=True)
    
    # 5. การจัดเรียง (Sorting) - แก้ไขจุดที่ทำให้เกิด ValueError
    # แปลงปีเป็นตัวเลขแบบปลอดภัย (ถ้าแปลงไม่ได้ให้เป็น 0)
    updated_df['Year_tmp'] = pd.to_numeric(updated_df['Year'], errors='coerce').fillna(0).astype(int)
    
    # ตั้งค่าลำดับเดือน
    updated_df['Month'] = pd.Categorical(updated_df['Month'], categories=MONTH_ORDER, ordered=True)
    
    if is_export:
        updated_df['Region'] = pd.Categorical(updated_df['Region'], categories=REGIONS_LIST, ordered=True)
        updated_df = updated_df.sort_values(by=['Year_tmp', 'Month', 'Region'])
    else:
        updated_df = updated_df.sort_values(by=['Year_tmp', 'Month'])

    # ลบคอลัมน์ชั่วคราวทิ้ง
    updated_df = updated_df.drop(columns=['Year_tmp'])

    # 6. บันทึกกลับแบบล้างแผ่นงานก่อน (ป้องกันข้อมูลค้าง)
    worksheet.clear()
    set_with_dataframe(worksheet, updated_df)
    return True

# --- UI ---
st.set_page_config(page_title="Data Management", layout="wide")
st.title("📑 ระบบจัดการข้อมูล (แก้ไขปัญหาการเรียงลำดับและการซ้ำ)")

with st.sidebar:
    st.header("📅 ช่วงเวลา")
    sel_month = st.selectbox("เดือน", MONTH_ORDER)
    sel_year = st.text_input("ปี (พ.ศ.)", value="2568")
    st.divider()
    category = st.radio("เลือกหัวข้อ:", list(GID_MAP.keys()))

# --- หมวด 1.3 ---
if category == "1.3 ยอดส่งออกรถยนต์":
    st.subheader(f"📍 {category}")
    raw_paste = st.text_area("วางข้อมูลแนวตั้งจาก Excel", height=200)
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
        st.write("🔍 ข้อมูลที่จะบันทึก (จะทับค่าเดิมในระบบ)")
        edited_final = st.data_editor(final_df, use_container_width=True)
        if st.button("🚀 ยืนยันการบันทึกทับ"):
            if update_google_sheet(GID_MAP[category], edited_final, sel_month, sel_year, is_export=True):
                st.success(f"บันทึกข้อมูลเดือน {sel_month} เรียบร้อยและจัดลำดับใหม่แล้ว")

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
    if st.button(f"💾 บันทึกทับข้อมูล {sel_month}"):
        if update_google_sheet(GID_MAP[category], edited_df, sel_month, sel_year):
            st.success(f"บันทึก {sel_month} สำเร็จ!")



