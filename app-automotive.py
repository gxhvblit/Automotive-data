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

# --- FUNCTIONS ---
def get_gspread_client():
    """เชื่อมต่อ Google Sheets ผ่าน Streamlit Secrets"""
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # ตรวจสอบการตั้งค่า Secrets
    if "gcp_service_account" not in st.secrets:
        st.error("❌ ไม่พบข้อมูล Secrets: กรุณาใส่ [gcp_service_account] ในหน้า Settings > Secrets")
        st.stop()
        
    creds_info = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)

def update_google_sheet(gid, input_df, month, year, is_export=False):
    """บันทึกข้อมูลแบบเขียนทับ (Overwrite) เมื่อเดือน/ปีซ้ำ และจัดเรียงข้อมูล"""
    gc = get_gspread_client()
    sh = gc.open_by_key(SHEET_ID)
    
    # แก้ไข AttributeError: เรียกใช้ worksheet โดยตรงจาก spreadsheet object (sh)
    worksheet = sh.get_worksheet_by_id(int(gid))
    
    # ดึงข้อมูลเดิมเพื่อนำมา Overwrite
    existing_df = get_as_dataframe(worksheet).dropna(how='all').dropna(axis=1, how='all')
    
    if not existing_df.empty:
        # แปลงเป็น string เพื่อป้องกันปัญหาการเปรียบเทียบข้อมูล
        existing_df['Month'] = existing_df['Month'].astype(str)
        existing_df['Year'] = existing_df['Year'].astype(str)
        
        # กรองข้อมูลเก่า: เลือกเก็บเฉพาะแถวที่ "ไม่ใช่" เดือนและปีที่กำลังบันทึก (Overwrite Logic)
        if is_export:
            # กรณี 1.3: เช็ค Month + Year + Region (เผื่อกรณีต้องการแก้รายภูมิภาค)
            mask = ~((existing_df['Month'] == str(month)) & 
                     (existing_df['Year'] == str(year)) & 
                     (existing_df['Region'].isin(input_df['Region'])))
        else:
            # กรณีทั่วไป: เช็ค Month + Year
            mask = ~((existing_df['Month'] == str(month)) & 
                     (existing_df['Year'] == str(year)))
        
        existing_df = existing_df[mask]

    # รวมข้อมูลเก่าที่เหลือกับข้อมูลใหม่
    updated_df = pd.concat([existing_df, input_df], ignore_index=True)
    
    # จัดเรียงลำดับเดือน (Jan -> Dec)
    month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    updated_df['Month'] = pd.Categorical(updated_df['Month'], categories=month_order, ordered=True)
    
    if is_export:
        updated_df['Region'] = pd.Categorical(updated_df['Region'], categories=REGIONS_LIST, ordered=True)
        updated_df = updated_df.sort_values(by=['Year', 'Region', 'Month'])
    else:
        updated_df = updated_df.sort_values(by=['Year', 'Month'])

    # บันทึกกลับ: ล้างข้อมูลเก่าทั้งหมดใน Sheet ก่อนเพื่อความสะอาด
    worksheet.clear()
    set_with_dataframe(worksheet, updated_df)
    return True

# --- UI INTERFACE ---
st.set_page_config(page_title="Data Management System", layout="wide")
st.title("🚀 ระบบจัดการข้อมูลยานยนต์ (Google Sheets)")

with st.sidebar:
    st.header("📅 ช่วงเวลา")
    sel_month = st.selectbox("เลือกเดือน", ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    sel_year = st.text_input("ระบุปี (พ.ศ.)", value="2568")
    st.divider()
    category = st.radio("เลือกหัวข้อที่ต้องการกรอก:", list(GID_MAP.keys()))

# --- หมวด 1.3 ยอดส่งออกรถยนต์ (Paste จาก Excel) ---
if category == "1.3 ยอดส่งออกรถยนต์":
    st.subheader(f"📍 {category}")
    st.info("💡 วิธีใช้: ก๊อปปี้คอลัมน์ตัวเลขจาก Excel แนวตั้งมาวาง (รวม 5 บรรทัดต่อภูมิภาค: 4 ประเภท + 1 Total)")
    
    raw_paste = st.text_area("วางคอลัมน์ตัวเลขจาก Excel ที่นี่", height=300)
    
    if raw_paste:
        lines = raw_paste.strip().split('\n')
        vals = [float(l.replace(',', '').strip()) if l.strip() not in ['-', '', ' '] else 0 for l in lines]

        rows = []
        idx = 0
        for reg in REGIONS_LIST:
            # ดึงข้อมูล 5 ลำดับ (Pickup, Passenger, PPV, Truck และบรรทัด Total)
            chunk = vals[idx : idx+5] if idx+5 <= len(vals) else [0,0,0,0,0]
            rows.append({
                "Month": sel_month, "Year": sel_year, "Region": reg,
                "Pickup": chunk[0], "Passenger": chunk[1], "PPV": chunk[2], "Truck": chunk[3],
                "Total_region": chunk[4] # ใช้ค่า Total ที่ก๊อปมาจริง ไม่ต้องคำนวณใหม่
            })
            idx += 5
            
        final_df = pd.DataFrame(rows)
        st.write("### 🔍 ตรวจสอบข้อมูลก่อนบันทึก")
        edited_final = st.data_editor(final_df, use_container_width=True)
        
        if st.button("🚀 บันทึกข้อมูล (เขียนทับค่าเก่า)"):
            if update_google_sheet(GID_MAP[category], edited_final, sel_month, sel_year, is_export=True):
                st.success(f"บันทึกยอดส่งออกเดือน {sel_month}/{sel_year} เรียบร้อย!")

# --- หมวดหมู่อื่นๆ (เพิ่มช่อง Total ให้กรอก) ---
else:
    st.subheader(f"📊 {category}")
    
    # กำหนดคอลัมน์ที่ต้องการแสดง (รวม Total)
    col_config = {
        "1.1 ยอดผลิตรถยนต์": ["Passenger", "Pickup", "Commercial", "Total"],
        "1.2 ยอดจำหน่ายรถยนต์": ["Passenger", "Pickup", "Commercial", "PPV_SUV", "Total"],
        "2.1 ยอดผลิตจักรยานยนต์": ["Family", "Sport", "EV", "ICONIC", "Total"],
        "2.2 ยอดจำหน่ายจักรยานยนต์": ["< 50 CC", "51-110 CC", "111-125 CC", "126-250 CC", "251-399 CC", "< 400 CC", "Total"],
        "2.3 ยอดส่งออกจักรยานยนต์": ["CBU", "CKD", "Value", "Total"]
    }
    
    current_cols = col_config[category]
    template_df = pd.DataFrame([{"Month": sel_month, "Year": sel_year, **{c: 0 for c in current_cols}}])
    
    st.write(f"กรุณากรอกตัวเลข (ข้อมูลจะเขียนทับหากเดือน {sel_month} ปี {sel_year} มีอยู่แล้ว)")
    edited_df = st.data_editor(template_df, num_rows="dynamic", hide_index=True, use_container_width=True)

    if st.button(f"💾 บันทึก {category}"):
        if update_google_sheet(GID_MAP[category], edited_df, sel_month, sel_year):
            st.success(f"บันทึกข้อมูล {category} เรียบร้อย!")





