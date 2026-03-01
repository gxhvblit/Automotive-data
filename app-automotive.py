import streamlit as st
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe, get_as_dataframe
from google.oauth2.service_account import Credentials

# --- CONFIG ---
SHEET_ID = "17Nq4MVLOKtdantiDayXwAgPRZKCvkI1FD4n7FJMZlJo"

GID_MAP = {
    "1.1 ยอดผลิตรถยนต์": "446114989",
    "1.2 ยอดจำหน่ายรถยนต์": "0",
    "1.3 ยอดส่งออกรถยนต์": "1329648063",
    "2.1 ยอดผลิตจักรยานยนต์": "2053730588",
    "2.2 ยอดจำหน่ายจักรยานยนต์": "787611063",
    "2.3 ยอดส่งออกจักรยานยนต์": "77561054"
}

# รายชื่อภูมิภาคตามรูปแนบที่ 1 สำหรับยอดส่งออก
REGIONS_LIST = [
    "Asia", "Australia, NZ & Other Oceania", "Middle East", 
    "Africa", "Europe", "North America", "Central & South America", "Others"
]

# --- FUNCTIONS ---
def get_gspread_client():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds_info = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)

def update_google_sheet(gid, input_df):
    gc = get_gspread_client()
    sh = gc.open_by_key(SHEET_ID)
    worksheet = sh.get_as_spreadsheet().get_worksheet_by_id(int(gid))
    
    # 1. ดึงข้อมูลเดิมจาก Google Sheet
    existing_df = get_as_dataframe(worksheet).dropna(how='all').dropna(axis=1, how='all')
    
    # 2. แปลงข้อมูล Year/Month ให้เป็น String เพื่อความแม่นยำในการเปรียบเทียบ
    input_df['Year'] = input_df['Year'].astype(str)
    input_df['Month'] = input_df['Month'].astype(str)
    
    if not existing_df.empty:
        existing_df['Year'] = existing_df['Year'].astype(str)
        existing_df['Month'] = existing_df['Month'].astype(str)
        
        # 3. ลบข้อมูลเก่าที่มี Year/Month (และ Region) ตรงกับที่กรอกใหม่ (Overwrite Logic)
        for _, row in input_df.iterrows():
            if 'Region' in input_df.columns:
                mask = ~((existing_df['Month'] == row['Month']) & 
                         (existing_df['Year'] == row['Year']) & 
                         (existing_df['Region'] == row['Region']))
            else:
                mask = ~((existing_df['Month'] == row['Month']) & 
                         (existing_df['Year'] == row['Year']))
            existing_df = existing_df[mask]

    # 4. รวมข้อมูลเก่าและใหม่
    updated_df = pd.concat([existing_df, input_df], ignore_index=True)
    
    # 5. จัดรูปแบบลำดับตาม Year -> Month -> Region (แบบรูปที่ 2)
    month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    updated_df['Month'] = pd.Categorical(updated_df['Month'], categories=month_order, ordered=True)
    
    if 'Region' in updated_df.columns:
        updated_df['Region'] = pd.Categorical(updated_df['Region'], categories=REGIONS_LIST, ordered=True)
        updated_df = updated_df.sort_values(by=['Year', 'Region', 'Month'])
    else:
        updated_df = updated_df.sort_values(by=['Year', 'Month'])

    # 6. บันทึกกลับลง Google Sheet
    set_with_dataframe(worksheet, updated_df)
    return True

# --- UI INTERFACE ---
st.set_page_config(page_title="Data Entry System", layout="wide")
st.title("📋 ระบบจัดการข้อมูลยานยนต์ (รองรับ Excel Copy/Paste)")

category = st.selectbox("เลือกหัวข้อที่ต้องการจัดการ", list(GID_MAP.keys()))

# เตรียมข้อมูลเบื้องต้น
month_default = st.sidebar.selectbox("Month (Default)", ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
year_default = st.sidebar.text_input("Year (Default)", value="2567")

st.divider()

if category == "1.3 ยอดส่งออกรถยนต์":
    st.subheader(f"📍 {category}")
    st.info("💡 คุณสามารถก๊อปปี้ข้อมูลจาก Excel (5 คอลัมน์: Pickup, Passenger, PPV, Truck) มาวางลงในตารางด้านล่างได้เลย")
    
    # สร้าง DataFrame เริ่มต้นตาม Region สำหรับยอดส่งออก
    init_data = {
        "Month": [month_default] * len(REGIONS_LIST),
        "Year": [year_default] * len(REGIONS_LIST),
        "Region": REGIONS_LIST,
        "Pickup": [0]*len(REGIONS_LIST),
        "Passenger": [0]*len(REGIONS_LIST),
        "PPV": [0]*len(REGIONS_LIST),
        "Truck": [0]*len(REGIONS_LIST)
    }
    df_to_edit = pd.DataFrame(init_data)
    
    # ตารางสำหรับแก้ไขและวางข้อมูล
    edited_df = st.data_editor(df_to_edit, num_rows="dynamic", hide_index=True, use_container_width=True)
    edited_df["Total_region"] = edited_df[["Pickup", "Passenger", "PPV", "Truck"]].sum(axis=1)

else:
    # หมวดหมู่อื่นๆ ยังคงใช้รูปแบบเดิม แต่ปรับเป็นตารางเพื่อให้วางข้อมูลได้เช่นกัน
    st.subheader(f"📊 {category}")
    
    col_maps = {
        "1.1 ยอดผลิตรถยนต์": ["Month", "Year", "Passenger", "Pickup", "Commercial"],
        "1.2 ยอดจำหน่ายรถยนต์": ["Month", "Year", "Passenger", "Pickup", "Commercial", "PPV_SUV"],
        "2.1 ยอดผลิตจักรยานยนต์": ["Month", "Year", "Family", "Sport", "EV", "ICONIC"],
        "2.2 ยอดจำหน่ายจักรยานยนต์": ["Month", "Year", "< 50 CC", "51-110 CC", "111-125 CC", "126-250 CC", "251-399 CC", "< 400 CC"],
        "2.3 ยอดส่งออกจักรยานยนต์": ["Month", "Year", "CBU", "CKD", "Value"]
    }
    
    current_cols = col_maps[category]
    df_template = pd.DataFrame([{"Month": month_default, "Year": year_default, **{c: 0 for c in current_cols if c not in ["Month", "Year"]}}])
    
    edited_df = st.data_editor(df_template, num_rows="dynamic", hide_index=True, use_container_width=True)
    
    # คำนวณ Total อัตโนมัติ
    numeric_cols = [c for c in edited_df.columns if c not in ["Month", "Year"]]
    edited_df["Total"] = edited_df[numeric_cols].sum(axis=1)

# ปุ่มบันทึก
if st.button("🚀 บันทึกข้อมูลและจัดรูปแบบลง Google Sheet", use_container_width=True):
    with st.spinner("กำลังประมวลผลและจัดเรียงข้อมูล..."):
        try:
            update_google_sheet(GID_MAP[category], edited_df)
            st.success(f"บันทึกและจัดรูปแบบข้อมูล {category} เรียบร้อยแล้ว!")
            st.balloons()
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")

