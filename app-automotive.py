import streamlit as st
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe, get_as_dataframe
from google.oauth2.service_account import Credentials

# --- CONFIG ---
SHEET_ID = "17Nq4MVLOKtdantiDayXwAgPRZKCvkI1FD4n7FJMZlJo"

# Mapping GID ตามที่คุณระบุ
GID_MAP = {
    "1.1 ยอดผลิตรถยนต์": "446114989",
    "1.2 ยอดจำหน่ายรถยนต์": "0",
    "1.3 ยอดส่งออกรถยนต์": "1329648063",
    "2.1 ยอดผลิตจักรยานยนต์": "2053730588",
    "2.2 ยอดจำหน่ายจักรยานยนต์": "787611063",
    "2.3 ยอดส่งออกจักรยานยนต์": "77561054"
}

# --- FUNCTIONS ---
def get_gspread_client():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds_info = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)

def update_google_sheet(gid, new_data_dict):
    gc = get_gspread_client()
    sh = gc.open_by_key(SHEET_ID)
    worksheet = sh.get_as_spreadsheet().get_worksheet_by_id(int(gid))
    
    # ดึงข้อมูลเดิม
    existing_df = get_as_dataframe(worksheet).dropna(how='all').dropna(axis=1, how='all')
    new_df = pd.DataFrame([new_data_dict])
    
    # Logic: ถ้า Month และ Year ซ้ำ ให้ลบอันเก่าออก (เขียนทับ)
    if not existing_df.empty:
        # ตรวจสอบว่ามีคอลัมน์ Month และ Year หรือไม่
        if 'Month' in existing_df.columns and 'Year' in existing_df.columns:
            mask = ~((existing_df['Month'] == new_data_dict['Month']) & 
                     (existing_df['Year'] == new_data_dict['Year']))
            existing_df = existing_df[mask]
    
    updated_df = pd.concat([existing_df, new_df], ignore_index=True)
    
    # กรณี 1.3 ยอดส่งออกรถยนต์: จัดเรียงตาม Region (ถ้ามี)
    if gid == GID_MAP["1.3 ยอดส่งออกรถยนต์"] and 'Region' in updated_df.columns:
        updated_df = updated_df.sort_values(by=['Year', 'Month', 'Region'])

    set_with_dataframe(worksheet, updated_df)
    return True

# --- UI INTERFACE ---
st.set_page_config(page_title="Data Entry System", layout="centered")
st.title("📋 ระบบกรอกข้อมูลอุตสาหกรรมยานยนต์")

category = st.selectbox("เลือกหัวข้อที่ต้องการกรอกข้อมูล", list(GID_MAP.keys()))

with st.form("entry_form", clear_on_submit=True):
    st.subheader(f"กรอกข้อมูล: {category}")
    
    # ส่วนกลางที่ทุกตารางต้องมี
    col1, col2 = st.columns(2)
    with col1:
        month = st.selectbox("Month", ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    with col2:
        year = st.text_input("Year (พ.ศ.)", value="2567")

    data = {"Month": month, "Year": year}

    # แยกฟิลด์ตามหมวดหมู่
    if category == "1.1 ยอดผลิตรถยนต์":
        data.update({
            "Passenger": st.number_input("Passenger", min_value=0),
            "Pickup": st.number_input("Pickup", min_value=0),
            "Commercial": st.number_input("Commercial", min_value=0)
        })
        data["Total"] = data["Passenger"] + data["Pickup"] + data["Commercial"]

    elif category == "1.2 ยอดจำหน่ายรถยนต์":
        data.update({
            "Passenger": st.number_input("Passenger", min_value=0),
            "Pickup": st.number_input("Pickup", min_value=0),
            "Commercial": st.number_input("Commercial", min_value=0),
            "PPV_SUV": st.number_input("PPV_SUV", min_value=0)
        })
        data["Total"] = data["Passenger"] + data["Pickup"] + data["Commercial"] + data["PPV_SUV"]

    elif category == "1.3 ยอดส่งออกรถยนต์":
        data.update({
            "Region": st.selectbox("Region", ["Asia", "Australia, NZ & Other Oceania", "Middle East", "Africa", "Europe", "North America", "Central & South America", "Others"]),
            "Pickup": st.number_input("Pickup", min_value=0),
            "Passenger": st.number_input("Passenger", min_value=0),
            "PPV": st.number_input("PPV", min_value=0),
            "Truck": st.number_input("Truck", min_value=0)
        })
        data["Total_region"] = data["Pickup"] + data["Passenger"] + data["PPV"] + data["Truck"]

    elif category == "2.1 ยอดผลิตจักรยานยนต์":
        data.update({
            "Family": st.number_input("Family", min_value=0),
            "Sport": st.number_input("Sport", min_value=0),
            "EV": st.number_input("EV", min_value=0),
            "ICONIC": st.number_input("ICONIC", min_value=0)
        })
        data["Total"] = data["Family"] + data["Sport"] + data["EV"] + data["ICONIC"]

    elif category == "2.2 ยอดจำหน่ายจักรยานยนต์":
        data.update({
            "< 50 CC": st.number_input("< 50 CC", min_value=0),
            "51-110 CC": st.number_input("51-110 CC", min_value=0),
            "111-125 CC": st.number_input("111-125 CC", min_value=0),
            "126-250 CC": st.number_input("126-250 CC", min_value=0),
            "251-399 CC": st.number_input("251-399 CC", min_value=0),
            "< 400 CC": st.number_input("< 400 CC", min_value=0)
        })
        data["Total"] = sum([data[k] for k in data if k not in ["Month", "Year"]])

    elif category == "2.3 ยอดส่งออกจักรยานยนต์":
        data.update({
            "CBU": st.number_input("CBU (Units)", min_value=0),
            "CKD": st.number_input("CKD (Sets)", min_value=0),
            "Value": st.number_input("Value (Mll. Baht)", min_value=0.0)
        })
        data["Total"] = data["CBU"] + data["CKD"]

    submit = st.form_submit_button("บันทึกข้อมูลลง Google Sheet")

if submit:
    with st.spinner("กำลังบันทึก..."):
        success = update_google_sheet(GID_MAP[category], data)
        if success:
            st.success(f"บันทึกข้อมูล {category} ของ {month}/{year} เรียบร้อยแล้ว!")
            st.write("ข้อมูลที่บันทึก:")
            st.json(data)
