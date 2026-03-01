/**
 * ฟังก์ชันหลักสำหรับบันทึกข้อมูลรถยนต์และรถจักรยานยนต์
 * ตรวจสอบ Month/Year ซ้ำเพื่อเขียนทับ (Overwrite)
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

// --- 1. ฟังก์ชันบันทึกข้อมูลทั่วไป (1.1, 1.2, 2.1, 2.2, 2.3) ---
function saveGeneralData(sheetName, dataObj) {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) {
    Browser.msgBox("ไม่พบ Sheet ชื่อ: " + sheetName);
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const allData = sheet.getDataRange().getValues();
  let targetRow = -1;

  // ค้นหา Month และ Year ซ้ำ (Col A และ B)
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0].toString() == dataObj.Month.toString() && 
        allData[i][1].toString() == dataObj.Year.toString()) {
      targetRow = i + 1;
      break;
    }
  }

  // เตรียมข้อมูลตาม Header
  const rowData = headers.map(h => (dataObj[h] !== undefined ? dataObj[h] : 0));

  if (targetRow !== -1) {
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

// --- 2. ฟังก์ชันพิเศษสำหรับยอดส่งออกรถยนต์ (1.3) ---
// ฟังก์ชันนี้จะดึงค่าจากหน้าฟอร์มที่จัดรูปแบบตามรูปที่ 1 แล้วส่งไปบันทึกแบบรูปที่ 2
function processExportCarData() {
  const inputSheet = SS.getSheetByName("Input_ExportCar"); // สมมติชื่อ Sheet ที่ใช้กรอก
  const dbSheet = SS.getSheetByName("Export_Car_DB");     // Sheet 1.3
  
  const month = inputSheet.getRange("B1").getValue(); // ตำแหน่ง Month ในฟอร์ม
  const year = inputSheet.getRange("B2").getValue();  // ตำแหน่ง Year ในฟอร์ม
  
  // นิยามช่วงข้อมูลของแต่ละภูมิภาค (ตามรูปที่ 1)
  // ตัวอย่าง: Asia อยู่แถว 3-7, Category อยู่ Col B, Value อยู่ Col C
  const regions = [
    { name: "Asia", range: "B3:C7" },
    { name: "Australia, NZ & Other Oceania", range: "B8:C12" },
    { name: "Middle East", range: "B13:C17" },
    { name: "Africa", range: "B18:C22" },
    { name: "Europe", range: "B23:C27" },
    { name: "North America", range: "B28:C32" },
    { name: "Central & South America", range: "B33:C37" },
    { name: "Others", range: "B38:C42" }
  ];

  regions.forEach(reg => {
    const rawValues = inputSheet.getRange(reg.range).getValues();
    let dataForDB = {
      "Month": month,
      "Year": year,
      "Region": reg.name,
      "Pickup": 0, "Passenger": 0, "PPV": 0, "Truck": 0, "Total_region": 0
    };

    rawValues.forEach(row => {
      const cat = row[0].toString();
      const val = parseFloat(row[1]) || 0;
      if (cat.includes("1 Ton P/U")) dataForDB["Pickup"] = val;
      else if (cat.includes("Passenger")) dataForDB["Passenger"] = val;
      else if (cat.includes("PPV")) dataForDB["PPV"] = val;
      else if (cat.includes("Truck")) dataForDB["Truck"] = val;
    });

    dataForDB["Total_region"] = dataForDB["Pickup"] + dataForDB["Passenger"] + dataForDB["PPV"] + dataForDB["Truck"];
    
    // เรียกใช้ฟังก์ชันบันทึกแบบเช็คซ้ำ (Overwrite) โดยเช็ค 3 เงื่อนไข: Month, Year, Region
    saveExportWithRegionCheck("Export_Car_DB", dataForDB);
  });
  
  Browser.msgBox("บันทึกข้อมูลยอดส่งออกเรียบร้อยแล้ว!");
}

function saveExportWithRegionCheck(sheetName, data) {
  const sheet = SS.getSheetByName(sheetName);
  const allData = sheet.getDataRange().getValues();
  let targetRow = -1;

  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] == data.Month && allData[i][1] == data.Year && allData[i][2] == data.Region) {
      targetRow = i + 1;
      break;
    }
  }

  const rowData = [data.Month, data.Year, data.Region, data.Pickup, data.Passenger, data.PPV, data.Truck, data.Total_region];
  if (targetRow !== -1) {
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}