from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import cloudinary
import cloudinary.uploader
import openpyxl
from openpyxl import load_workbook
import os, json, re
from datetime import datetime
from typing import Optional
import httpx

app = FastAPI(title="BigC Champion API")

# CORS — อนุญาต GitHub Pages หรือทุก origin ระหว่าง dev
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ============================================================
# CONFIG — ใส่ค่าจริงใน Railway environment variables
# ============================================================
CLOUDINARY_CLOUD_NAME = os.getenv("CLOUDINARY_CLOUD_NAME", "")
CLOUDINARY_API_KEY    = os.getenv("CLOUDINARY_API_KEY", "")
CLOUDINARY_API_SECRET = os.getenv("CLOUDINARY_API_SECRET", "")

# OneDrive direct-download links ของ 2 ไฟล์ Excel
# วิธีได้ลิ้ง: เปิดไฟล์ใน OneDrive → Share → Copy link → เปลี่ยน ?e=xxx ท้าย URL เป็น download=1
RESPOND_EXCEL_URL  = os.getenv("RESPOND_EXCEL_URL", "")   # BigC_Champion_Respond.xlsx
DATABASE_EXCEL_URL = os.getenv("DATABASE_EXCEL_URL", "")  # ไฟล์ employee database

LOCAL_RESPOND_PATH  = "/tmp/respond.xlsx"
LOCAL_DATABASE_PATH = "/tmp/database.xlsx"
# ============================================================

cloudinary.config(
    cloud_name=CLOUDINARY_CLOUD_NAME,
    api_key=CLOUDINARY_API_KEY,
    api_secret=CLOUDINARY_API_SECRET,
    secure=True,
)


# ---- helpers -----------------------------------------------

async def download_excel(url: str, dest: str):
    """ดาวน์โหลด Excel จาก OneDrive มาเก็บที่ /tmp"""
    async with httpx.AsyncClient(follow_redirects=True, timeout=30) as client:
        r = await client.get(url)
        r.raise_for_status()
        with open(dest, "wb") as f:
            f.write(r.content)


def excel_serial_to_date(serial) -> str:
    """แปลง Excel date serial เป็น YYYY-MM-DD"""
    if isinstance(serial, (int, float)):
        base = datetime(1899, 12, 30)
        delta = __import__("datetime").timedelta(days=int(serial))
        return (base + delta).strftime("%Y-%m-%d")
    return str(serial)


# ---- routes ------------------------------------------------

@app.get("/")
def root():
    return {"status": "BigC Champion API is running 🚀"}


@app.get("/employee/{emp_id}")
async def get_employee(emp_id: str):
    """ค้นหาพนักงานจาก Database sheet"""
    try:
        if DATABASE_EXCEL_URL:
            await download_excel(DATABASE_EXCEL_URL, LOCAL_DATABASE_PATH)
        
        if not os.path.exists(LOCAL_DATABASE_PATH):
            raise HTTPException(status_code=503, detail="Database file not available")

        wb = load_workbook(LOCAL_DATABASE_PATH, read_only=True, data_only=True)
        ws = wb.active

        headers = [str(c.value).strip() if c.value else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]

        # หาตำแหน่ง column (case-insensitive)
        col = {h.lower(): i for i, h in enumerate(headers)}
        idx_code  = next((col[k] for k in col if "store code" in k or k == "store_code"), None)
        idx_sname = next((col[k] for k in col if "store name" in k or k == "store_name" or "storename" in k), None)
        idx_eid   = next((col[k] for k in col if "employee id" in k or k == "employee_id"), None)
        idx_ename = next((col[k] for k in col if "employee name" in k or k == "employee_name"), None)

        if idx_eid is None:
            raise HTTPException(status_code=500, detail="ไม่พบคอลัมน์ Employee ID ใน database")

        emp_id_clean = emp_id.strip()

        for row in ws.iter_rows(min_row=2, values_only=True):
            cell_id = str(row[idx_eid]).strip().split(".")[0]  # ตัด .0 ออก
            if cell_id == emp_id_clean:
                return {
                    "employeeId":   emp_id_clean,
                    "employeeName": str(row[idx_ename]) if idx_ename is not None else "",
                    "storeCode":    str(row[idx_code]).split(".")[0] if idx_code is not None else "",
                    "storeName":    str(row[idx_sname]) if idx_sname is not None else "",
                }

        raise HTTPException(status_code=404, detail="ไม่พบรหัสพนักงานนี้")

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/upload")
async def upload_photo(
    file: UploadFile = File(...),
    slot: str = Form(...),
    storeCode: str = Form(...),
):
    """อัปโหลดรูปไป Cloudinary แล้วคืน URL"""
    try:
        contents = await file.read()
        folder   = f"bigc_champion/{storeCode}"
        public_id = f"{storeCode}_{slot}_{int(datetime.now().timestamp())}"

        result = cloudinary.uploader.upload(
            contents,
            folder=folder,
            public_id=public_id,
            resource_type="image",
            overwrite=True,
        )
        return {"url": result["secure_url"], "slot": slot}

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Upload failed: {str(e)}")


class SubmitPayload(BaseModel):
    storeCode:    str
    storeName:    str
    employeeId:   str
    employeeName: str
    date:         str
    photos:       dict          # { "HCL1.1": "https://...", ... }
    freqQuestion: Optional[str] = ""
    hasBgc:       Optional[bool] = False


@app.post("/submit")
async def submit_respond(payload: SubmitPayload):
    """เพิ่มแถวใหม่ใน Respond.xlsx แล้วอัปโหลดกลับ OneDrive"""
    try:
        # ดาวน์โหลด Excel ล่าสุดมาก่อน
        if RESPOND_EXCEL_URL:
            await download_excel(RESPOND_EXCEL_URL, LOCAL_RESPOND_PATH)

        if not os.path.exists(LOCAL_RESPOND_PATH):
            # สร้างใหม่ถ้าไม่มี
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Respond"
            headers = [
                "Store Code","Store Name","Employee ID","Employee Name","Date",
                "HCL1.1","HCL1.2","HCL1.3",
                "HCL2.1","HCL2.2","HCL2.3",
                "HCL3.1","HCL3.2","HCL3.3",
                "BGC1.1","BGC1.2","BGC1.3",
                "BGC2.1","BGC2.2","BGC2.3",
                "BGC3.1","BGC3.2","BGC3.3",
                "BGC4.1","BGC4.2","BGC4.3",
                "BGC5.1","BGC5.2","BGC5.3",
                "Freq.Qs.","HAS_BGC",
            ]
            ws.append(headers)
        else:
            wb = load_workbook(LOCAL_RESPOND_PATH)
            ws = wb["Respond"] if "Respond" in wb.sheetnames else wb.active

        # map headers → column index
        header_row = [str(c.value).strip() if c.value else "" for c in ws[1]]
        h_idx = {h: i+1 for i, h in enumerate(header_row)}

        # สร้างแถวใหม่
        new_row = [""] * len(header_row)

        def set_col(name, val):
            if name in h_idx:
                new_row[h_idx[name] - 1] = val

        set_col("Store Code",    payload.storeCode)
        set_col("Store Name",    payload.storeName)
        set_col("Employee ID",   payload.employeeId)
        set_col("Employee Name", payload.employeeName)
        set_col("Date",          payload.date)
        set_col("HAS_BGC",       "YES" if payload.hasBgc else "NO")
        set_col("Freq.Qs.",      payload.freqQuestion or "")

        for slot_key, url in payload.photos.items():
            # normalize: HCL1.1 → HCL1.1
            col_name = re.sub(r'(\d)(\d)', r'\1.\2', slot_key) if '.' not in slot_key else slot_key
            set_col(col_name, url)

        ws.append(new_row)
        wb.save(LOCAL_RESPOND_PATH)

        # TODO: อัปโหลด LOCAL_RESPOND_PATH กลับขึ้น OneDrive
        # ตอนนี้บันทึกที่ /tmp ก่อน — เพิ่ม Microsoft Graph upload ตรงนี้ได้ทีหลัง

        return {
            "status": "success",
            "message": f"บันทึกข้อมูล {payload.employeeName} — {payload.storeName} เรียบร้อย",
            "row_added": True,
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
