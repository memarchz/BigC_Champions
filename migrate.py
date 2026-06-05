import os # จัดการ File/Folder
import csv # อ่านไฟล์ตาราง
import re # สเเกนตัดคำ/ข้อความ
import time
import requests # โหลดข้อมูลจากเน็ต
import threading # เเบ่งงานเพื่อทำพร้อมกันหลาย Thread
from concurrent.futures import ThreadPoolExecutor

# ==================== [ ตั้งค่าระบบ ] ====================
INPUT_CSV = 'BigC_Champion_Respond.csv'         
PROGRESS_FILE = 'progress.csv' # ไฟล์ log ที่บอกว่า download อะไรไปเเล้วบ้าง          
ONEDRIVE_FOLDER = r'C:\Users\Win10\Berli Jucker Public Company Limited\COS Theerapat Pongpanich - Pic' # Folder OneDrive ในเครื่องที่จะเอาไปลง

MAX_WORKERS = 5 # จำนวน Thread ที่จะทำงานร่วมกัน              
BATCH_SIZE = 100 # ทำรอบละ 100 rows                   
# ============================================================

csv_lock = threading.Lock() # ป้องกันไม่ให้ thread เขียนข้อมูลลงพร้อมกัน 1 thread / 1 access
os.makedirs(ONEDRIVE_FOLDER, exist_ok=True) # สร้าง Folder ใน OneDrive

def extract_google_drive_id(url):
    if not url or not isinstance(url, str): return None # ตรวจว่า url ต้องไม่ว่างเปล่า เเละต้องเป็น string
    match = re.search(r'/d/([^/]+)', url)
    if match: return match.group(1)
    match = re.search(r'id=([^&]+)', url)
    if match: return match.group(1)
    # เพิ่มเติมเผื่อเคสลิงก์ติดกันแบบ https://drive.google.com01 
    match = re.search(r'drive\.google\.com[^\w]*([a-zA-Z0-9_-]{25,})', url) 
    if match: return match.group(1)
    return None

def download_file(file_id, output_path, retries=3, backoff_in_seconds=2): # รับ file_id และ path ที่จะบันทึก พร้อม retry สูงสุด 3 ครั้งถ้าล้มเหลว
    download_url = f'https://drive.google.com/uc?export=download&id={file_id}' # สร้าง URL ดาวน์โหลดตรง (direct download URL) จาก File ID — เป็น URL พิเศษของ Google Drive ที่บังคับให้ดาวน์โหลดไฟล์โดยตรง
    for i in range(retries):
        try:
            response = requests.get(download_url, stream=True, timeout=15)
            if response.status_code == 200:
                with open(output_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                return True
            elif response.status_code == 429:
                time.sleep(backoff_in_seconds * (i + 1))
            else:
                return False
        except Exception:
            time.sleep(backoff_in_seconds * (i + 1))
    return False

def process_single_image(store_code, store_name, date_str, url, col_name, processed_urls):
    if url in processed_urls:
        return
        
    file_id = extract_google_drive_id(url)
    if not file_id:
        return

    clean_store_name = re.sub(r'[\\/*?:"<>|]', "", str(store_name))
    store_folder_name = f"{store_code}_{clean_store_name}"
    store_folder_path = os.path.join(ONEDRIVE_FOLDER, store_folder_name)
    os.makedirs(store_folder_path, exist_ok=True)

    clean_col_name = re.sub(r'[\\/*?:"<>|]', "", str(col_name))
    
    # จุดที่แก้ไข: เพิ่มวันที่ (date_str) เข้าไปในชื่อไฟล์รูปภาพ
    filename = f"{store_code}_{date_str}_{clean_col_name}.jpg"
    output_path = os.path.join(store_folder_path, filename)

    # ตรวจสอบว่าในโฟลเดอร์มีไฟล์นี้โหลดเสร็จสมบูรณ์อยู่แล้วหรือยัง
    if os.path.exists(output_path):
        with csv_lock:
            with open(PROGRESS_FILE, 'a', encoding='utf-8', newline='') as pf:
                writer = csv.writer(pf)
                writer.writerow([url, store_code, 'ALREADY_EXISTS', time.strftime('%Y-%m-%d %H:%M:%S')])
        processed_urls.add(url)
        print(f"⏩ ข้าม (มีไฟล์แล้ว): {store_code} -> {filename}")
        return
    # ==================================

    is_success = download_file(file_id, output_path)
    
    if is_success:
        with csv_lock:
            with open(PROGRESS_FILE, 'a', encoding='utf-8', newline='') as pf:
                writer = csv.writer(pf)
                writer.writerow([url, store_code, 'SUCCESS', time.strftime('%Y-%m-%d %H:%M:%S')])
        processed_urls.add(url)
        print(f"สำเร็จ: {store_code} [{date_str}] -> {filename}")
    else:
        print(f"ล้มเหลว: {store_code} โหลดรูปจากคอลัมน์ {col_name} ไม่ได้")

def main():
    print("เริ่มการย้ายข้อมูล")
    
    processed_urls = set()
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            next(reader, None)
            for row in reader:
                if row: processed_urls.add(row[0])
                
    if not os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['url', 'store_code', 'status', 'timestamp'])

    with open(INPUT_CSV, 'r', encoding='utf-8-sig') as f:
        reader = list(csv.DictReader(f))
        total_rows = len(reader)
        print(f"📊 พบข้อมูลสำหรับทดสอบ: {total_rows} แถว")

        for i in range(0, total_rows, BATCH_SIZE):
            batch_rows = reader[i:i+BATCH_SIZE]
            print(f"\n📦 กำลังเริ่มทำ Batch ที่ {int(i/BATCH_SIZE)+1}")
            
            download_tasks = []
            for row in batch_rows:
                store_code = row.get('Store Code')
                store_name = row.get('Store Name')
                
                # ดึงคอลัมน์ Date และล้างเครื่องหมายสแลช (ถ้ามี) ให้กลายเป็นขีดกลาง
                date_val = row.get('Date', 'NoDate')
                clean_date = re.sub(r'[\\/*?:"<>|]', "-", str(date_val).strip())
                
                if not store_code: continue
                
                for col_name, value in row.items():
                    if value and 'drive.google.com' in str(value):
                        # ส่ง clean_date พ่วงไปด้วยในรายการงาน
                        download_tasks.append((store_code, store_name, clean_date, str(value).strip(), col_name))

            if download_tasks:
                with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                    for store_code, store_name, date_str, url, col_name in download_tasks:
                        executor.submit(process_single_image, store_code, store_name, date_str, url, col_name, processed_urls)
            
            time.sleep(3)

    print("\n Done")
if __name__ == '__main__':
    main()