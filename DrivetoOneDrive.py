import csv
import os
import re
from urllib.parse import quote  # 🔥 นำเข้าเครื่องมือแปลงภาษาไทย/ช่องว่าง ให้เป็นลิงก์เว็บ

# ==================== [ ตั้งค่าสำหรับ BJC OneDrive ] ====================
INPUT_CSV = 'BigC_Champion_Respond.csv'         
OUTPUT_CSV = 'BigC_Champion_Respond_Final.csv'   

# โครงสร้างใหม่ตามภาพ: Documents -> BigC Champion Data -> Pic
# (ใช้ %20 แทนช่องว่าง เพื่อให้เป็น URL มาตรฐานตั้งแต่แรก)
BJC_ONEDRIVE_BASE = "https://bjcgrp-my.sharepoint.com/personal/theerapat_pon_bjc_co_th/Documents/BigC%20Champion%20Data/Pic"
# =======================================================================

def main():
    print("🔄 กำลังแปลงลิงก์ (ปรับตามโครงสร้างโฟลเดอร์ใหม่)...")

    if not os.path.exists(INPUT_CSV):
        print(f"❌ ไม่พบไฟล์ {INPUT_CSV}")
        return

    with open(INPUT_CSV, 'r', encoding='utf-8-sig') as f_in, \
         open(OUTPUT_CSV, 'w', encoding='utf-8', newline='') as f_out:
         
        reader_obj = csv.DictReader(f_in)
        fieldnames = reader_obj.fieldnames 
        
        # ทดสอบเฉพาะ 10 แถวแรก
        rows = list(reader_obj)[:10]
        
        writer = csv.DictWriter(f_out, fieldnames=fieldnames)
        writer.writeheader() 
        
        row_count = 0
        link_count = 0

        for row in rows:
            store_code = row.get('Store Code')
            store_name = row.get('Store Name')
            
            date_val = row.get('Date', 'NoDate')
            clean_date = re.sub(r'[\\/*?:"<>|]', "-", str(date_val).strip())
            
            if store_code:
                clean_store_name = re.sub(r'[\\/*?:"<>|]', "", str(store_name))
                store_folder_name = f"{store_code}_{clean_store_name}"
                
                for col_name, value in row.items():
                    if value and 'drive.google.com' in str(value):
                        clean_col_name = re.sub(r'[\\/*?:"<>|]', "", str(col_name))
                        filename = f"{store_code}_{clean_date}_{clean_col_name}.jpg"
                        
                        # เข้ารหัสชื่อโฟลเดอร์และชื่อไฟล์ (แปลงภาษาไทย/ช่องว่าง ให้เบราว์เซอร์อ่านออก)
                        encoded_folder = quote(store_folder_name)
                        encoded_file = quote(filename)
                        
                        # ประกอบร่างเป็นลิงก์ที่สมบูรณ์
                        new_onedrive_url = f"{BJC_ONEDRIVE_BASE}/{encoded_folder}/{encoded_file}"
                        
                        row[col_name] = new_onedrive_url
                        link_count += 1
            
            writer.writerow(row)
            row_count += 1

    print(f"\n[ทดสอบเสร็จสิ้น] สร้างไฟล์ '{OUTPUT_CSV}' เรียบร้อย!")
    print(f" อัปเดตลิงก์ทดสอบไปทั้งหมด: {link_count} ลิงก์")

if __name__ == '__main__':
    main()