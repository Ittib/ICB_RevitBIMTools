import clr
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
doc = DocumentManager.Instance.CurrentDBDocument

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

# 🔹 กำหนดค่าสีดำ (RGB: 0, 0, 0)
black = Color(0, 0, 0)

# 🔹 เริ่มบันทึกการเปลี่ยนแปลงใน Revit (Transaction)
t = Transaction(doc, "Change Object and Line Styles to Black")
t.Start()

# 🔹 รายชื่อหมวดหมู่ (Category) ที่เราไม่ต้องการเปลี่ยนสี
excluded_names = [
    "Revision Clouds",
    "Revision Cloud Tags"
]

# เวอร์ชั่นของสคริปต์สำหรับเก็บประวัติ
script_version = 'v1.13'

# 🔹 ฟังก์ชันเช็คว่าชื่อนี้อยู่ในรายการยกเว้น หรือเป็นไฟล์ Link (.dwg) หรือไม่
def is_excluded(name):
    return name in excluded_names or ".dwg" in name.lower()

# 🔹 เริ่มวนลูปตรวจสอบทุก Categories ในโปรเจกต์
categories = doc.Settings.Categories
for cat in categories:
    try:
        # ถ้าไม่อยู่ในรายการยกเว้น ให้เปลี่ยนเส้นหลักเป็นสีดำ
        if not is_excluded(cat.Name):
            cat.LineColor = black
        
        # ถ้ามีหมวดหมู่ย่อย (SubCategories) ให้ไล่เปลี่ยนสีทีละอันด้วย
        if cat.SubCategories.Size > 0:
            for subcat in cat.SubCategories:
                # ตรวจสอบซ้ำว่าหมวดหมู่ย่อยหรือหมวดหมู่หลักอยู่ในรายการยกเว้นหรือไม่
                if not is_excluded(subcat.Name) and not is_excluded(cat.Name):
                    subcat.LineColor = black
    except:
        # หากบาง Category เปลี่ยนสีไม่ได้ (เช่น System-defined) ให้ข้ามไป
        pass

# 🔹 ยืนยันการเปลี่ยนแปลงลงในไฟล์ Revit
t.Commit()
OUT = "✅ เปลี่ยนสี Object Style และ Subcategories เป็นสีดำ ยกเว้น Revision Cloud และไฟล์ .dwg"

# 🔹 ส่วนการบันทึกประวัติการใช้งาน (Logging)
import csv
from datetime import datetime
import os

# ดึงข้อมูลผู้ใช้, ชื่อโปรเจกต์ และเวลาปัจจุบัน
user_name = os.getenv("USERNAME")
project_name = doc.Title
log_action = "Change Object & Line Style"
log_status = "Success"
log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ที่อยู่ไฟล์ CSV สำหรับเก็บ Log กลางของทีม
log_path = r"R:\05_MTC_Library Development\MTC R24\00_BIM_Tools_Library\02_Dynamo_Scripts\00_Log\00_Dynamo Script_log.csv"

try:
    # เปิดไฟล์ Log และเขียนข้อมูลลงไปต่อท้าย (Append mode)
    with open(log_path, mode='a', newline='', encoding='utf-8') as logfile:
        writer = csv.writer(logfile)
        writer.writerow([log_time, user_name, project_name, log_action, log_status, script_version])
except Exception as e:
    # หากบันทึก Log ไม่สำเร็จ ให้แสดง Error แทน
    OUT = "❌ Error saving log: " + str(e)