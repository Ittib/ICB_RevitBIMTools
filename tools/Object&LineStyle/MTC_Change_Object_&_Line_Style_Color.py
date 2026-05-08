import clr
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
doc = DocumentManager.Instance.CurrentDBDocument

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

# 🔹 กำหนดสีดำ
black = Color(0, 0, 0)

# 🔹 เริ่ม Transaction
t = Transaction(doc, "Change Object and Line Styles to Black")
t.Start()

# 🔹 รายชื่อ Category ที่ต้องยกเว้น
excluded_names = [
    "Revision Clouds",
    "Revision Cloud Tags"
]

#Version Script ที่เก็บไว้ใน log
script_version = 'v1.13'


# 🔹 ฟังก์ชันตรวจสอบชื่อที่ควรยกเว้น
def is_excluded(name):
    return name in excluded_names or ".dwg" in name.lower()

# 🔹 ลูปเปลี่ยนสี Category และ SubCategory
categories = doc.Settings.Categories
for cat in categories:
    try:
        if not is_excluded(cat.Name):
            cat.LineColor = black
        if cat.SubCategories.Size > 0:
            for subcat in cat.SubCategories:
                if not is_excluded(subcat.Name) and not is_excluded(cat.Name):
                    subcat.LineColor = black
    except:
        pass

t.Commit()
OUT = "✅ เปลี่ยนสี Object Style และ Subcategories เป็นสีดำ ยกเว้น Revision Cloud และไฟล์ .dwg"

# 🔹 เขียน Log
import csv
from datetime import datetime
import os

user_name = os.getenv("USERNAME")
project_name = doc.Title
log_action = "Change Object & Line Style"
log_status = "Success"
log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
log_path = r"R:\05_MTC_Library Development\MTC R24\00_BIM_Tools_Library\02_Dynamo_Scripts\00_Log\00_Dynamo Script_log.csv"

try:
    with open(log_path, mode='a', newline='', encoding='utf-8') as logfile:
        writer = csv.writer(logfile)
        writer.writerow([log_time, user_name, project_name, log_action, log_status, script_version])
except Exception as e:
    OUT = "❌ Error saving log: " + str(e)
