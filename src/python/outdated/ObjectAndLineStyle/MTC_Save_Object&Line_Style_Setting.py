import clr
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import SaveFileDialog, DialogResult

import csv
import os
from datetime import datetime

# 🔹 กำหนดเวอร์ชันของสคริปต์
script_version = "v1.10"
#v.1.10 Update Log Path

# 🔹 ดึงเอกสาร Revit และข้อมูลผู้ใช้
doc = DocumentManager.Instance.CurrentDBDocument
categories = doc.Settings.Categories
project_name = doc.Title.replace(" ", "_").replace(".", "_").replace("-", "_")
user_name = os.getenv("USERNAME")
export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# 🔹 เตรียมข้อมูล ObjectStyle และ LineStyle แยกกัน
object_styles = []
line_styles_dict = {}

# 🔹 เงื่อนไขกรอง Category ที่ไม่ต้อง Export Subcategory
def is_excluded_category(cat_name):
    keywords = ["Import", "Link", "Analytical", "Analysis"]
    if any(keyword in cat_name for keyword in keywords):
        return True
    if ".dwg" in cat_name.lower():  # ตรวจสอบว่ามี .dwg อยู่ในชื่อ
        return True
    return False


# 🔹 ดึง Object Styles
for cat in categories:
    if is_excluded_category(cat.Name):
        continue
    try:
        object_styles.append([
            "ObjectStyle",
            cat.Name,
            "",
            cat.LineColor.Red,
            cat.LineColor.Green,
            cat.LineColor.Blue,
            cat.GetLineWeight(GraphicsStyleType.Projection),
            cat.GetLinePatternId(GraphicsStyleType.Projection).IntegerValue
        ])
        line_styles_dict[cat.Name] = []
    except:
        pass

# 🔹 ดึง Line Styles (SubCategories)
for cat in categories:
    if is_excluded_category(cat.Name):
        continue
    if cat.SubCategories.Size > 0:
        for subcat in cat.SubCategories:
            try:
                line_styles_dict[cat.Name].append([
                    "LineStyle",
                    "  " + subcat.Name,
                    cat.Name,
                    subcat.LineColor.Red,
                    subcat.LineColor.Green,
                    subcat.LineColor.Blue,
                    subcat.GetLineWeight(GraphicsStyleType.Projection),
                    subcat.GetLinePatternId(GraphicsStyleType.Projection).IntegerValue
                ])
            except:
                pass

# 🔹 รวมข้อมูลทั้งหมดแบบจัดลำดับ Category > Subcategory
data = [["Type", "Category", "Parent Category", "Color_R", "Color_G", "Color_B", "LineWeight", "LinePatternId"]]
for obj in sorted(object_styles, key=lambda x: x[1].lower()):
    data.append(obj)
    subcats = line_styles_dict.get(obj[1], [])
    subcats_sorted = sorted(subcats, key=lambda x: x[1].lower())
    data.extend(subcats_sorted)

# 🔹 เปิดหน้าต่างให้ผู้ใช้เลือก path สำหรับเซฟ CSV
dialog = SaveFileDialog()
dialog.Title = "เลือกที่เซฟไฟล์ CSV"
dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
dialog.DefaultExt = "csv"
dialog.FileName = project_name + "_style_backup.csv"

result = dialog.ShowDialog()

# 🔹 เซฟไฟล์ CSV
if result == DialogResult.OK:
    csv_path = dialog.FileName
    try:
        with open(csv_path, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(data)
        OUT = "✅ Exported to: " + csv_path
        log_status = "Success"
    except Exception as e:
        OUT = "❌ Error saving CSV: " + str(e)
        log_status = "Failed"
else:
    OUT = "❌ ไม่ได้เลือก path สำหรับเซฟไฟล์"
    log_status = "Cancelled"

# 🔹 เขียน Log ลงไฟล์ CSV พร้อมเวอร์ชัน
log_action = "Save Object & Line Style"
log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
log_path = r"R:\05_MTC_Library Development\MTC R24\00_BIM_Tools_Library\02_Dynamo_Scripts\00_Log\00_Dynamo Script_log.csv"

try:
    with open(log_path, mode='a', newline='', encoding='utf-8') as logfile:
        writer = csv.writer(logfile)
        writer.writerow([log_time, user_name, project_name, log_action, log_status, script_version])
except Exception as e:
    OUT += "\n❌ Error saving log: " + str(e)
