import clr
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog, DialogResult

import csv
import os
from datetime import datetime

# 🔹 เวอร์ชันของสคริปต์
script_version = "v1.10"
#v.1.09 Update Log Path

# 🔹 ฟังก์ชันตรวจสอบ LineWeight
def safe_line_weight(value):
    try:
        lw = int(value)
        return max(1, min(16, lw))  # ปรับให้อยู่ในช่วงที่ Revit ยอมรับ
    except:
        return 1

# 🔹 ฟังก์ชันกรองชื่อ Parent Category ที่ไม่ต้องโหลด
def is_excluded_parent(parent_name):
    return ".dwg" in parent_name.lower()

# 🔹 เริ่มต้น
doc = DocumentManager.Instance.CurrentDBDocument
TransactionManager.Instance.EnsureInTransaction(doc)
project_name = doc.Title.replace(" ", "_").replace(".", "_").replace("-", "_")
user_name = os.getenv("USERNAME")

# 🔹 เปิดหน้าต่างเลือกไฟล์ CSV
dialog = OpenFileDialog()
dialog.Title = "เลือกไฟล์ CSV ที่จะโหลด"
dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
dialog.Multiselect = False
result = dialog.ShowDialog()

success_count = 0
skip_count = 0

if result == DialogResult.OK:
    csv_path = dialog.FileName

    # 🔹 ตรวจสอบ LinePatternId ที่มีอยู่จริง
    valid_line_pattern_ids = {lpe.Id.IntegerValue for lpe in FilteredElementCollector(doc).OfClass(LinePatternElement)}

    # 🔹 สร้าง dictionary ของ Categories ที่มีอยู่จริง
    available_categories = {cat.Name: cat for cat in doc.Settings.Categories}

    try:
        with open(csv_path, mode='r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                cat_name = row["Category"].strip()
                parent_name = row.get("Parent Category", "").strip()

                # 🔹 ข้ามถ้าเป็น Subcategory ของ .dwg
                if row["Type"] == "LineStyle" and is_excluded_parent(parent_name):
                    skip_count += 1
                    continue

                try:
                    color = Color(int(row["Color_R"]), int(row["Color_G"]), int(row["Color_B"]))
                    line_weight = safe_line_weight(row["LineWeight"])
                    line_pattern_id = int(row["LinePatternId"])
                except:
                    skip_count += 1
                    continue

                # 🔹 ObjectStyle
                if row["Type"] == "ObjectStyle":
                    if cat_name in available_categories:
                        cat = available_categories[cat_name]
                        cat.LineColor = color
                        cat.SetLineWeight(line_weight, GraphicsStyleType.Projection)
                        if line_pattern_id in valid_line_pattern_ids:
                            cat.SetLinePatternId(ElementId(line_pattern_id), GraphicsStyleType.Projection)
                        success_count += 1
                    else:
                        skip_count += 1

                # 🔹 LineStyle (SubCategory) โดยใช้ Parent Category
                elif row["Type"] == "LineStyle":
                    found = False
                    for cat in doc.Settings.Categories:
                        if cat.Name == parent_name and cat.SubCategories.Size > 0:
                            for subcat in cat.SubCategories:
                                if subcat.Name == cat_name:
                                    subcat.LineColor = color
                                    subcat.SetLineWeight(line_weight, GraphicsStyleType.Projection)
                                    if line_pattern_id in valid_line_pattern_ids:
                                        subcat.SetLinePatternId(ElementId(line_pattern_id), GraphicsStyleType.Projection)
                                    success_count += 1
                                    found = True
                                    break
                        if found:
                            break
                    if not found:
                        skip_count += 1

        log_status = "Success"
        OUT = f"✅ Loaded from {csv_path}\nเปลี่ยนสำเร็จ: {success_count} รายการ\nข้าม: {skip_count} รายการ"

    except Exception as e:
        log_status = "Failed"
        OUT = "❌ Error loading CSV: " + str(e)
else:
    log_status = "Cancelled"
    OUT = "❌ ไม่ได้เลือกไฟล์ CSV"

TransactionManager.Instance.TransactionTaskDone()

# 🔹 บันทึก Log พร้อมเวอร์ชัน
log_action = "Load Object & Line Style"
log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
log_path = r"R:\05_MTC_Library Development\MTC R24\00_BIM_Tools_Library\02_Dynamo_Scripts\00_Log\00_Dynamo Script_log.csv"

try:
    with open(log_path, mode='a', newline='', encoding='utf-8') as logfile:
        writer = csv.writer(logfile)
        writer.writerow([log_time, user_name, project_name, log_action, log_status, script_version])
except Exception as e:
    OUT += "\n❌ Error saving log: " + str(e)
