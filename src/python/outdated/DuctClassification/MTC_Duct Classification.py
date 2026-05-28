import clr
import os

# --- Revit API Reference ---
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# --- ข้อมูล Metadata ---
script_version = "v1.2.0"

doc            = DocumentManager.Instance.CurrentDBDocument

# รับข้อมูลจาก Dynamo Input และทำการ Unwrap เพื่อใช้งาน Revit API
raw_input = IN[0] if isinstance(IN[0], list) else [IN[0]]
items = [UnwrapElement(i) for i in raw_input if i is not None]

success_count = 0

# --- ส่วนหลัก: Duct Classification ---
TransactionManager.Instance.EnsureInTransaction(doc)

for item in items:
    try:
        # ตรวจสอบว่าเป็น Revit Element หรือไม่[cite: 1]
        if not hasattr(item, "Id"): continue
            
        w_param = item.LookupParameter("Width")
        h_param = item.LookupParameter("Height")
        d_param = item.LookupParameter("Diameter")
        
        # ดึงค่าและแปลงหน่วย (Feet -> mm)[cite: 1]
        w = w_param.AsDouble() * 304.8 if w_param and w_param.HasValue else 0
        h = h_param.AsDouble() * 304.8 if h_param and h_param.HasValue else 0
        d = d_param.AsDouble() * 304.8 if d_param and d_param.HasValue else 0
        
        max_dim = max(w, h, d)
        
        # Logic การแบ่งประเภทตามขนาด[cite: 1]
        if max_dim > 2299:
            val = "Extra Heavy"
        elif max_dim > 1524:
            val = "Heavy"
        elif max_dim > 774:
            val = "Medium"
        elif max_dim > 599:
            val = "Light-Medium"
        else:
            val = "Light"
            
        # เขียนค่าลงใน Parameter[cite: 1]
        out_param = item.LookupParameter("Duct Classification")
        if out_param and not out_param.IsReadOnly:
            out_param.Set(val)
            success_count += 1
            
    except:
        continue

TransactionManager.Instance.TransactionTaskDone()

# ส่งผลลัพธ์กลับไปยัง Dynamo[cite: 1]
OUT = "Success: {} elements processed | Version: {}".format(success_count, script_version)

# 🔹 เขียน Log
import csv
from datetime import datetime
import os

user_name = os.getenv("USERNAME")
project_name = doc.Title
log_action = "1-Click Duct Classification"
log_status = "Success"
log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
log_path = r"R:\05_MTC_Library Development\MTC R24\00_BIM_Tools_Library\02_Dynamo_Scripts\00_Log\00_Dynamo Script_log.csv"

try:
    with open(log_path, mode='a', newline='', encoding='utf-8') as logfile:
        writer = csv.writer(logfile)
        writer.writerow([log_time, user_name, project_name, log_action, log_status, script_version])
except Exception as e:
    OUT = "❌ Error saving log: " + str(e)
