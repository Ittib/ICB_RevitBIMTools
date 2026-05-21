import clr
import sys

# 1. Import Revit API & Services
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# 2. Import Windows Forms สำหรับ UI เลือกไฟล์
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog, DialogResult

# 3. Import Excel Interop
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

# เตรียมข้อมูลเบื้องต้น
doc = DocumentManager.Instance.CurrentDBDocument
project_info = doc.ProjectInformation

# --- ส่วนที่ 1: UI เลือกไฟล์ Excel ---
file_path = ""
dialog = OpenFileDialog()
dialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
dialog.Title = "กรุณาเลือกไฟล์ Excel สำหรับข้อมูลผู้ออกแบบ"

if dialog.ShowDialog() == DialogResult.OK:
    file_path = dialog.FileName

if not file_path:
    OUT = "ยกเลิกการเลือกไฟล์"
    sys.exit()

# --- ส่วนที่ 2: อ่านข้อมูลจาก Excel และ Update Revit ---
excel_app = Excel.ApplicationClass()
# ตั้งค่าให้ทำงานเบื้องหลัง (ไม่ต้องเปิดหน้าต่าง Excel)
excel_app.Visible = False 
target_sheet = "Eng name"


try:
    workbook = excel_app.Workbooks.Open(file_path)
    worksheet = workbook.Worksheets[target_sheet] # เลือก Sheet แรก
    
    # เริ่ม Transaction ใน Revit
    TransactionManager.Instance.EnsureInTransaction(doc)
    
    results = []
    # วนลูปอ่านตั้งแต่แถวที่ 4 เป็นต้นไป (ตามรูปภาพ)
    # สมมติว่าอ่านถึงแถวที่ 100 (หรือปรับตามต้องการ)
    for row in range(4, 100):
        # Col A (1): Parameter Name (เช่น Architect 1)
        param_name = worksheet.Cells[row, 1].Value2
        # Col B (2): Thai Name
        thai_name = worksheet.Cells[row, 4].Value2
        # Col D (4): Id (License No.)
        license_id = worksheet.Cells[row, 3].Value2
        
        if not param_name: continue # ถ้าชื่อตำแหน่งว่างให้ข้ามแถวนั้น
        
        # กรองข้อมูลและนำไปกรอกใน Revit
        # 1. กรอกชื่อภาษาไทย (Thai Name)
        p1 = project_info.LookupParameter(str(param_name))
        if p1:
            p1.Set(str(thai_name) if thai_name else "")
            
        # 2. กรอกเลขใบอนุญาต (License No.)
        # สร้างชื่อ Parameter ตามเงื่อนไข: ชื่อตำแหน่ง + " License No."
        license_param_name = str(param_name) + " License No."
        p2 = project_info.LookupParameter(license_param_name)
        if p2:
            p2.Set(str(license_id) if license_id else "")
            
        results.append("Updated: " + str(param_name))

    TransactionManager.Instance.TransactionTaskDone()
    workbook.Close(False)
    OUT = results

except Exception as ex:
    OUT = "เกิดข้อผิดพลาด: " + str(ex)

finally:
    # สำคัญ: ต้องปิดกระบวนการ Excel ในเครื่องเสมอ
    excel_app.Quit()