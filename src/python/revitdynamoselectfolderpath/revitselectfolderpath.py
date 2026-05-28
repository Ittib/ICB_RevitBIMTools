import clr
import System

# โหลดไลบรารีสำหรับใช้งาน Windows UI
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import DialogResult, FolderBrowserDialog

# รับค่าสวิตช์เปิด/ปิดจาก IN[0] (ต่อ Boolean Node ค่า True ไว้เพื่อสั่งทำงาน)
run_it = IN[0]

if run_it:
    # สร้างหน้าต่างเลือกโฟลเดอร์
    dialog = FolderBrowserDialog()
    dialog.Description = "Please select a folder"
    dialog.ShowNewFolderButton = True # อนุญาตให้ผู้ใช้สร้างโฟลเดอร์ใหม่ได้ด้วย
    
    # ถ้าผู้ใช้เลือกโฟลเดอร์และกด OK
    if dialog.ShowDialog() == DialogResult.OK:
        # ส่งที่อยู่โฟลเดอร์ออกไปที่ OUT เพื่อไปใช้ต่อใน Node ถัดไป
        OUT = dialog.SelectedPath
    else:
        OUT = "Canceled"
else:
    OUT = "Set IN[0] to True to run"
