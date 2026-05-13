import clr
import os

from System.Text import Encoding
from System.IO import File

local_path = r"R:\05_MTC_Library Development\MTC R24\00_BIM_Tools_Library\02_Dynamo_Scripts\ICB\Python Script Demo\MTC_ChangeObject&LineStyleColor_1click.py"

script_content = ""
source_info = ""

    # --- [LOCAL]
if os.path.exists(local_path):
    try:
        # ใช้ System.IO.File ของ .NET แทนการใช้ open() แบบเก่า เพื่อให้อ่านภาษาไทย/UTF-8 ได้ไม่ Error
        script_content = File.ReadAllText(local_path, Encoding.UTF8)
        source_info = "📁 Loaded from Local"
    except Exception as read_e:
        source_info = "❌ Error: Cannot read Local Fallback file! (" + str(read_e) + ")"
else:
    source_info = "❌ Error: Cannot reach Local file is missing! "

# --- รันโค้ดที่ดึงมาได้ ---
if script_content:
    try:
        # ป้องกันปัญหาเรื่องรูปแบบการขึ้นบรรทัดใหม่ (\r\n) ที่อาจทำให้เกิด Syntax Error ในการรัน exec()
        script_content = script_content.replace("\r\n", "\n")
        exec(script_content)
        OUT = source_info
    except Exception as run_e:
        OUT = "❌ Script Execution Error: " + str(run_e)
else:
    OUT = source_info