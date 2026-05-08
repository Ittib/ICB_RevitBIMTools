import clr
import time
import os

# นำเข้าไลบรารีของ .NET สำหรับจัดการ HTTP Request และ File
clr.AddReference('System')
from System.Net import WebClient, ServicePointManager, SecurityProtocolType
from System.Text import Encoding
from System.IO import File

# --- บังคับให้ใช้ TLS 1.2 --- 
# (สำคัญมาก: GitHub ปฏิเสธการเชื่อมต่อจาก SSL/TLS เวอร์ชั่นเก่า การตั้งค่านี้ช่วยให้ IronPython 2 โหลดข้อมูลได้)
ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

# --- [SETTING] Path ต่างๆ ---
token_file_path = r"R:\05_MTC_Library Development\MTC R24\00_BIM_Tools_Library\01_Python-Scripts\02_Git Code Importer\00_Data\ICB_Note.txt"
url = "https://raw.githubusercontent.com/Ittib/RevitFilterDeleteTools/refs/heads/main/MTC_Filter%20Cleaner%20Tools.py"
local_fallback_path = r"R:\05_MTC_Library Development\MTC R24\00_BIM_Tools_Library\02_Dynamo_Scripts\ICB\Python Script Demo\MTC_Filter Cleaner Tools.py"

def get_token(path):
    try:
        with open(path, 'r') as f:
            return f.read().strip()
    except:
        return None

script_content = ""
source_info = ""

# --- เริ่มการดึงโค้ด ---
try:
    token = get_token(token_file_path)
    if not token:
        raise Exception("Token file not found")

    req_url = url + "?t=" + str(int(time.time()))

    # ใช้ .NET WebClient แทน urllib เพื่อหลีกเลี่ยงปัญหา HTTPS ของ IronPython 2
    client = WebClient()
    client.Headers.Add('Authorization', 'token %s' % token)
    client.Encoding = Encoding.UTF8

    script_content = client.DownloadString(req_url)
    source_info = "🌐 Loaded from GitHub (Latest)"

except Exception as e:
    # --- [LOCAL FALLBACK] หากเน็ตหลุดหรือ Token มีปัญหา ---
    if os.path.exists(local_fallback_path):
        try:
            # ใช้ System.IO.File ของ .NET แทนการใช้ open() แบบเก่า เพื่อให้อ่านภาษาไทย/UTF-8 ได้ไม่ Error
            script_content = File.ReadAllText(local_fallback_path, Encoding.UTF8)
            source_info = "⚠️ Network Offline: Loaded from Local Fallback"
        except Exception as read_e:
            source_info = "❌ Error: Cannot read Local Fallback file! (" + str(read_e) + ")"
    else:
        source_info = "❌ Error: Cannot reach GitHub and Local file is missing! (" + str(e) + ")"

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