import clr
import urllib.request
import time
import os

# --- [SETTING] Path ต่างๆ ---
# 1. Path เก็บ Token 
token_file_path = r"R:\05_MTC_Library Development\MTC R24\00_BIM_Tools_Library\01_Python-Scripts\02_Git Code Importer\00_Data\ICB_Note.txt"
# 2. URL ของสคริปต์หลักบน GitHub
url = "https://raw.githubusercontent.com/Ittib/RevitFilterDeleteTools/refs/heads/main/MTC_Filter%20Cleaner%20Tools.py"
# 3. Path ของสคริปต์สำรอง (Local Fallback)
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

    # ดึงจาก GitHub
    req_url = url + "?t=" + str(int(time.time()))
    req = urllib.request.Request(req_url)
    req.add_header('Authorization', 'token %s' % token)

    with urllib.request.urlopen(req, timeout=10) as response:
        script_content = response.read().decode('utf-8')
        source_info = "🌐 Loaded from GitHub (Latest)"

except Exception as e:
    # --- [LOCAL FALLBACK] หากเน็ตหลุดหรือ Token มีปัญหา ---
    if os.path.exists(local_fallback_path):
        with open(local_fallback_path, 'r', encoding='utf-8') as f:
            script_content = f.read()
            source_info = "⚠️ Network Offline: Loaded from Local Fallback"
    else:
        source_info = "❌ Error: Cannot reach GitHub and Local file is missing!"

# --- รันโค้ดที่ดึงมาได้ ---
if script_content:
    try:
        exec(script_content)
        OUT = source_info
    except Exception as run_e:
        OUT = "❌ Script Execution Error: " + str(run_e)
else:
    OUT = source_info