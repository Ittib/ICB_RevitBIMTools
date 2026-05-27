import os
import datetime
from collections import Counter

# 🎯 ตั้งค่าโฟลเดอร์ที่ต้องการสแกน
TARGET_DIR = r"R:\01_MTC_Library\MTC R24 (Update)\03_ME_Library" 
OUTPUT_FILE = r"C:\Users\ittichai.b\Downloads\Full_Server_Structure_Report.txt"

def generate_tree(startpath):
    tree_output = []
    tree_output.append(f"🔍 สแกนโครงสร้างเซิร์ฟเวอร์แบบเจาะลึก (Unlimited Depth & All Files)")
    tree_output.append(f"📅 วันเวลาที่สแกน: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    tree_output.append(f"📁 เป้าหมาย: {startpath}")
    tree_output.append("="*80 + "\n")

    if not os.path.exists(startpath):
        return "❌ หาโฟลเดอร์ไม่เจอครับ ตรวจสอบ Path ใน TARGET_DIR อีกครั้งนะครับ"

    # os.walk จะมุดลงไปเรื่อยๆ จนกว่าจะสุดทาง
    for root, dirs, files in os.walk(startpath):
        level = root.replace(startpath, '').count(os.sep)
        indent = '    ' * level 
        folder_name = os.path.basename(root) if root != startpath else os.path.basename(startpath)
        
        file_info = ""
        if files:
            # ดึงนามสกุลไฟล์ออกมาทั้งหมด แล้วนับจำนวน
            extensions = []
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                extensions.append(ext if ext else "ไฟล์ไม่มีนามสกุล")
            
            ext_counts = Counter(extensions)
            
            # จัดรูปแบบให้อ่านง่าย เช่น "10 .rfa, 2 .dwg"
            ext_str = ", ".join([f"{count} {ext}" for ext, count in ext_counts.items()])
            file_info = f"  ---> [พบไฟล์: {ext_str}]"
            
        tree_output.append(f"{indent}📂 {folder_name}{file_info}")

    return '\n'.join(tree_output)

def main():
    print("⏳ กำลังทะลวงสแกนโครงสร้างเซิร์ฟเวอร์ทั้งหมด กรุณารอสักครู่ (อาจใช้เวลาแป๊บโหลนะ)...")
    
    # 1. เริ่มสแกน
    report_data = generate_tree(TARGET_DIR)
    
    # 2. เซฟลงไฟล์ (ส่วนนี้ที่หายไปในโค้ดของคุณ)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(report_data)
        
    print(f"\n✅ สแกนเสร็จสมบูรณ์! บันทึกรายงานไว้ที่ไฟล์:\n{os.path.abspath(OUTPUT_FILE)}")
    print("\n💡 คุณสามารถเปิดไฟล์ Text นี้ขึ้นมาดูได้เลยครับ!")

# 3. ตัว Trigger ให้โปรแกรมเริ่มทำงาน (ส่วนนี้สำคัญมาก ห้ามลืมครับ)
if __name__ == "__main__":
    main()