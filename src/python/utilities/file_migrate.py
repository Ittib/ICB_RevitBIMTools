import shutil
import re
from pathlib import Path

def safe_move(src, dst):
    """
    ฟังก์ชันย้ายไฟล์/โฟลเดอร์ที่ทนทานขึ้น 
    รองรับการย้ายข้ามไดรฟ์ ข้ามไฟล์ที่เปิดอยู่ และข้ามไฟล์ที่ติด Permission
    """
    try:
        if src.is_dir():
            # สำหรับโฟลเดอร์ข้ามไดรฟ์: Copy ไปก่อน แล้วพยายามลบต้นทางแบบ Ignore Errors (ป้องกันสคริปต์พัง)
            shutil.copytree(src, dst, dirs_exist_ok=True)
            shutil.rmtree(src, ignore_errors=True) 
        else:
            # สำหรับไฟล์เดี่ยว
            shutil.move(str(src), str(dst))
        return True
    except PermissionError:
        print(f"  [ข้าม] ⚠️ ไฟล์หรือโฟลเดอร์นี้ถูกเปิดใช้งานอยู่: {src.name}")
        return False
    except Exception as e:
        print(f"  [ข้อผิดพลาด] ❌ ไม่สามารถย้าย {src.name}: {e}")
        return False

def organize_downloads_to_d_drive():
    source_dir = Path(r"C:\Users\ittichai.b\Downloads")
    target_dir = Path(r"D:\itti")
    
    # หมวดหมู่ไฟล์
    categories = {
        "BIM_Library": [".rfa", ".rvt", ".rte", ".rft"],
        "Automation_Scripts": [".py", ".dyn", ".cs", ".json", ".xml", ".uxml", ".pxml"],
        "References/CAD": [".dwg", ".dxf", ".dgn", ".pcp", ".pat"],
        "References/Documents": [".pdf", ".docx", ".doc", ".txt", ".csv", ".xlsx", ".xls", ".xlsm", ".log"],
        "References/Images": [".jpg", ".jpeg", ".png", ".webp", ".jfif"],
        "Software_Installers": [".exe", ".msi", ".zip", ".rar"]
    }
    
    # โฟลเดอร์เป้าหมาย
    trash_dir = target_dir / "Revit_Cleaned_Backups"
    projects_dir = target_dir / "Projects_and_Source"
    misc_dir = target_dir / "Others"

    # สร้างโครงสร้างโฟลเดอร์
    for category in categories.keys():
        (target_dir / category).mkdir(parents=True, exist_ok=True)
    trash_dir.mkdir(parents=True, exist_ok=True)
    projects_dir.mkdir(parents=True, exist_ok=True)
    misc_dir.mkdir(parents=True, exist_ok=True)

    # Regex สำหรับดักไฟล์ Backup ของ Revit
    revit_backup_pattern = re.compile(r'\.\d{4}\.(rvt|rfa|rte)$', re.IGNORECASE)

    print(f"กำลังเริ่มจัดระเบียบไฟล์จาก: {source_dir}")
    print(f"เป้าหมาย: {target_dir}\n" + "="*50)

    for item in source_dir.iterdir():
        if not item.exists():
            continue

        target_path = None

        # ---------------------------
        # จัดการโฟลเดอร์
        # ---------------------------
        if item.is_dir():
            if item.name.endswith("_backup") or item.name == "Revit_temp":
                target_path = trash_dir / item.name
            else:
                target_path = projects_dir / item.name
                
            print(f"[กำลังย้ายโฟลเดอร์] 📁 {item.name} ... (อาจใช้เวลาสักครู่)")
            if safe_move(item, target_path):
                print(f"  -> สำเร็จ: ย้ายไปที่ {target_path.parent.name}/")
                
        # ---------------------------
        # จัดการไฟล์เดี่ยว
        # ---------------------------
        elif item.is_file():
            if revit_backup_pattern.search(item.name):
                target_path = trash_dir / item.name
            else:
                file_ext = item.suffix.lower()
                moved = False
                for category, extensions in categories.items():
                    if file_ext in extensions:
                        target_path = target_dir / category / item.name
                        moved = True
                        break
                if not moved:
                    target_path = misc_dir / item.name

            # จัดการชื่อไฟล์ซ้ำ
            if target_path and target_path.exists():
                base = target_path.stem
                ext = target_path.suffix
                counter = 1
                while target_path.exists():
                    target_path = target_path.parent / f"{base}_{counter}{ext}"
                    counter += 1

            if safe_move(item, target_path):
                print(f"[ย้ายไฟล์] 📄 {item.name} -> {target_path.parent.name}/")

    print("=" * 50)
    print("จัดระเบียบเสร็จสมบูรณ์!")

if __name__ == "__main__":
    organize_downloads_to_d_drive()