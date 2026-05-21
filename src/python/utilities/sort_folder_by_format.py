import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import sys

# กำหนดหมวดหมู่ของไฟล์เพื่อความเป็นระเบียบมากขึ้น
FILE_CATEGORIES = {
    "Images": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg", ".webp"],
    "Documents": [".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt", ".csv"],
    "Videos": [".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv"],
    "Audio": [".mp3", ".wav", ".aac", ".flac", ".ogg"],
    "Archives": [".zip", ".rar", ".7z", ".tar", ".gz"],
    "Programs": [".exe", ".msi", ".bat", ".cmd"],
    "Code": [".py", ".html", ".css", ".js", ".json", ".cpp", ".c", ".java"]
}

def get_category(extension):
    """ คืนค่าชื่อหมวดหมู่จากนามสกุลไฟล์ ถ้าไม่เจอจะคืนค่าเป็น 'Others' """
    extension = extension.lower()
    for category, exts in FILE_CATEGORIES.items():
        if extension in exts:
            return category
    # หากไม่ตรงหมวดไหนเลย ให้ใช้นามสกุลไฟล์เป็นชื่อโฟลเดอร์ (หรือจะใช้คำว่า "Others" ก็ได้)
    return extension[1:].upper()

def organize_files(target_path=None):
    # ถ้าไม่ได้ส่ง path มาให้เปิดหน้าต่างเลือกโฟลเดอร์
    if not target_path:
        root = tk.Tk()
        root.withdraw()
        target_path = filedialog.askdirectory(title="เลือกโฟลเดอร์ที่ต้องการจัดระเบียบไฟล์")

    # ถ้าผู้ใช้ไม่ได้เลือก (กด Cancel)
    if not target_path:
        return

    try:
        # เก็บรายชื่อไฟล์ในโฟลเดอร์
        files = [f for f in os.listdir(target_path) if os.path.isfile(os.path.join(target_path, f))]
        
        count = 0
        for file in files:
            # ป้องกันไม่ให้โปรแกรมย้ายไฟล์สคริปต์ของตัวมันเอง (กรณีรันในโฟลเดอร์เดียวกัน)
            if file == os.path.basename(__file__):
                continue

            name, extension = os.path.splitext(file)
            
            # ข้ามไฟล์ที่ไม่มีนามสกุล
            if not extension:
                continue

            # กำหนดชื่อโฟลเดอร์ตามหมวดหมู่
            folder_name = get_category(extension)
            folder_full_path = os.path.join(target_path, folder_name)

            # สร้างโฟลเดอร์ใหม่ถ้ายังไม่มี
            if not os.path.exists(folder_full_path):
                os.makedirs(folder_full_path)

            # ย้ายไฟล์
            shutil.move(os.path.join(target_path, file), os.path.join(folder_full_path, file))
            count += 1

        # แสดงข้อความเมื่อทำงานเสร็จ (แสดงเฉพาะเมื่อไม่ได้รันผ่าน command line เงียบๆ)
        messagebox.showinfo("สำเร็จ", f"จัดระเบียบโฟลเดอร์เรียบร้อยแล้ว!\nจำนวนไฟล์ที่ถูกย้าย: {count} ไฟล์")

    except Exception as e:
        messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถดำเนินการได้:\n{e}")

if __name__ == "__main__":
    # รองรับการส่ง Path เข้ามาผ่าน Command Line (มีประโยชน์ถ้านำไปทำเมนูคลิกขวา)
    if len(sys.argv) > 1:
        folder_to_sort = sys.argv[1]
        if os.path.isdir(folder_to_sort):
            organize_files(folder_to_sort)
        else:
            messagebox.showwarning("คำเตือน", "Path ที่ระบุไม่ใช่โฟลเดอร์")
    else:
        organize_files()