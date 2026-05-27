# 🏗️ Navisworks Clash Reporter

ปลั๊กอินสำหรับ Autodesk Navisworks ที่ช่วยดึงข้อมูลจุดตัด (Clash) จาก Saved Viewpoints และส่งออกเป็นรายงาน (Report) พร้อมรูปภาพประกอบโดยอัตโนมัติ

## 📖 คู่มือการใช้งาน
โปรเจกต์นี้มีคู่มือการใช้งานและสรุปตรรกะการทำงานฉบับเต็มในรูปแบบ HTML สามารถเปิดดูได้จากลิงก์ด้านล่างครับ:
👉 [คลิกที่นี่เพื่อเปิดดูคู่มือ (Navisworks_Clash_Reporter_Manual.html)](bin/Release/net48/Navisworks_Clash_Reporter_Manual.html)

---

### ✨ ฟีเจอร์หลัก (Key Features)
* **ดึงข้อมูลมุมมอง (Extract Viewpoints):** กวาดข้อมูล View ทั้งหมดในโฟลเดอร์ที่ตั้งชื่อตามมาตรฐาน (กรองเอาเฉพาะตึกที่ผู้ใช้เลือกได้)
* **จับภาพอัตโนมัติ (Auto Screenshot):** สร้างไฟล์ภาพจากมุมมองและตั้งชื่อภาพตาม View นั้นๆ
* **โหมดการทำงาน 2 รูปแบบ:**
  - **XML / HTML Mode:** อ่านข้อมูลจาก Report ที่ Navisworks สร้างออกมา
  - **API Mode:** ใช้ Navisworks API ดึงข้อมูลและถ่ายภาพแบบ Real-time
* **เชื่อมต่อกับ Excel:** นำข้อมูลภาพและรายชื่อ View ไปจัดวางลงในไฟล์ Excel Template ให้โดยอัตโนมัติ (พร้อมข้าม Status ที่เป็น "Pending")

### 📂 โครงสร้างโปรเจกต์ (Folder Structure)
* `MainAddIn.cs` - ไฟล์เริ่มต้นสำหรับติดตั้ง Command เข้ากับปุ่มใน Navisworks
* `UI/` - หน้าต่าง WPF (XAML) สำหรับตั้งค่าและเลือกโฟลเดอร์ต่างๆ
* `Logic/` - คลาสสำหรับประมวลผล เช่น `ClashDataExtractor` (ดึงข้อมูล View), `ImageExporter` (จัดการเรื่องรูป), และ `ExcelReportUpdater` (อัปเดตไฟล์ Excel)
