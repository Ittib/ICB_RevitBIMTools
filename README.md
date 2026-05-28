# 🏗️ ICB Revit BIM Tools 
### Custom Automation & Productivity Tools for Revit Ecosystem

ชุดเครื่องมือเสริมสำหรับ Revit ที่พัฒนาขึ้นเพื่อเพิ่มประสิทธิภาพการทำงาน (Workflow Automation) ภายในองค์กร ช่วยลดขั้นตอนการทำงานที่ซ้ำซ้อน และเพิ่มความแม่นยำในการจัดการข้อมูล BIM

---

### 🚀 Key Features
* **Automated Workflow:** ลดเวลาการทำงาน Manual ใน Revit ด้วยสคริปต์ที่ออกแบบมาเฉพาะทาง
* **Data Management:** จัดการข้อมูล Elements, Parameters และ Schedules ได้อย่างรวดเร็ว
* **Standardization:** ควบคุมมาตรฐานการทำงานให้เป็นไปในทิศทางเดียวกันทั้งทีม
* **User-Friendly:** ออกแบบมาให้เรียกใช้งานได้ง่ายผ่าน Dynamo Player หรือ Revit UI

### 📂 Project Structure
* **`src/csharp/`** : โปรเจกต์และเครื่องมือที่พัฒนาด้วย C# (เช่น NwdHtmlImportApp, MTC_1ClickPlot)
* **`src/python/apps/`** : แอปพลิเคชันและสคริปต์แยกส่วน (เช่น เครื่องมือเปลี่ยนชื่อไฟล์, auto_project_info)
* **`src/python/revit_tools/`** : เครื่องมือหลักสำหรับใช้งานใน Revit
* **`src/python/lib/`** : รวบรวม Python Scripts ที่เป็น Template File และ Dynamo Files (.dyn)
* **`src/python/bim_aec/`** : เครื่องมือและสคริปต์สำหรับจัดการข้อมูล BIM/AEC ทั่วไป
* **`src/python/utilities/`** : สคริปต์อรรถประโยชน์ต่างๆ สำหรับช่วยเหลือการทำงาน
* **`src/python/autocad/`** : สคริปต์ Python และ Lisp สำหรับการทำงานและสร้างไฟล์ร่วมกับ AutoCAD

### 🔧 Current Revit Tools
* **Duct Classification** (`src/revit_tools/DuctClassification`) ระบุชนิดของวัสดุลงสู่ Parameter โดยอัตโนมัติ
* **Filter Cleaner** (`src/revit_tools/FilterCleaner`) จัดการลบ Filter หลายๆชิ้นในโปรเจคได้ในทีเดียว
* **1clickMonoPrint** (`src/revit_tools/outdated/ObjectAndLineStyle`) Plot Drawing ให้เป็นสีดำโดยยังคง Object หรือ Line Style ที่เลือกให้เป็นสีเหมือนเดิม
* **MTC 1ClickPlot** (`src/csharp/MTC_1ClickPlot`) เครื่องมือสั่งพิมพ์ PDF แบบแยกแผ่นและเปลี่ยนสีอัจฉริยะ (ดู [คู่มือการใช้งาน](src/csharp/MTC_1ClickPlot/bin/Release/MTC_1ClickPlot_Manual.html))
* **Auto Project Info** (`src/python/apps/auto_project_info`) สคริปต์อัปเดต Project Info แบบอัตโนมัติจากไฟล์ Excel (ดู [คู่มือและตรรกะการทำงาน](src/python/apps/auto_project_info/Auto_Project_Info_Manual.html))

### 🏗️ Navisworks Tools
* **Navisworks Clash Reporter** (`src/csharp/plugins/NavisworksClashReporter`) ดึงข้อมูล View และรูปภาพจาก Navisworks ลง Excel อัตโนมัติ (ดู [คู่มือการใช้งาน](src/csharp/NavisworksClashReporter/bin/Release/net48/Navisworks_Clash_Reporter_Manual.html))