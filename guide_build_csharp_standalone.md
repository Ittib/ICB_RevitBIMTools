# คู่มือการ Build โปรเจกต์ C# .NET (แบบไม่ต้องใช้สิทธิ์ Admin)

คู่มือนี้จัดทำขึ้นเพื่อเป็นแนวทางสำหรับกรณีที่คุณต้องการเขียนโปรแกรมหรือสร้าง Navisworks Add-in (C#) โปรเจกต์อื่นๆ ในอนาคต โดยใช้ตัวแปลภาษา (Compiler) แบบ Standalone ที่ติดตั้งไว้ให้แล้วบนเครื่องของคุณ

## 📌 1. ทำไมถึงใช้การ Build แบบนี้?
ปกติการเขียนและ Build โค้ด C# จำเป็นต้องพึ่งพาโปรแกรมขนาดใหญ่อย่าง **Visual Studio** หรือ **.NET SDK** แบบติดตั้งเต็มรูปแบบ ซึ่งมักจะติดข้อจำกัดเรื่อง **สิทธิ์ผู้ดูแลระบบ (Administrator)** บนเครื่องคอมพิวเตอร์ของบริษัท 

ดังนั้น เราจึงแก้ไขปัญหาโดยการนำ .NET SDK (เวอร์ชันแตกไฟล์) มาวางไว้ในโฟลเดอร์ส่วนตัวของคุณ เพื่อให้คุณสามารถคอมไพล์โค้ดได้ด้วยตัวเอง 100%

* **ตำแหน่งของ Compiler ของคุณ:** `D:\itti\mtc\dotnet-sdk-10\dotnet.exe`

---

## 🛠️ 2. วิธี Build แบบใช้ไฟล์ สคริปต์ (.bat) (แนะนำ)
วิธีที่ง่ายที่สุดเมื่อคุณขึ้นโปรเจกต์ใหม่ คือการสร้างไฟล์สคริปต์ไว้รันคำสั่งแทนการพิมพ์เองทุกครั้ง

1. ในโฟลเดอร์โปรเจกต์ใหม่ของคุณ (ที่มีไฟล์ `.csproj`) ให้คลิกขวา -> **New** -> **Text Document**
2. ตั้งชื่อไฟล์ว่า `build.bat` (ลบนามสกุล .txt ออก)
3. คลิกขวาที่ไฟล์ `build.bat` เลือก **Edit** (หรือเปิดด้วย Notepad/VS Code)
4. คัดลอกโค้ดด้านล่างนี้ไปวาง:

```bat
@echo off
echo =======================================================
echo Building My C# Plugin
echo =======================================================

:: กำหนดที่อยู่ของ .NET SDK บนเครื่อง
set DOTNET_EXE=D:\itti\mtc\dotnet-sdk-10\dotnet.exe

:: สั่ง Build โปรเจกต์
"%DOTNET_EXE%" build -c Release

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed! Check the errors above.
    pause
    exit /b %errorlevel%
)

echo.
echo [SUCCESS] Build completed successfully.
pause
```
5. เมื่อเขียนโค้ดเสร็จแล้ว คุณแค่ **ดับเบิลคลิก** ที่ไฟล์ `build.bat` ระบบก็จะทำการคอมไพล์โค้ดให้คุณทันที ไฟล์ `.dll` ที่ได้จะไปอยู่ในโฟลเดอร์ `bin\Release\net48\`

---

## 💻 3. วิธี Build ผ่าน Terminal ใน VS Code (แบบพิมพ์คำสั่ง)
หากคุณกำลังเขียนโค้ดอยู่ในโปรแกรม Visual Studio Code (VS Code) และไม่อยากสลับหน้าจอ คุณสามารถสั่ง Build ผ่าน Terminal ได้โดยตรง:

1. ใน VS Code กดเมนู **Terminal** -> **New Terminal**
2. พิมพ์คำสั่งนี้ลงไป แล้วกด Enter:
```powershell
& "D:\itti\mtc\dotnet-sdk-10\dotnet.exe" build -c Release
```

---

## ⚙️ 4. วิธีตั้งค่า VS Code Task (กดปุ่มเดียว Build ได้เลย)
เพื่อความโปรเหมือนนักพัฒนา المحترف คุณสามารถตั้งค่าให้ VS Code รู้จัก Compiler ตัวนี้ เพื่อให้กดปุ่ม `Ctrl + Shift + B` แล้ว Build ได้ทันที:

1. ในโฟลเดอร์โปรเจกต์ของคุณ ให้สร้างโฟลเดอร์ชื่อ `.vscode`
2. สร้างไฟล์ชื่อ `tasks.json` ไว้ข้างในโฟลเดอร์นั้น
3. นำโค้ดนี้ไปวาง:

```json
{
  "version": "2.0.0",
  "tasks": [
    {
      "label": "Build C# Project",
      "type": "shell",
      "command": "& \"D:\\itti\\mtc\\dotnet-sdk-10\\dotnet.exe\" build -c Release",
      "group": {
        "kind": "build",
        "isDefault": true
      },
      "problemMatcher": "$msCompile"
    }
  ]
}
```
> [!TIP]
> ครั้งต่อไปเมื่อคุณเขียนโค้ดเสร็จ เพียงแค่กด **`Ctrl + Shift + B`** บนคีย์บอร์ด VS Code จะทำการ Build ให้คุณทันที!

---

## 📦 5. กฎเหล็กของการทำไฟล์ .csproj ให้อ่าน Reference ได้
เพื่อให้ .NET SDK แบบ Standalone สามารถ Build โปรเจกต์ .NET Framework รุ่นเก่า (เช่น `net48`) ได้โดยไม่ต้องพึ่ง Visual Studio 
คุณต้องใส่บรรทัดนี้ลงไปในกลุ่ม `<ItemGroup>` ของไฟล์ `.csproj` เสมอ:

```xml
<PackageReference Include="Microsoft.NETFramework.ReferenceAssemblies" Version="1.0.3" PrivateAssets="All" />
```
(บรรทัดนี้จะสั่งให้โปรแกรมแอบไปดาวน์โหลดไฟล์ระบบของ Windows มาใช้สำหรับ Build ชั่วคราว)

---

## 📂 6. การนำ Plugin ไปติดตั้งใน Navisworks (แบบไม่มี Admin)
หลังจาก Build ได้ไฟล์ `.dll` แล้ว ปกติเราจะไม่สามารถเอาไปวางใน `C:\Program Files` ได้ ให้คุณนำไปวางที่โฟลเดอร์ส่วนตัวนี้แทนครับ:

`%APPDATA%\Autodesk\ApplicationPlugins`
*(หรือก็คือ `C:\Users\<ชื่อคอมพิวเตอร์ของคุณ>\AppData\Roaming\Autodesk\ApplicationPlugins`)*

**โครงสร้างที่ถูกต้องสำหรับ Navisworks:**
```text
ApplicationPlugins/
 └── ชื่อโปรเจกต์ของคุณ.bundle/
      └── Contents/
           ├── โปรเจกต์ของคุณ.dll
           └── ... (ไฟล์อ้างอิงอื่นๆ เช่น HtmlAgilityPack.dll)
```
เมื่อวางถูกโครงสร้าง Navisworks จะอ่าน Plugin โหลดขึ้นมาให้ตอนเปิดโปรแกรมทันทีครับ
