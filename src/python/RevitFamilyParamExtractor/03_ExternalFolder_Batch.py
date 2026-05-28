import clr
import System
import os

# To use Windows UI for Folder Selection and Pop-ups
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import DialogResult, OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon

# Import RevitAPI
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

# Import DocumentManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager

# Import Excel Interop
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

# รับค่าจาก Boolean Node (True = ทำงาน)
run_it = IN[0]

OUT = []
log = []

if run_it:
    # ใช้เทคนิคยืมหน้าต่าง OpenFileDialog ของ Windows เพื่อให้ได้ UI แบบใหม่
    dialog = OpenFileDialog()
    dialog.Title = "Select a folder containing Revit Family (.rfa) files (Click Open)"
    dialog.ValidateNames = False
    dialog.CheckFileExists = False
    dialog.CheckPathExists = True
    dialog.FileName = "Select_This_Folder" # ใส่ชื่อหลอกไว้
    dialog.Filter = "Folders Only|*.none" # ซ่อนไฟล์ทั้งหมด ให้เห็นแต่โฟลเดอร์
    
    if dialog.ShowDialog() == DialogResult.OK:
        # ตัดชื่อไฟล์หลอกทิ้ง เอาแค่ชื่อโฟลเดอร์ที่ผู้ใช้ยืนอยู่
        folder_path = os.path.dirname(dialog.FileName)
        
        app = DocumentManager.Instance.CurrentUIApplication.Application
        
        # 2. ค้นหาไฟล์ .rfa ทั้งหมดแบบทะลุซับโฟลเดอร์
        rfa_files = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(".rfa"):
                    rfa_files.append(os.path.join(root, file))
        
        if not rfa_files:
            msg = "No .rfa files found in: " + folder_path
            MessageBox.Show(msg, "Family Extractor", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            OUT = msg
        else:
            all_family_data = []
            all_param_names = []
            
            # 3. วนลูปเปิดไฟล์ทีละไฟล์เพื่อดึง Parameter
            for rfa_path in rfa_files:
                try:
                    model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(rfa_path)
                    open_options = OpenOptions()
                    
                    # เปิดไฟล์แบบเบื้องหลัง
                    doc = app.OpenDocumentFile(model_path, open_options)
                    
                    if doc.IsFamilyDocument:
                        fam_name = doc.Title
                        
                        # เอาแค่ชื่อโฟลเดอร์เต็มรูปแบบที่สังกัดอยู่ (ไม่ต้องเอาชื่อไฟล์)
                        display_path = os.path.dirname(rfa_path)
                        
                        fam_data = {"File Path": display_path, "Family Name": fam_name}
                        
                        fam_manager = doc.FamilyManager
                        current_type = fam_manager.CurrentType
                        
                        # ดึงพารามิเตอร์ทั้งหมดใน Family นั้น
                        for param in fam_manager.Parameters:
                            base_param_name = param.Definition.Name
                            
                            # ตรวจสอบชนิดของ Parameter (Shared, BuiltIn/System, Family)
                            if param.IsShared:
                                p_type = "Shared"
                            elif param.Id.IntegerValue < 0:
                                p_type = "BuiltIn"
                            else:
                                p_type = "Family"
                                
                            param_name = "{} [{}]".format(base_param_name, p_type)
                            
                            val = ""
                            
                            # ดึงค่าพารามิเตอร์ของ Current Type
                            if current_type is not None and current_type.HasValue(param):
                                storage_type = param.StorageType
                                
                                # พยายามดึงค่าแบบข้อความที่แสดงบนหน้าจอ (พร้อมหน่วย) ก่อน
                                try:
                                    val_string = current_type.AsValueString(param)
                                except:
                                    val_string = None
                                
                                if val_string:
                                    val = val_string
                                else:
                                    # ถ้ายึกค่าแบบติดหน่วยไม่ได้ ให้ดึงค่าดิบตามชนิดตัวแปร
                                    if storage_type == StorageType.Double:
                                        val = current_type.AsDouble(param)
                                    elif storage_type == StorageType.Integer:
                                        val = current_type.AsInteger(param)
                                    elif storage_type == StorageType.String:
                                        val = current_type.AsString(param)
                                    elif storage_type == StorageType.ElementId:
                                        val = current_type.AsElementId(param).IntegerValue
                            
                            fam_data[param_name] = str(val) if val is not None else ""
                            
                            # บันทึกรายชื่อพารามิเตอร์เพื่อเอาไปสร้างเป็นคอลัมน์ Excel (Dynamic Column)
                            if param_name not in all_param_names:
                                all_param_names.append(param_name)
                        
                        all_family_data.append(fam_data)
                    
                    # ปิดไฟล์แบบไม่เซฟเพื่อล้าง Memory
                    doc.Close(False)
                    
                except Exception as e:
                    log.append("Error opening file: {} | Msg: {}".format(rfa_path, str(e)))
            
            # 4. สร้างและเขียนลง Excel โดยตรง
            
            # ฟังก์ชันช่วยแปลงสี RGB (0-255) เป็นรหัสสีของ Excel
            def rgb_to_excel(r, g, b):
                return r + (g * 256) + (b * 65536)
                
            try:
                excel = Excel.ApplicationClass()
                excel.Visible = True 
                
                # ปิดการอัปเดตหน้าจอและการโต้ตอบชั่วคราว เพื่อป้องกัน Error 0x800AC472 และทำให้เขียนไวขึ้น
                excel.ScreenUpdating = False
                excel.Interactive = False
                
                workbook = excel.Workbooks.Add()
                worksheet = workbook.Worksheets[1]
                worksheet.Name = "Family Parameters"
                
                # จัดกลุ่ม (Sort) ตามชนิดของ Parameter (BuiltIn, Family, Shared)
                all_param_names.sort(key=lambda x: (x.split(' [')[-1], x))
                
                # เพิ่มคอลัมน์คั่น (Separator) ระหว่างกลุ่มเพื่อบังคับให้ Excel แยก Group
                final_param_names = []
                current_type = None
                for name in all_param_names:
                    p_type = name.split(' [')[-1]
                    if current_type is not None and p_type != current_type:
                        final_param_names.append("") # คอลัมน์ว่างสำหรับคั่น
                    final_param_names.append(name)
                    current_type = p_type
                
                # เขียน Headers
                headers = ["File Path", "Family Name"] + final_param_names
                
                for col_idx, header in enumerate(headers):
                    cell = worksheet.Cells[1, col_idx + 1]
                    
                    if header == "":
                        # ตกแต่งคอลัมน์คั่น
                        worksheet.Columns[col_idx + 1].ColumnWidth = 2
                        cell.Interior.Color = rgb_to_excel(160, 160, 160) # สีเทาอ่อน (Light Grey)
                        continue
                        
                    cell.Value2 = header
                    
                    # ตกแต่งสี Header แยกตาม Group
                    cell.Font.Bold = True
                    cell.Font.Color = rgb_to_excel(255, 255, 255) # สีขาว
                    
                    if "[BuiltIn]" in header:
                        cell.Interior.Color = rgb_to_excel(51, 51, 153) # สีน้ำเงินเข้ม (Dark Blue)
                    elif "[Family]" in header:
                        cell.Interior.Color = rgb_to_excel(34, 139, 34) # สีเขียว (Dark Green)
                    elif "[Shared]" in header:
                        cell.Interior.Color = rgb_to_excel(255, 140, 0) # สีส้ม (Dark Orange)
                    else:
                        cell.Interior.Color = rgb_to_excel(70, 70, 70) # สีเทาเข้ม (Dark Grey)
                        
                    cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                
                # เรียงลำดับข้อมูลตาม File Path เพื่อให้ไฟล์ในโฟลเดอร์เดียวกันอยู่ติดกัน
                all_family_data.sort(key=lambda x: x.get("File Path", ""))
                
                # แทรกแถวคั่น (Blank Row) ระหว่างกลุ่มโฟลเดอร์ เพื่อบังคับให้ Excel แยก Group Outline ออกจากกัน
                new_all_family_data = []
                for i in range(len(all_family_data)):
                    if i > 0:
                        prev_path = all_family_data[i-1].get("File Path", "")
                        curr_path = all_family_data[i].get("File Path", "")
                        if prev_path != curr_path:
                            # หา Common Path เพื่อให้ Group แม่คลุมต่อเนื่อง แต่ Group ลูกขาดออกจากกัน
                            parts1 = prev_path.split('\\')
                            parts2 = curr_path.split('\\')
                            common = []
                            for a, b in zip(parts1, parts2):
                                if a == b:
                                    common.append(a)
                                else:
                                    break
                            common_path = "\\".join(common)
                            new_all_family_data.append({"File Path": common_path, "IsSeparator": True})
                    new_all_family_data.append(all_family_data[i])
                    
                all_family_data = new_all_family_data
                
                # เขียนข้อมูลลงตารางแบบรวดเดียว (Bulk Write) เพื่อแก้ปัญหาช้าจาก COM
                import System
                rows_count = len(all_family_data)
                cols_count = len(headers)
                
                if rows_count > 0:
                    data_arr = System.Array.CreateInstance(System.Object, rows_count, cols_count)
                    for r in range(rows_count):
                        fam_data = all_family_data[r]
                        
                        if fam_data.get("IsSeparator"):
                            # เป็นแถวคั่น ปล่อยให้เซลล์ว่างไว้
                            continue
                            
                        for c in range(cols_count):
                            header = headers[c]
                            if c == 0:
                                data_arr[r, c] = fam_data.get("File Path", "")
                            elif c == 1:
                                data_arr[r, c] = fam_data.get("Family Name", "")
                            elif header != "" and header in fam_data:
                                data_arr[r, c] = str(fam_data[header])
                                
                    start_cell = worksheet.Cells[2, 1]
                    end_cell = worksheet.Cells[rows_count + 1, cols_count]
                    data_range = worksheet.Range(start_cell, end_cell)
                    
                    # วางข้อมูลทั้งหมดใน 1 คำสั่ง (ไวมาก)
                    data_range.Value2 = data_arr
                    
                    # ตกแต่งสีพื้นหลังรวดเดียวด้วย SpecialCells
                    if cols_count >= 3:
                        param_range = worksheet.Range(worksheet.Cells(2, 3), worksheet.Cells(rows_count + 1, cols_count))
                        # เทสีเทาเป็นค่าเริ่มต้น (Parameter ที่ไม่มี)
                        param_range.Interior.Color = rgb_to_excel(220, 220, 220) 
                        try:
                            # xlCellTypeConstants = 2 เลือกเฉพาะช่องที่มีข้อมูล เพื่อเปลี่ยนเป็นสีขาว
                            populated_cells = param_range.SpecialCells(2)
                            populated_cells.Interior.Color = rgb_to_excel(255, 255, 255) # สีขาว
                        except:
                            pass
                
                # ปรับขนาดคอลัมน์อัตโนมัติและตีเส้นตาราง
                worksheet.Columns.AutoFit()
                used_range = worksheet.UsedRange
                used_range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                used_range.Borders.Weight = Excel.XlBorderWeight.xlThin
                
                # สร้าง Group (Outline) สำหรับแต่ละชนิดของ Parameter แบบยุบ/ขยายได้ (+/-)
                current_type = None
                start_col = 3 # พารามิเตอร์เริ่มต้นที่คอลัมน์ 3 (คอลัมน์ C)
                
                for i, param_name in enumerate(final_param_names):
                    col_idx_excel = i + 3
                    
                    if param_name == "":
                        p_type = "Separator"
                    elif "[BuiltIn]" in param_name:
                        p_type = "BuiltIn"
                    elif "[Family]" in param_name:
                        p_type = "Family"
                    elif "[Shared]" in param_name:
                        p_type = "Shared"
                    else:
                        p_type = "Unknown"
                        
                    if p_type != current_type:
                        if current_type is not None and current_type != "Separator":
                            end_col = col_idx_excel - 1
                            if start_col <= end_col:
                                worksheet.Range(worksheet.Cells(1, start_col), worksheet.Cells(1, end_col)).EntireColumn.Group()
                        current_type = p_type
                        start_col = col_idx_excel
                
                # Group ชุดสุดท้าย
                if current_type is not None and current_type != "Separator":
                    end_col = len(final_param_names) + 2
                    if start_col <= end_col:
                        worksheet.Range(worksheet.Cells(1, start_col), worksheet.Cells(1, end_col)).EntireColumn.Group()
                
                # สร้าง Group สำหรับแถว (Row Grouping) ตามลำดับชั้นของโฟลเดอร์ (รองรับสูงสุด 7 ระดับตามข้อจำกัด Excel)
                if len(all_family_data) > 0:
                    max_depth = max([len(d.get("File Path", "").split('\\')) for d in all_family_data])
                    max_depth = min(7, max_depth) # Excel limit 8 levels total
                    
                    for lvl in range(max_depth):
                        start_row = None
                        current_val = None
                        
                        for row_offset in range(len(all_family_data)):
                            rfa_path = all_family_data[row_offset].get("File Path", "")
                            if not rfa_path: continue
                            
                            parts = rfa_path.split('\\') # ไม่ต้องตัด -1 แล้ว เพราะเอาเฉพาะโฟลเดอร์มาตั้งแต่ต้น
                            
                            # ตรวจสอบว่าไฟล์นี้อยู่ในโฟลเดอร์ลึกพอสำหรับระดับนี้ไหม
                            if lvl < len(parts):
                                val = "\\".join(parts[:lvl+1])
                            else:
                                val = None
                            
                            if val != current_val:
                                # ปิด Group ก่อนหน้า
                                if current_val is not None and start_row is not None:
                                    end_row = row_offset + 1
                                    if start_row <= end_row:
                                        worksheet.Range(worksheet.Cells(start_row, 1), worksheet.Cells(end_row, 1)).EntireRow.Group()
                                
                                current_val = val
                                if val is not None:
                                    start_row = row_offset + 2
                                else:
                                    start_row = None
                                
                        # ปิด Group สุดท้ายที่ยังเปิดค้างอยู่
                        if current_val is not None and start_row is not None:
                            end_row = len(all_family_data) + 1
                            if start_row <= end_row:
                                worksheet.Range(worksheet.Cells(start_row, 1), worksheet.Cells(end_row, 1)).EntireRow.Group()
                
                # เปิดการทำงานของหน้าจอกลับมา
                excel.ScreenUpdating = True
                excel.Interactive = True
                
                # บังคับให้หน้าต่าง Excel เด้งมาอยู่หน้าสุด (Force to Foreground)
                try:
                    excel.WindowState = Excel.XlWindowState.xlMinimized
                    excel.WindowState = Excel.XlWindowState.xlMaximized
                except:
                    pass
                
                result_msg = "Success! Processed {} families and wrote to Excel.".format(len(all_family_data))
                if log:
                    result_msg += "\n(With {} errors)".format(len(log))
                
                # แสดง Pop-up แจ้งเตือนเมื่อทำงานเสร็จ (เหมาะสำหรับรันใน Dynamo Player)
                MessageBox.Show(result_msg, "Family Extractor: Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                
                OUT = [result_msg, log]
                
            except Exception as e:
                try:
                    if 'excel' in locals():
                        excel.ScreenUpdating = True
                        excel.Interactive = True
                except:
                    pass
                err_msg = "Error writing to Excel: " + str(e)
                MessageBox.Show(err_msg, "Family Extractor: Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                OUT = err_msg
    else:
        OUT = "User canceled folder selection."
else:
    OUT = "Waiting... (Set IN[0] to True to run the script)"
 