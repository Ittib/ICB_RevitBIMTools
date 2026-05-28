import clr
import System
import os

# To use Windows UI for File Selection and Pop-ups
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
    # ใช้หน้าต่าง OpenFileDialog ของ Windows เพื่อเลือกไฟล์เดี่ยว
    dialog = OpenFileDialog()
    dialog.Title = "Select Revit Family (.rfa) files"
    dialog.Filter = "Revit Family Files (*.rfa)|*.rfa"
    dialog.Multiselect = True
    
    if dialog.ShowDialog() == DialogResult.OK:
        app = DocumentManager.Instance.CurrentUIApplication.Application
        
        all_type_data = []
        all_param_names = []
        
        for rfa_path in dialog.FileNames:
            try:
                model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(rfa_path)
                open_options = OpenOptions()
                
                # เปิดไฟล์แบบเบื้องหลัง
                doc = app.OpenDocumentFile(model_path, open_options)
                
                if doc.IsFamilyDocument:
                    fam_name = doc.Title
                    fam_manager = doc.FamilyManager
                    
                    # วนลูปดึงข้อมูลจากแต่ละ Type ใน Family
                    for fam_type in fam_manager.Types:
                        type_name = fam_type.Name
                        
                        type_data = {"Family Name": fam_name, "Type Name": type_name}
                        
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
                            
                            # ดึงค่าพารามิเตอร์ของ Type นั้นๆ
                            if fam_type.HasValue(param):
                                storage_type = param.StorageType
                                
                                # พยายามดึงค่าแบบข้อความที่แสดงบนหน้าจอ (พร้อมหน่วย) ก่อน
                                try:
                                    val_string = fam_type.AsValueString(param)
                                except:
                                    val_string = None
                                
                                if val_string:
                                    val = val_string
                                else:
                                    # ถ้ายึกค่าแบบติดหน่วยไม่ได้ ให้ดึงค่าดิบตามชนิดตัวแปร
                                    if storage_type == StorageType.Double:
                                        val = fam_type.AsDouble(param)
                                    elif storage_type == StorageType.Integer:
                                        val = fam_type.AsInteger(param)
                                    elif storage_type == StorageType.String:
                                        val = fam_type.AsString(param)
                                    elif storage_type == StorageType.ElementId:
                                        val = fam_type.AsElementId(param).IntegerValue
                            
                            type_data[param_name] = str(val) if val is not None else ""
                            
                            # บันทึกรายชื่อพารามิเตอร์เพื่อเอาไปสร้างเป็นคอลัมน์ Excel
                            if param_name not in all_param_names:
                                all_param_names.append(param_name)
                        
                        all_type_data.append(type_data)
                
                # ปิดไฟล์แบบไม่เซฟเพื่อล้าง Memory
                doc.Close(False)
                
            except Exception as e:
                log.append("Error processing file: {} | Msg: {}".format(rfa_path, str(e)))
        
        if all_type_data:
            # ถามผู้ใช้ว่าต้องการแยก Sheet ตามแฟมิลี่หรือไม่
            separate_res = MessageBox.Show(
                "Do you want to separate each Family into different Sheets?\\n(ต้องการแยก Sheet ตามชื่อแฟมิลี่หรือไม่?)",
                "Separate Sheets?",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            )
            separate_sheets = (separate_res == DialogResult.Yes)
            
            # เรียงลำดับตามชื่อ Family Name ก่อน แล้วตามด้วย Type Name (A-Z)
            all_type_data.sort(key=lambda x: (x.get("Family Name", ""), x.get("Type Name", "")))
            
            # ฟังก์ชันช่วยแปลงสี RGB (0-255) เป็นรหัสสีของ Excel
            def rgb_to_excel(r, g, b):
                return r + (g * 256) + (b * 65536)
                
            try:
                excel = Excel.ApplicationClass()
                excel.Visible = True 
                excel.ScreenUpdating = False
                excel.Interactive = False
                workbook = excel.Workbooks.Add()
                
                # --- NEW LOGIC FOR GROUPING ---
                groups = []
                if separate_sheets:
                    family_map = {}
                    for td in all_type_data:
                        f_name = td.get("Family Name", "Unknown")
                        if f_name not in family_map:
                            family_map[f_name] = []
                        family_map[f_name].append(td)
                    for f_name, t_list in family_map.items():
                        groups.append((f_name, t_list))
                else:
                    groups.append(("Type Parameters", all_type_data))
                
                first_sheet = True
                import re
                import random
                
                for sheet_title, types_list in groups:
                    if first_sheet:
                        worksheet = workbook.Worksheets[1]
                        first_sheet = False
                    else:
                        worksheet = workbook.Worksheets.Add(After=workbook.Worksheets[workbook.Worksheets.Count])
                        
                    safe_sheet_name = re.sub(r'[\\\\/*?:\\[\\]]', '', sheet_title)
                    if len(safe_sheet_name) > 31: safe_sheet_name = safe_sheet_name[:28] + "..."
                    worksheet.Name = safe_sheet_name if safe_sheet_name else "Sheet"
                    
                    if separate_sheets:
                        worksheet.Tab.Color = rgb_to_excel(random.randint(150, 255), random.randint(150, 255), random.randint(150, 255))
                    
                    group_params = []
                    for td in types_list:
                        for k in td.keys():
                            if k not in ["Family Name", "Type Name"] and k not in group_params:
                                group_params.append(k)
                    group_params.sort(key=lambda x: (x.split(' [')[-1], x))
                    
                    final_param_names = []
                    current_type = None
                    for name in group_params:
                        p_type = name.split(' [')[-1]
                        if current_type is not None and p_type != current_type: final_param_names.append("")
                        final_param_names.append(name)
                        current_type = p_type
                    
                    headers = ["Family Name", "Type Name"] + final_param_names
                    
                    for col_idx, header in enumerate(headers):
                        cell = worksheet.Cells[1, col_idx + 1]
                        if header == "":
                            worksheet.Columns[col_idx + 1].ColumnWidth = 2
                            cell.Interior.Color = rgb_to_excel(160, 160, 160)
                            continue
                        cell.Value2 = header
                        cell.Font.Bold = True
                        cell.Font.Color = rgb_to_excel(255, 255, 255)
                        if "[BuiltIn]" in header: cell.Interior.Color = rgb_to_excel(51, 51, 153)
                        elif "[Family]" in header: cell.Interior.Color = rgb_to_excel(34, 139, 34)
                        elif "[Shared]" in header: cell.Interior.Color = rgb_to_excel(255, 140, 0)
                        else: cell.Interior.Color = rgb_to_excel(70, 70, 70)
                        cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    
                    import System
                    rows_count = len(types_list)
                    cols_count = len(headers)
                    if rows_count > 0:
                        data_arr = System.Array.CreateInstance(System.Object, rows_count, cols_count)
                        for r in range(rows_count):
                            type_data = types_list[r]
                            for c in range(cols_count):
                                header = headers[c]
                                if c == 0: data_arr[r, c] = type_data.get("Family Name", "")
                                elif c == 1: data_arr[r, c] = type_data.get("Type Name", "")
                                elif header != "" and header in type_data: data_arr[r, c] = str(type_data[header])
                        worksheet.Range(worksheet.Cells(2, 1), worksheet.Cells(rows_count + 1, cols_count)).Value2 = data_arr
                        
                        if cols_count >= 3:
                            param_range = worksheet.Range(worksheet.Cells(2, 3), worksheet.Cells(rows_count + 1, cols_count))
                            param_range.Interior.Color = rgb_to_excel(220, 220, 220) 
                            try: param_range.SpecialCells(2).Interior.Color = rgb_to_excel(255, 255, 255)
                            except: pass
                    
                    worksheet.Columns.AutoFit()
                    used_range = worksheet.UsedRange
                    used_range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                    used_range.Borders.Weight = Excel.XlBorderWeight.xlThin
                    
                    current_type = None
                    start_col = 3
                    for i, param_name in enumerate(final_param_names):
                        col_idx_excel = i + 3
                        if param_name == "": p_type = "Separator"
                        elif "[BuiltIn]" in param_name: p_type = "BuiltIn"
                        elif "[Family]" in param_name: p_type = "Family"
                        elif "[Shared]" in param_name: p_type = "Shared"
                        else: p_type = "Unknown"
                        if p_type != current_type:
                            if current_type is not None and current_type != "Separator":
                                if start_col <= col_idx_excel - 1: worksheet.Range(worksheet.Cells(1, start_col), worksheet.Cells(1, col_idx_excel - 1)).EntireColumn.Group()
                            current_type = p_type
                            start_col = col_idx_excel
                    if current_type is not None and current_type != "Separator":
                        if start_col <= len(final_param_names) + 2: worksheet.Range(worksheet.Cells(1, start_col), worksheet.Cells(1, len(final_param_names) + 2)).EntireColumn.Group()
                # --- END NEW LOGIC ---
                
                # เปิดการทำงานของหน้าจอกลับมา
                excel.ScreenUpdating = True
                excel.Interactive = True
                
                # บังคับให้หน้าต่าง Excel เด้งมาอยู่หน้าสุด
                try:
                    excel.WindowState = Excel.XlWindowState.xlMinimized
                    excel.WindowState = Excel.XlWindowState.xlMaximized
                except:
                    pass
                
                result_msg = "Success! Processed {} types and wrote to Excel.".format(len(all_type_data))
                if log:
                    result_msg += "\n(With {} errors)".format(len(log))
                
                MessageBox.Show(result_msg, "Family Type Extractor: Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                
                OUT = [result_msg, log]
                
            except Exception as e:
                try:
                    if 'excel' in locals():
                        excel.ScreenUpdating = True
                        excel.Interactive = True
                except:
                    pass
                err_msg = "Error writing to Excel: " + str(e)
                MessageBox.Show(err_msg, "Family Type Extractor: Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                OUT = err_msg
        else:
            if not log:
                msg = "No types found or unable to read."
                MessageBox.Show(msg, "Family Type Extractor", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                OUT = msg
            else:
                OUT = log
    else:
        OUT = "User canceled file selection."
else:
    OUT = "Waiting... (Set IN[0] to True to run the script)"
