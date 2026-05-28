import clr
import System
import os
import re
import random

# Import WPF/Windows API for UI
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("System.Xml")
clr.AddReference("System.Windows.Forms")

from System.Windows.Markup import XamlReader
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows.Controls import CheckBox
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

# Import RevitAPI
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Import DocumentManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager

# Import Excel Interop
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

# รับค่าจาก Boolean Node (True = ทำงาน)
run_it = IN[0] if 'IN' in globals() else True

OUT = []
log = []

# ==========================================
# 1. UI XAML Config
# ==========================================
XAML_UI = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Active Project Family Parameter Extractor" Width="450" Height="550"
        Background="#F5F5F5" Topmost="True" WindowStartupLocation="CenterScreen">
    <StackPanel Margin="15">
        <Label Content="1. เลือก Family ที่ต้องการดึงข้อมูล:" FontWeight="Bold" Background="#E0E0E0" Padding="5"/>
        <TextBlock Text="พิมพ์ค้นหา Family:" Margin="0,8,0,2"/>
        <TextBox Name="txtSearch" Height="25" Margin="0,0,0,5" />
        <StackPanel Orientation="Horizontal" Margin="0,0,0,5">
            <Button Name="btnSelectAll" Content="Select All" Width="80" Height="22" Margin="0,0,5,0"/>
            <Button Name="btnClearAll"  Content="Clear All"  Width="80" Height="22"/>
        </StackPanel>
        <ListBox Name="lbFamilies" Height="300" Margin="0,0,0,10"/>
        
        <Button Name="btnRun" Content="EXTRACT PARAMETERS TO EXCEL" Height="45" Background="#212121" Foreground="#FFFFFF" FontWeight="Bold"/>
    </StackPanel>
</Window>
"""

# ==========================================
# 2. UI Controller
# ==========================================
class FamilySelectionUI(object):
    def __init__(self, family_names):
        self.checkbox_map = {}
        self.sorted_names = sorted(family_names)
        self.selected_families = []
        
        self.window = XamlReader.Load(XmlReader.Create(StringReader(XAML_UI)))
        
        # Bind Controls
        self.txtSearch = self.window.FindName("txtSearch")
        self.lbFamilies = self.window.FindName("lbFamilies")
        self.btnSelectAll = self.window.FindName("btnSelectAll")
        self.btnClearAll = self.window.FindName("btnClearAll")
        self.btnRun = self.window.FindName("btnRun")
        
        # Events
        self.txtSearch.TextChanged += self.on_search_changed
        self.btnSelectAll.Click += self.on_select_all
        self.btnClearAll.Click += self.on_clear_all
        self.btnRun.Click += self.on_submit
        
        # Init List
        self.refresh_list("")

    def refresh_list(self, search_text):
        self.lbFamilies.Items.Clear()
        search_text = search_text.lower()
        for name in self.sorted_names:
            if search_text in name.lower():
                # รักษาสถานะเดิมถ้าเคยติ๊กไว้
                prev_checked = self.checkbox_map[name].IsChecked if name in self.checkbox_map else False
                cb = CheckBox()
                cb.Content = name
                cb.IsChecked = prev_checked
                self.checkbox_map[name] = cb
                self.lbFamilies.Items.Add(cb)

    def on_search_changed(self, sender, e):
        self.refresh_list(self.txtSearch.Text)

    def on_select_all(self, sender, e):
        for cb in self.checkbox_map.values():
            cb.IsChecked = True

    def on_clear_all(self, sender, e):
        for cb in self.checkbox_map.values():
            cb.IsChecked = False

    def on_submit(self, sender, e):
        self.selected_families = [name for name, cb in self.checkbox_map.items() if cb.IsChecked]
        self.window.DialogResult = True
        self.window.Close()


def rgb_to_excel(r, g, b):
    return r + (g * 256) + (b * 65536)


def main_process():
    doc = DocumentManager.Instance.CurrentDBDocument
    
    # 1. ค้นหา Family ทั้งหมดที่โหลดอยู่ใน Project และแก้ไขได้
    families = FilteredElementCollector(doc).OfClass(Family).ToElements()
    editable_families = {fam.Name: fam for fam in families if fam.IsEditable}
    
    if not editable_families:
        return ["No editable families found in the active project.", log]
        
    # 2. แสดง UI ให้เลือก Family
    ui = FamilySelectionUI(editable_families.keys())
    if not ui.window.ShowDialog():
        return ["User canceled family selection.", log]
        
    selected_fam_names = ui.selected_families
    if not selected_fam_names:
        return ["No families selected.", log]
        
    family_type_data_map = {} # dict of list of dicts: { "FamilyName": [ {type_data}, ... ] }
    
    # 3. วนลูปเปิดไฟล์ Family (EditFamily) ทีละไฟล์เพื่อดึง Parameter ของ "ทุก Type"
    for fam_name in selected_fam_names:
        fam = editable_families[fam_name]
        all_type_data = []
        all_param_names = []
        
        try:
            fam_doc = doc.EditFamily(fam)
            
            if fam_doc and fam_doc.IsFamilyDocument:
                fam_manager = fam_doc.FamilyManager
                
                # วนลูปดึงข้อมูลจากแต่ละ Type ใน Family
                for fam_type in fam_manager.Types:
                    type_name = fam_type.Name
                    type_data = {"Family Name": fam_name, "Type Name": type_name}
                    
                    # ดึงพารามิเตอร์ทั้งหมดใน Family นั้น
                    for param in fam_manager.Parameters:
                        base_param_name = param.Definition.Name
                        
                        # ตรวจสอบชนิดของ Parameter
                        if param.IsShared:
                            p_type = "Shared"
                        elif param.Id.IntegerValue < 0:
                            p_type = "BuiltIn"
                        else:
                            p_type = "Family"
                            
                        param_name_formatted = "{} [{}]".format(base_param_name, p_type)
                        val = ""
                        
                        # ดึงค่าพารามิเตอร์ของ Type นั้นๆ
                        if fam_type.HasValue(param):
                            storage_type = param.StorageType
                            try:
                                val_string = fam_type.AsValueString(param)
                            except:
                                val_string = None
                            
                            if val_string:
                                val = val_string
                            else:
                                if storage_type == StorageType.Double:
                                    val = fam_type.AsDouble(param)
                                elif storage_type == StorageType.Integer:
                                    val = fam_type.AsInteger(param)
                                elif storage_type == StorageType.String:
                                    val = fam_type.AsString(param)
                                elif storage_type == StorageType.ElementId:
                                    val = fam_type.AsElementId(param).IntegerValue
                        
                        type_data[param_name_formatted] = str(val) if val is not None else ""
                        
                        if param_name_formatted not in all_param_names:
                            all_param_names.append(param_name_formatted)
                    
                    all_type_data.append(type_data)
                
                fam_doc.Close(False)
                
                if all_type_data:
                    # เรียงลำดับ (Sort) ข้อมูลตามชื่อ Type Name
                    all_type_data.sort(key=lambda x: x.get("Type Name", ""))
                    
                    family_type_data_map[fam_name] = {
                        "types": all_type_data,
                        "params": all_param_names
                    }
                
        except Exception as e:
            log.append("Error opening family: {} | Msg: {}".format(fam_name, str(e)))
                
    if not family_type_data_map:
        return ["No parameters extracted.", log]

    # 4. สร้างและเขียนลง Excel (แยก Sheet ตาม Family)
    try:
        excel = Excel.ApplicationClass()
        excel.Visible = True 
        
        excel.ScreenUpdating = False
        excel.Interactive = False
        
        workbook = excel.Workbooks.Add()
        
        first_sheet = True
        
        # วนลูปตามแต่ละ Family เพื่อสร้าง Sheet ใหม่
        for fam_name, fam_info in family_type_data_map.items():
            if first_sheet:
                worksheet = workbook.Worksheets[1]
                first_sheet = False
            else:
                worksheet = workbook.Worksheets.Add(After=workbook.Worksheets[workbook.Worksheets.Count])
                
            # ตั้งชื่อ Sheet (ตัดตัวอักษรไม่ให้เกิน 31 ตัวและลบอักขระพิเศษตามข้อจำกัดของ Excel)
            safe_sheet_name = re.sub(r'[\\\\/*?:\\[\\]]', '', fam_name)
            if len(safe_sheet_name) > 31:
                safe_sheet_name = safe_sheet_name[:28] + "..."
            if not safe_sheet_name:
                safe_sheet_name = "Family"
            worksheet.Name = safe_sheet_name
            
            # ใส่สีที่ Tab (สุ่มสีโทนพาสเทลเพื่อให้มองเห็นชัดเจน)
            r = random.randint(150, 255)
            g = random.randint(150, 255)
            b = random.randint(150, 255)
            worksheet.Tab.Color = rgb_to_excel(r, g, b)
            
            all_type_data = fam_info["types"]
            all_param_names = fam_info["params"]
                        
            # จัดกลุ่ม (Sort) ตามชนิดของ Parameter
            all_param_names.sort(key=lambda x: (x.split(' [')[-1], x))
            
            final_param_names = []
            current_type = None
            for name in all_param_names:
                p_type = name.split(' [')[-1]
                if current_type is not None and p_type != current_type:
                    final_param_names.append("") # คอลัมน์ว่างสำหรับคั่น
                final_param_names.append(name)
                current_type = p_type
                
            # เขียน Headers
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
                
                if "[BuiltIn]" in header:
                    cell.Interior.Color = rgb_to_excel(51, 51, 153)
                elif "[Family]" in header:
                    cell.Interior.Color = rgb_to_excel(34, 139, 34)
                elif "[Shared]" in header:
                    cell.Interior.Color = rgb_to_excel(255, 140, 0)
                else:
                    cell.Interior.Color = rgb_to_excel(70, 70, 70)
                    
                cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                
            # เขียนข้อมูลลงตารางรวดเดียว (Bulk Write)
            import System
            rows_count = len(all_type_data)
            cols_count = len(headers)
            
            if rows_count > 0:
                data_arr = System.Array.CreateInstance(System.Object, rows_count, cols_count)
                for row_idx in range(rows_count):
                    type_data = all_type_data[row_idx]
                    for col_idx in range(cols_count):
                        header = headers[col_idx]
                        if col_idx == 0:
                            data_arr[row_idx, col_idx] = type_data.get("Family Name", "")
                        elif col_idx == 1:
                            data_arr[row_idx, col_idx] = type_data.get("Type Name", "")
                        elif header != "" and header in type_data:
                            data_arr[row_idx, col_idx] = str(type_data[header])
                            
                start_cell = worksheet.Cells[2, 1]
                end_cell = worksheet.Cells[rows_count + 1, cols_count]
                data_range = worksheet.Range(start_cell, end_cell)
                data_range.Value2 = data_arr
                
                # ตกแต่งสีพื้นหลัง
                if cols_count >= 3:
                    param_range = worksheet.Range(worksheet.Cells(2, 3), worksheet.Cells(rows_count + 1, cols_count))
                    param_range.Interior.Color = rgb_to_excel(220, 220, 220) 
                    try:
                        populated_cells = param_range.SpecialCells(2)
                        populated_cells.Interior.Color = rgb_to_excel(255, 255, 255)
                    except:
                        pass
            
            # ปรับขนาดคอลัมน์อัตโนมัติและตีเส้นตาราง
            worksheet.Columns.AutoFit()
            used_range = worksheet.UsedRange
            used_range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            used_range.Borders.Weight = Excel.XlBorderWeight.xlThin
            
            # สร้าง Group (Outline) คอลัมน์สำหรับแต่ละชนิดของ Parameter
            current_type = None
            start_col = 3
            
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
                    
            if current_type is not None and current_type != "Separator":
                end_col = len(final_param_names) + 2
                if start_col <= end_col:
                    worksheet.Range(worksheet.Cells(1, start_col), worksheet.Cells(1, end_col)).EntireColumn.Group()
        
        excel.ScreenUpdating = True
        excel.Interactive = True
        
        try:
            excel.WindowState = Excel.XlWindowState.xlMinimized
            excel.WindowState = Excel.XlWindowState.xlMaximized
        except:
            pass
        
        result_msg = "Success! Processed {} families into {} sheets.".format(len(family_type_data_map), len(family_type_data_map))
        if log:
            result_msg += "\n(With {} errors)".format(len(log))
            
        MessageBox.Show(result_msg, "Family Type Extractor: Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
        return [result_msg, log]
        
    except Exception as e:
        try:
            if 'excel' in locals():
                excel.ScreenUpdating = True
                excel.Interactive = True
        except:
            pass
        err_msg = "Error writing to Excel: " + str(e)
        MessageBox.Show(err_msg, "Family Type Extractor: Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        return [err_msg, log]


if run_it:
    OUT = main_process()
else:
    OUT = "Waiting... (Set IN[0] to True to run the script)"
