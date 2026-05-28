import clr
import System
import os

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

# Import WPF/Windows API for UI
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("System.Xml")
clr.AddReference("System.Windows.Forms")

from System.Windows.Markup import XamlReader
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows.Controls import CheckBox

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
        <Label Content="1. เลือก Category ของ Family ที่ต้องการดึงข้อมูล:" FontWeight="Bold" Background="#E0E0E0" Padding="5"/>
        <TextBlock Text="พิมพ์ค้นหา Category:" Margin="0,8,0,2"/>
        <TextBox Name="txtSearch" Height="25" Margin="0,0,0,5" />
        <StackPanel Orientation="Horizontal" Margin="0,0,0,5">
            <Button Name="btnSelectAll" Content="Select All" Width="80" Height="22" Margin="0,0,5,0"/>
            <Button Name="btnClearAll"  Content="Clear All"  Width="80" Height="22"/>
        </StackPanel>
        <ListBox Name="lbCategories" Height="300" Margin="0,0,0,10"/>
        
        <Button Name="btnRun" Content="EXTRACT PARAMETERS TO EXCEL" Height="45" Background="#212121" Foreground="#FFFFFF" FontWeight="Bold"/>
    </StackPanel>
</Window>
"""

# ==========================================
# 2. UI Controller
# ==========================================
class CategorySelectionUI(object):
    def __init__(self, cat_names):
        self.checkbox_map = {}
        self.sorted_names = sorted(cat_names)
        self.selected_categories = []
        
        self.window = XamlReader.Load(XmlReader.Create(StringReader(XAML_UI)))
        
        # Bind Controls
        self.txtSearch = self.window.FindName("txtSearch")
        self.lbCategories = self.window.FindName("lbCategories")
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
        self.lbCategories.Items.Clear()
        search_text = search_text.lower()
        for name in self.sorted_names:
            if search_text in name.lower():
                # รักษาสถานะเดิมถ้าเคยติ๊กไว้
                prev_checked = self.checkbox_map[name].IsChecked if name in self.checkbox_map else False
                cb = CheckBox()
                cb.Content = name
                cb.IsChecked = prev_checked
                self.checkbox_map[name] = cb
                self.lbCategories.Items.Add(cb)

    def on_search_changed(self, sender, e):
        self.refresh_list(self.txtSearch.Text)

    def on_select_all(self, sender, e):
        for cb in self.checkbox_map.values():
            cb.IsChecked = True

    def on_clear_all(self, sender, e):
        for cb in self.checkbox_map.values():
            cb.IsChecked = False

    def on_submit(self, sender, e):
        self.selected_categories = [name for name, cb in self.checkbox_map.items() if cb.IsChecked]
        self.window.DialogResult = True
        self.window.Close()


def rgb_to_excel(r, g, b):
    return r + (g * 256) + (b * 65536)


def main_process():
    doc = DocumentManager.Instance.CurrentDBDocument
    
    # 1. ค้นหา Family ทั้งหมดที่โหลดอยู่ใน Project
    families = FilteredElementCollector(doc).OfClass(Family).ToElements()
    
    # จัดกลุ่ม Family ตาม Category
    category_family_map = {}
    for fam in families:
        if fam.IsEditable: # เช็คว่าสามารถ Edit Family ได้
            cat = fam.FamilyCategory
            if cat is not None:
                cat_name = cat.Name
                if cat_name not in category_family_map:
                    category_family_map[cat_name] = []
                category_family_map[cat_name].append(fam)
                
    if not category_family_map:
        return ["No editable families found in the active project.", log]
        
    # 2. แสดง UI ให้เลือก Category
    ui = CategorySelectionUI(category_family_map.keys())
    if not ui.window.ShowDialog():
        return ["User canceled category selection.", log]
        
    selected_cats = ui.selected_categories
    if not selected_cats:
        return ["No categories selected.", log]
        
    all_family_data = []
    
    # 3. วนลูปเปิดไฟล์ Family (EditFamily) ทีละไฟล์เพื่อดึง Parameter
    for cat_name in selected_cats:
        fams_in_cat = category_family_map[cat_name]
        
        # จัดเรียงชื่อ Family ให้เป็นระเบียบ
        fams_in_cat.sort(key=lambda x: x.Name)
        
        for fam in fams_in_cat:
            try:
                # Option B: เปิด Family เบื้องหลังเพื่อให้ได้โครงสร้างเป๊ะๆ 100%
                fam_doc = doc.EditFamily(fam)
                
                if fam_doc and fam_doc.IsFamilyDocument:
                    fam_name = fam.Name
                    
                    fam_data = {"Category": cat_name, "Family Name": fam_name}
                    
                    fam_manager = fam_doc.FamilyManager
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
                            if storage_type == StorageType.Double:
                                val = current_type.AsDouble(param)
                            elif storage_type == StorageType.Integer:
                                val = current_type.AsInteger(param)
                            elif storage_type == StorageType.String:
                                val = current_type.AsString(param)
                            elif storage_type == StorageType.ElementId:
                                val = current_type.AsElementId(param).IntegerValue
                        
                        fam_data[param_name] = str(val) if val is not None else ""
                    
                    all_family_data.append(fam_data)
                    
                    # ปิดไฟล์เบื้องหลังโดยไม่เซฟ
                    fam_doc.Close(False)
                    
            except Exception as e:
                log.append("Error opening family: {} | Msg: {}".format(fam.Name, str(e)))
                
    if not all_family_data:
        return ["No parameters extracted (maybe no editable families in selected categories).", log]

    # 4. สร้างและเขียนลง Excel (แยก Sheet ตาม Category)
    try:
        excel = Excel.ApplicationClass()
        excel.Visible = True 
        
        excel.ScreenUpdating = False
        excel.Interactive = False
        
        workbook = excel.Workbooks.Add()
        
        import random
        
        # จัดกลุ่มข้อมูลตาม Category
        data_by_cat = {}
        for d in all_family_data:
            c = d["Category"]
            if c not in data_by_cat:
                data_by_cat[c] = []
            data_by_cat[c].append(d)
            
        first_sheet = True
        
        # วนลูปตามแต่ละ Category เพื่อสร้าง Sheet ใหม่
        for cat_name, cat_data in data_by_cat.items():
            if first_sheet:
                worksheet = workbook.Worksheets[1]
                first_sheet = False
            else:
                worksheet = workbook.Worksheets.Add(After=workbook.Worksheets[workbook.Worksheets.Count])
                
            # ตั้งชื่อ Sheet (ตัดตัวอักษรไม่ให้เกิน 31 ตัวและลบอักขระพิเศษตามข้อจำกัดของ Excel)
            import re
            safe_sheet_name = re.sub(r'[\\\\/*?:\\[\\]]', '', cat_name)
            if len(safe_sheet_name) > 31:
                safe_sheet_name = safe_sheet_name[:28] + "..."
            if not safe_sheet_name:
                safe_sheet_name = "Category"
            worksheet.Name = safe_sheet_name
            
            # ใส่สีที่ Tab (สุ่มสีโทนพาสเทลเพื่อให้มองเห็นชัดเจน)
            r = random.randint(150, 255)
            g = random.randint(150, 255)
            b = random.randint(150, 255)
            worksheet.Tab.Color = rgb_to_excel(r, g, b)
            
            # หา Parameter ทั้งหมดเฉพาะใน Category นี้
            cat_param_names = []
            for d in cat_data:
                for k in d.keys():
                    if k not in ["Category", "Family Name", "IsSeparator"] and k not in cat_param_names:
                        cat_param_names.append(k)
                        
            # จัดกลุ่ม (Sort) ตามชนิดของ Parameter
            cat_param_names.sort(key=lambda x: (x.split(' [')[-1], x))
            
            final_param_names = []
            current_type = None
            for name in cat_param_names:
                p_type = name.split(' [')[-1]
                if current_type is not None and p_type != current_type:
                    final_param_names.append("") # คอลัมน์ว่างสำหรับคั่น
                final_param_names.append(name)
                current_type = p_type
                
            # เขียน Headers (ไม่ต้องมีคอลัมน์ Category แล้วเพราะแยก Sheet แล้ว)
            headers = ["Family Name"] + final_param_names
            
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
            rows_count = len(cat_data)
            cols_count = len(headers)
            
            if rows_count > 0:
                data_arr = System.Array.CreateInstance(System.Object, rows_count, cols_count)
                for r in range(rows_count):
                    fam_data = cat_data[r]
                    for c in range(cols_count):
                        header = headers[c]
                        if c == 0:
                            data_arr[r, c] = fam_data.get("Family Name", "")
                        elif header != "" and header in fam_data:
                            data_arr[r, c] = str(fam_data[header])
                            
                start_cell = worksheet.Cells[2, 1]
                end_cell = worksheet.Cells[rows_count + 1, cols_count]
                data_range = worksheet.Range(start_cell, end_cell)
                data_range.Value2 = data_arr
                
                # ตกแต่งสีพื้นหลัง
                if cols_count >= 2:
                    param_range = worksheet.Range(worksheet.Cells(2, 2), worksheet.Cells(rows_count + 1, cols_count))
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
            start_col = 2
            
            for i, param_name in enumerate(final_param_names):
                col_idx_excel = i + 2
                
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
                end_col = len(final_param_names) + 1
                if start_col <= end_col:
                    worksheet.Range(worksheet.Cells(1, start_col), worksheet.Cells(1, end_col)).EntireColumn.Group()
        
        excel.ScreenUpdating = True
        excel.Interactive = True
        
        result_msg = "Success! Processed {} families into {} sheets.".format(len(all_family_data), len(data_by_cat))
        if log:
            result_msg += " (With {} errors)".format(len(log))
            
        return [result_msg, log]
        
    except Exception as e:
        try:
            if 'excel' in locals():
                excel.ScreenUpdating = True
                excel.Interactive = True
        except:
            pass
        return ["Error writing to Excel: " + str(e), log]


if run_it:
    OUT = main_process()
else:
    OUT = "Waiting... (Set IN[0] to True to run the script)"
