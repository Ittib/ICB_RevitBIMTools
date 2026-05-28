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
from System.Windows.Forms import DialogResult, OpenFileDialog, MessageBox, MessageBoxButtons, MessageBoxIcon

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
# 0. Main Selection UI XAML Config
# ==========================================
XAML_MAIN_UI = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Revit Parameter Extractor - Select Mode" Width="430" Height="340"
        Background="#F5F5F5" Topmost="True" WindowStartupLocation="CenterScreen">
    <StackPanel Margin="20">
        <Label Content="Please select the extraction mode:" FontWeight="Bold" FontSize="14" Margin="0,0,0,10"/>
        
        <RadioButton Name="rbActiveFam" Content="1. Active Project: Select by Family" IsChecked="True" Margin="5"/>
        <RadioButton Name="rbActiveCat" Content="2. Active Project: Select by Category" Margin="5"/>
        <RadioButton Name="rbBatchFolder" Content="3. External Files: Batch process Folder (.rfa)" Margin="5"/>
        <RadioButton Name="rbSingleFile" Content="4. External Files: Select Multiple Files (.rfa)" Margin="5,5,5,5"/>
        <CheckBox Name="chkSeparateSheets" Content="  ↳ Separate sheets by Family (แยกหน้าตามแฟมิลี่)" Margin="25,0,5,20" IsEnabled="False"/>
        
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Stretch" Margin="0,10,0,0">
            <Button Name="btnGuide" Content="📖 อ่านคู่มือ (Manual)" Height="40" Width="150" Margin="0,0,10,0" Background="#E0E0E0" FontWeight="Bold"/>
            <Button Name="btnContinue" Content="CONTINUE" Height="40" Width="190" Background="#212121" Foreground="#FFFFFF" FontWeight="Bold"/>
        </StackPanel>
    </StackPanel>
</Window>
"""

XAML_GUIDE_UI = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="คู่มือการใช้งาน (User Manual)" Width="750" Height="650"
        Background="#F9F9F9" Topmost="True" WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="60"/>
        </Grid.RowDefinitions>
        
        <Border Grid.Row="0" Background="#212121" Padding="20">
            <TextBlock Text="คู่มือการใช้งาน (Revit Parameter Extractor)" Foreground="White" FontSize="20" FontWeight="Bold" HorizontalAlignment="Center"/>
        </Border>
        
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="25">
            <StackPanel>
                <TextBlock Text="โปรแกรมนี้จะช่วยคุณดึงข้อมูลและพารามิเตอร์ต่างๆ จากไฟล์ Revit Family ออกมาจัดเรียงในตาราง Excel ให้โดยอัตโนมัติ เพื่อให้ง่ายต่อการนำข้อมูลไปเช็คหรือใช้งานต่อ" FontSize="14" TextWrapping="Wrap" Margin="0,0,0,20"/>
                
                <Border Background="#E8F4F8" BorderBrush="#B3D9DF" BorderThickness="1" Padding="15" Margin="0,0,0,20" CornerRadius="5">
                    <StackPanel>
                        <TextBlock Text="ความหมายของสีในตาราง Excel" FontWeight="Bold" FontSize="16" Foreground="#1F4E79" Margin="0,0,0,10"/>
                        <TextBlock Text="เมื่อเปิดไฟล์ Excel ที่ได้ คุณจะพบว่าหัวตารางมีการแยกสีไว้เพื่อบอกประเภทของพารามิเตอร์ ดังนี้:" FontSize="13" TextWrapping="Wrap" Margin="0,0,0,10"/>
                        <StackPanel Margin="10,0,0,0">
                            <TextBlock FontSize="13" Margin="0,2"><Run Foreground="#1F497D" FontWeight="Bold">🟦 สีน้ำเงิน (Dark Blue):</Run><Run> BuiltIn / System Parameter (พารามิเตอร์พื้นฐานที่มากับระบบ)</Run></TextBlock>
                            <TextBlock FontSize="13" Margin="0,2"><Run Foreground="#00B050" FontWeight="Bold">🟩 สีเขียว (Dark Green):</Run><Run> Family Parameter (พารามิเตอร์ที่ถูกสร้างขึ้นเฉพาะภายใน Family นั้นๆ)</Run></TextBlock>
                            <TextBlock FontSize="13" Margin="0,2"><Run Foreground="#E36C09" FontWeight="Bold">🟧 สีส้ม (Dark Orange):</Run><Run> Shared Parameter (พารามิเตอร์ที่ดึงมาจากไฟล์ Shared Parameter)</Run></TextBlock>
                            <TextBlock FontSize="13" Margin="0,2"><Run Foreground="#333333" FontWeight="Bold">⬛ สีเทา/สีดำ:</Run><Run> ข้อมูลทั่วไป (เช่น Path ของไฟล์, ชื่อ Family, ชื่อ Type)</Run></TextBlock>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <TextBlock Text="รูปแบบการดึงข้อมูลทั้ง 4 โหมด" FontWeight="Bold" FontSize="18" Foreground="#333333" Margin="0,0,0,15"/>
                
                <TextBlock Text="1. Active Project: Select by Family" FontWeight="Bold" FontSize="15" Foreground="#0078D7" Margin="0,0,0,5"/>
                <TextBlock TextWrapping="Wrap" Margin="15,0,0,15" FontSize="13">
                    <Run FontWeight="Bold">เหมาะสำหรับ:</Run><Run> ต้องการดึงข้อมูล Family ที่กำลังใช้งานอยู่ในโปรเจกต์ปัจจุบัน</Run><LineBreak/>
                    <Run FontWeight="Bold">การทำงาน:</Run><Run> จะมีหน้าต่างขึ้นมาให้คุณคลิกเลือกรายชื่อ Family ที่ต้องการ</Run><LineBreak/>
                    <Run FontWeight="Bold">ผลลัพธ์ที่ได้:</Run><Run> โปรแกรมจะสร้างไฟล์ Excel โดยแยก 1 Sheet ต่อ 1 Family และดึงข้อมูลของ "ทุก Type" ที่มีใน Family นั้นออกมาให้ครบถ้วน</Run>
                </TextBlock>

                <TextBlock Text="2. Active Project: Select by Category" FontWeight="Bold" FontSize="15" Foreground="#0078D7" Margin="0,0,0,5"/>
                <TextBlock TextWrapping="Wrap" Margin="15,0,0,15" FontSize="13">
                    <Run FontWeight="Bold">เหมาะสำหรับ:</Run><Run> ต้องการดึงข้อมูลแบบเหมารวมตามหมวดหมู่ (Category) ในโปรเจกต์ปัจจุบัน เช่น ดึงข้อมูลประตู (Doors) หรือหน้าต่าง (Windows) ทั้งหมด</Run><LineBreak/>
                    <Run FontWeight="Bold">การทำงาน:</Run><Run> จะมีหน้าต่างขึ้นมาให้คุณคลิกเลือกหมวดหมู่ที่ต้องการ</Run><LineBreak/>
                    <Run FontWeight="Bold">ผลลัพธ์ที่ได้:</Run><Run> โปรแกรมจะสร้างไฟล์ Excel โดยแยก 1 Sheet ต่อ 1 หมวดหมู่ และดึงข้อมูลเฉพาะ "Type ปัจจุบันที่ถูกเลือกไว้" ของแต่ละ Family มาสรุปให้</Run>
                </TextBlock>

                <TextBlock Text="3. External Files: Batch process Folder (.rfa)" FontWeight="Bold" FontSize="15" Foreground="#0078D7" Margin="0,0,0,5"/>
                <TextBlock TextWrapping="Wrap" Margin="15,0,0,15" FontSize="13">
                    <Run FontWeight="Bold">เหมาะสำหรับ:</Run><Run> มีไฟล์ Family (.rfa) เก็บไว้ในโฟลเดอร์ในคอมพิวเตอร์เป็นจำนวนมาก และต้องการดึงข้อมูลทั้งหมดรวดเดียว (ไม่ต้องโหลดเข้าไปในโปรเจกต์)</Run><LineBreak/>
                    <Run FontWeight="Bold">การทำงาน:</Run><Run> ให้คุณเลือก "โฟลเดอร์" ที่เก็บไฟล์ไว้ (โปรแกรมจะค้นหาไฟล์ในซับโฟลเดอร์ย่อยให้ด้วย)</Run><LineBreak/>
                    <Run FontWeight="Bold">ผลลัพธ์ที่ได้:</Run><Run> ข้อมูลของทุกไฟล์จะถูกดึงมาสรุปรวมกันใน Excel หน้าเดียว (ดึงเฉพาะ "Type ปัจจุบัน") พร้อมกับจัดกลุ่มตามโฟลเดอร์เพื่อให้คุณรู้ว่าไฟล์มาจากที่ไหน</Run>
                </TextBlock>

                <TextBlock Text="4. External Files: Select Multiple Files (.rfa)" FontWeight="Bold" FontSize="15" Foreground="#0078D7" Margin="0,0,0,5"/>
                <TextBlock TextWrapping="Wrap" Margin="15,0,0,15" FontSize="13">
                    <Run FontWeight="Bold">เหมาะสำหรับ:</Run><Run> ต้องการเลือกดึงข้อมูลจากไฟล์ Family (.rfa) ในคอมพิวเตอร์แบบเจาะจงทีละไฟล์ หรือหลายๆ ไฟล์</Run><LineBreak/>
                    <Run FontWeight="Bold">การทำงาน:</Run><Run> จะมีหน้าต่างขึ้นมาให้คุณเลือกไฟล์ .rfa (คุณสามารถกด Ctrl หรือ Shift ค้างไว้เพื่อเลือกพร้อมกันหลายๆ ไฟล์ได้)</Run><LineBreak/>
                    <Run FontWeight="Bold">ผลลัพธ์ที่ได้:</Run><Run> ข้อมูลจะถูกนำมาสรุปรวมใน Excel และล่าสุดมีตัวเลือกให้คุณเลือกว่าจะ "รวมข้อมูลหน้าเดียว" หรือ "แยก Sheet ละ 1 Family"</Run>
                </TextBlock>
                
            </StackPanel>
        </ScrollViewer>
        
        <Border Grid.Row="2" Background="#EEEEEE" BorderBrush="#DDDDDD" BorderThickness="0,1,0,0">
            <Button Name="btnClose" Content="CLOSE" Width="120" Height="35" Background="#212121" Foreground="#FFFFFF" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center" Cursor="Hand"/>
        </Border>
    </Grid>
</Window>
"""

class MainSelectionUI(object):
    def __init__(self):
        self.selected_mode = 1
        self.separate_sheets = False
        self.window = XamlReader.Load(XmlReader.Create(StringReader(XAML_MAIN_UI)))
        
        self.rbActiveFam = self.window.FindName("rbActiveFam")
        self.rbActiveCat = self.window.FindName("rbActiveCat")
        self.rbBatchFolder = self.window.FindName("rbBatchFolder")
        self.rbSingleFile = self.window.FindName("rbSingleFile")
        self.chkSeparateSheets = self.window.FindName("chkSeparateSheets")
        self.btnContinue = self.window.FindName("btnContinue")
        self.btnGuide = self.window.FindName("btnGuide")
        
        self.btnContinue.Click += self.on_continue
        self.btnGuide.Click += self.on_guide
        
        self.rbActiveFam.Checked += self.on_radio_changed
        self.rbActiveCat.Checked += self.on_radio_changed
        self.rbBatchFolder.Checked += self.on_radio_changed
        self.rbSingleFile.Checked += self.on_radio_changed
        
    def on_radio_changed(self, sender, e):
        if self.rbSingleFile.IsChecked:
            self.chkSeparateSheets.IsEnabled = True
        else:
            self.chkSeparateSheets.IsEnabled = False
            self.chkSeparateSheets.IsChecked = False
        
    def on_continue(self, sender, e):
        if self.rbActiveFam.IsChecked:
            self.selected_mode = 1
        elif self.rbActiveCat.IsChecked:
            self.selected_mode = 2
        elif self.rbBatchFolder.IsChecked:
            self.selected_mode = 3
        elif self.rbSingleFile.IsChecked:
            self.selected_mode = 4
            self.separate_sheets = bool(self.chkSeparateSheets.IsChecked)
            
        self.window.DialogResult = True
        self.window.Close()
        
    def on_guide(self, sender, e):
        guide_window = XamlReader.Load(XmlReader.Create(StringReader(XAML_GUIDE_UI)))
        btnClose = guide_window.FindName("btnClose")
        
        def close_guide(s, ev):
            guide_window.Close()
            
        btnClose.Click += close_guide
        guide_window.ShowDialog()



# ==========================================
# Shared Utilities
# ==========================================
def rgb_to_excel(r, g, b):
    return r + (g * 256) + (b * 65536)

# ==========================================
# UI Controllers for Active Project Modes
# ==========================================
XAML_LIST_UI = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{Binding WindowTitle}" Width="450" Height="550"
        Background="#F5F5F5" Topmost="True" WindowStartupLocation="CenterScreen">
    <StackPanel Margin="15">
        <Label Name="lblHeader" Content="1. เลือกรายการที่ต้องการดึงข้อมูล:" FontWeight="Bold" Background="#E0E0E0" Padding="5"/>
        <TextBlock Name="tbSearch" Text="พิมพ์ค้นหา:" Margin="0,8,0,2"/>
        <TextBox Name="txtSearch" Height="25" Margin="0,0,0,5" />
        <StackPanel Orientation="Horizontal" Margin="0,0,0,5">
            <Button Name="btnSelectAll" Content="Select All" Width="80" Height="22" Margin="0,0,5,0"/>
            <Button Name="btnClearAll"  Content="Clear All"  Width="80" Height="22"/>
        </StackPanel>
        <ListBox Name="lbItems" Height="300" Margin="0,0,0,10"/>
        
        <Button Name="btnRun" Content="EXTRACT PARAMETERS TO EXCEL" Height="45" Background="#212121" Foreground="#FFFFFF" FontWeight="Bold"/>
    </StackPanel>
</Window>
"""

class ListSelectionUI(object):
    def __init__(self, item_names, title="Selection", header_text="Select Items:"):
        self.checkbox_map = {}
        self.sorted_names = sorted(item_names)
        self.selected_items = []
        
        self.window = XamlReader.Load(XmlReader.Create(StringReader(XAML_LIST_UI)))
        self.window.Title = title
        
        self.window.FindName("lblHeader").Content = header_text
        
        self.txtSearch = self.window.FindName("txtSearch")
        self.lbItems = self.window.FindName("lbItems")
        self.btnSelectAll = self.window.FindName("btnSelectAll")
        self.btnClearAll = self.window.FindName("btnClearAll")
        self.btnRun = self.window.FindName("btnRun")
        
        self.txtSearch.TextChanged += self.on_search_changed
        self.btnSelectAll.Click += self.on_select_all
        self.btnClearAll.Click += self.on_clear_all
        self.btnRun.Click += self.on_submit
        
        self.refresh_list("")

    def refresh_list(self, search_text):
        self.lbItems.Items.Clear()
        search_text = search_text.lower()
        for name in self.sorted_names:
            if search_text in name.lower():
                prev_checked = self.checkbox_map[name].IsChecked if name in self.checkbox_map else False
                cb = CheckBox()
                cb.Content = name
                cb.IsChecked = prev_checked
                self.checkbox_map[name] = cb
                self.lbItems.Items.Add(cb)

    def on_search_changed(self, sender, e):
        self.refresh_list(self.txtSearch.Text)

    def on_select_all(self, sender, e):
        for cb in self.checkbox_map.values():
            cb.IsChecked = True

    def on_clear_all(self, sender, e):
        for cb in self.checkbox_map.values():
            cb.IsChecked = False

    def on_submit(self, sender, e):
        self.selected_items = [name for name, cb in self.checkbox_map.items() if cb.IsChecked]
        self.window.DialogResult = True
        self.window.Close()


# ==========================================
# MODE 1: Active Project - Family Extractor
# ==========================================
def process_active_project_families(doc, app):
    families = FilteredElementCollector(doc).OfClass(Family).ToElements()
    editable_families = {fam.Name: fam for fam in families if fam.IsEditable}
    
    if not editable_families:
        return ["No editable families found in the active project.", log]
        
    ui = ListSelectionUI(editable_families.keys(), "Active Project Family Parameter Extractor", "1. เลือก Family ที่ต้องการดึงข้อมูล:")
    if not ui.window.ShowDialog():
        return ["User canceled family selection.", log]
        
    selected_fam_names = ui.selected_items
    if not selected_fam_names:
        return ["No families selected.", log]
        
    family_type_data_map = {}
    
    for fam_name in selected_fam_names:
        fam = editable_families[fam_name]
        all_type_data = []
        all_param_names = []
        
        try:
            fam_doc = doc.EditFamily(fam)
            if fam_doc and fam_doc.IsFamilyDocument:
                fam_manager = fam_doc.FamilyManager
                for fam_type in fam_manager.Types:
                    type_name = fam_type.Name
                    type_data = {"Family Name": fam_name, "Type Name": type_name}
                    
                    for param in fam_manager.Parameters:
                        base_param_name = param.Definition.Name
                        if param.IsShared: p_type = "Shared"
                        elif param.Id.IntegerValue < 0: p_type = "BuiltIn"
                        else: p_type = "Family"
                        param_name_formatted = "{} [{}]".format(base_param_name, p_type)
                        
                        val = ""
                        if fam_type.HasValue(param):
                            storage_type = param.StorageType
                            try: val_string = fam_type.AsValueString(param)
                            except: val_string = None
                            
                            if val_string: val = val_string
                            else:
                                if storage_type == StorageType.Double: val = fam_type.AsDouble(param)
                                elif storage_type == StorageType.Integer: val = fam_type.AsInteger(param)
                                elif storage_type == StorageType.String: val = fam_type.AsString(param)
                                elif storage_type == StorageType.ElementId: val = fam_type.AsElementId(param).IntegerValue
                        
                        type_data[param_name_formatted] = str(val) if val is not None else ""
                        if param_name_formatted not in all_param_names:
                            all_param_names.append(param_name_formatted)
                    all_type_data.append(type_data)
                fam_doc.Close(False)
                
                if all_type_data:
                    all_type_data.sort(key=lambda x: x.get("Type Name", ""))
                    family_type_data_map[fam_name] = {"types": all_type_data, "params": all_param_names}
        except Exception as e:
            log.append("Error opening family: {} | Msg: {}".format(fam_name, str(e)))
                
    if not family_type_data_map:
        return ["No parameters extracted.", log]

    try:
        excel = Excel.ApplicationClass()
        excel.Visible = True 
        excel.ScreenUpdating = False
        excel.Interactive = False
        workbook = excel.Workbooks.Add()
        
        first_sheet = True
        for fam_name, fam_info in family_type_data_map.items():
            if first_sheet:
                worksheet = workbook.Worksheets[1]
                first_sheet = False
            else:
                worksheet = workbook.Worksheets.Add(After=workbook.Worksheets[workbook.Worksheets.Count])
                
            safe_sheet_name = re.sub(r'[\\\\/*?:\\[\\]]', '', fam_name)
            if len(safe_sheet_name) > 31: safe_sheet_name = safe_sheet_name[:28] + "..."
            worksheet.Name = safe_sheet_name if safe_sheet_name else "Family"
            
            worksheet.Tab.Color = rgb_to_excel(random.randint(150, 255), random.randint(150, 255), random.randint(150, 255))
            
            all_type_data = fam_info["types"]
            all_param_names = fam_info["params"]
            all_param_names.sort(key=lambda x: (x.split(' [')[-1], x))
            
            final_param_names = []
            current_type = None
            for name in all_param_names:
                p_type = name.split(' [')[-1]
                if current_type is not None and p_type != current_type:
                    final_param_names.append("")
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
                
            rows_count = len(all_type_data)
            cols_count = len(headers)
            if rows_count > 0:
                data_arr = System.Array.CreateInstance(System.Object, rows_count, cols_count)
                for row_idx in range(rows_count):
                    type_data = all_type_data[row_idx]
                    for col_idx in range(cols_count):
                        header = headers[col_idx]
                        if col_idx == 0: data_arr[row_idx, col_idx] = type_data.get("Family Name", "")
                        elif col_idx == 1: data_arr[row_idx, col_idx] = type_data.get("Type Name", "")
                        elif header != "" and header in type_data: data_arr[row_idx, col_idx] = str(type_data[header])
                            
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
                        if start_col <= col_idx_excel - 1:
                            worksheet.Range(worksheet.Cells(1, start_col), worksheet.Cells(1, col_idx_excel - 1)).EntireColumn.Group()
                    current_type = p_type
                    start_col = col_idx_excel
            if current_type is not None and current_type != "Separator":
                if start_col <= len(final_param_names) + 2:
                    worksheet.Range(worksheet.Cells(1, start_col), worksheet.Cells(1, len(final_param_names) + 2)).EntireColumn.Group()
        
        excel.ScreenUpdating = True
        excel.Interactive = True
        try:
            excel.WindowState = Excel.XlWindowState.xlMinimized
            excel.WindowState = Excel.XlWindowState.xlMaximized
        except: pass
        
        result_msg = "Success! Processed {} families into {} sheets.".format(len(family_type_data_map), len(family_type_data_map))
        if log: result_msg += "\n(With {} errors)".format(len(log))
        MessageBox.Show(result_msg, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        return [result_msg, log]
    except Exception as e:
        try:
            if 'excel' in locals():
                excel.ScreenUpdating = True
                excel.Interactive = True
        except: pass
        MessageBox.Show(str(e), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        return [str(e), log]


# ==========================================
# MODE 2: Active Project - Category Extractor
# ==========================================
def process_active_project_categories(doc, app):
    families = FilteredElementCollector(doc).OfClass(Family).ToElements()
    category_family_map = {}
    for fam in families:
        if fam.IsEditable:
            cat = fam.FamilyCategory
            if cat is not None:
                cat_name = cat.Name
                if cat_name not in category_family_map:
                    category_family_map[cat_name] = []
                category_family_map[cat_name].append(fam)
                
    if not category_family_map:
        return ["No editable families found in the active project.", log]
        
    ui = ListSelectionUI(category_family_map.keys(), "Active Project Parameter Extractor", "1. เลือก Category ของ Family ที่ต้องการดึงข้อมูล:")
    if not ui.window.ShowDialog():
        return ["User canceled category selection.", log]
        
    selected_cats = ui.selected_items
    if not selected_cats:
        return ["No categories selected.", log]
        
    all_family_data = []
    
    for cat_name in selected_cats:
        fams_in_cat = category_family_map[cat_name]
        fams_in_cat.sort(key=lambda x: x.Name)
        for fam in fams_in_cat:
            try:
                fam_doc = doc.EditFamily(fam)
                if fam_doc and fam_doc.IsFamilyDocument:
                    fam_name = fam.Name
                    fam_data = {"Category": cat_name, "Family Name": fam_name}
                    fam_manager = fam_doc.FamilyManager
                    current_type = fam_manager.CurrentType
                    
                    for param in fam_manager.Parameters:
                        base_param_name = param.Definition.Name
                        if param.IsShared: p_type = "Shared"
                        elif param.Id.IntegerValue < 0: p_type = "BuiltIn"
                        else: p_type = "Family"
                        param_name = "{} [{}]".format(base_param_name, p_type)
                        
                        val = ""
                        if current_type is not None and current_type.HasValue(param):
                            storage_type = param.StorageType
                            if storage_type == StorageType.Double: val = current_type.AsDouble(param)
                            elif storage_type == StorageType.Integer: val = current_type.AsInteger(param)
                            elif storage_type == StorageType.String: val = current_type.AsString(param)
                            elif storage_type == StorageType.ElementId: val = current_type.AsElementId(param).IntegerValue
                        fam_data[param_name] = str(val) if val is not None else ""
                    all_family_data.append(fam_data)
                    fam_doc.Close(False)
            except Exception as e:
                log.append("Error opening family: {} | Msg: {}".format(fam.Name, str(e)))
                
    if not all_family_data:
        return ["No parameters extracted.", log]

    try:
        excel = Excel.ApplicationClass()
        excel.Visible = True 
        excel.ScreenUpdating = False
        excel.Interactive = False
        workbook = excel.Workbooks.Add()
        
        data_by_cat = {}
        for d in all_family_data:
            c = d["Category"]
            if c not in data_by_cat: data_by_cat[c] = []
            data_by_cat[c].append(d)
            
        first_sheet = True
        for cat_name, cat_data in data_by_cat.items():
            if first_sheet:
                worksheet = workbook.Worksheets[1]
                first_sheet = False
            else:
                worksheet = workbook.Worksheets.Add(After=workbook.Worksheets[workbook.Worksheets.Count])
                
            safe_sheet_name = re.sub(r'[\\\\/*?:\\[\\]]', '', cat_name)
            if len(safe_sheet_name) > 31: safe_sheet_name = safe_sheet_name[:28] + "..."
            worksheet.Name = safe_sheet_name if safe_sheet_name else "Category"
            worksheet.Tab.Color = rgb_to_excel(random.randint(150, 255), random.randint(150, 255), random.randint(150, 255))
            
            cat_param_names = []
            for d in cat_data:
                for k in d.keys():
                    if k not in ["Category", "Family Name", "IsSeparator"] and k not in cat_param_names:
                        cat_param_names.append(k)
            cat_param_names.sort(key=lambda x: (x.split(' [')[-1], x))
            
            final_param_names = []
            current_type = None
            for name in cat_param_names:
                p_type = name.split(' [')[-1]
                if current_type is not None and p_type != current_type: final_param_names.append("")
                final_param_names.append(name)
                current_type = p_type
                
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
                if "[BuiltIn]" in header: cell.Interior.Color = rgb_to_excel(51, 51, 153)
                elif "[Family]" in header: cell.Interior.Color = rgb_to_excel(34, 139, 34)
                elif "[Shared]" in header: cell.Interior.Color = rgb_to_excel(255, 140, 0)
                else: cell.Interior.Color = rgb_to_excel(70, 70, 70)
                cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                
            rows_count = len(cat_data)
            cols_count = len(headers)
            if rows_count > 0:
                data_arr = System.Array.CreateInstance(System.Object, rows_count, cols_count)
                for r in range(rows_count):
                    fam_data = cat_data[r]
                    for c in range(cols_count):
                        header = headers[c]
                        if c == 0: data_arr[r, c] = fam_data.get("Family Name", "")
                        elif header != "" and header in fam_data: data_arr[r, c] = str(fam_data[header])
                            
                worksheet.Range(worksheet.Cells(2, 1), worksheet.Cells(rows_count + 1, cols_count)).Value2 = data_arr
                if cols_count >= 2:
                    param_range = worksheet.Range(worksheet.Cells(2, 2), worksheet.Cells(rows_count + 1, cols_count))
                    param_range.Interior.Color = rgb_to_excel(220, 220, 220) 
                    try: param_range.SpecialCells(2).Interior.Color = rgb_to_excel(255, 255, 255)
                    except: pass
            
            worksheet.Columns.AutoFit()
            used_range = worksheet.UsedRange
            used_range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            used_range.Borders.Weight = Excel.XlBorderWeight.xlThin
            
            current_type = None
            start_col = 2
            for i, param_name in enumerate(final_param_names):
                col_idx_excel = i + 2
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
                if start_col <= len(final_param_names) + 1: worksheet.Range(worksheet.Cells(1, start_col), worksheet.Cells(1, len(final_param_names) + 1)).EntireColumn.Group()
        
        excel.ScreenUpdating = True
        excel.Interactive = True
        result_msg = "Success! Processed {} families into {} sheets.".format(len(all_family_data), len(data_by_cat))
        if log: result_msg += " (With {} errors)".format(len(log))
        MessageBox.Show(result_msg, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        return [result_msg, log]
    except Exception as e:
        try:
            if 'excel' in locals():
                excel.ScreenUpdating = True
                excel.Interactive = True
        except: pass
        MessageBox.Show(str(e), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        return [str(e), log]


# ==========================================
# MODE 3: Batch Process Folder (Families)
# ==========================================
def process_batch_folder(doc, app):
    dialog = OpenFileDialog()
    dialog.Title = "Select a folder containing Revit Family (.rfa) files (Click Open)"
    dialog.ValidateNames = False
    dialog.CheckFileExists = False
    dialog.CheckPathExists = True
    dialog.FileName = "Select_This_Folder"
    dialog.Filter = "Folders Only|*.none"
    
    if dialog.ShowDialog() == DialogResult.OK:
        folder_path = os.path.dirname(dialog.FileName)
        rfa_files = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(".rfa"):
                    rfa_files.append(os.path.join(root, file))
        
        if not rfa_files:
            msg = "No .rfa files found in: " + folder_path
            MessageBox.Show(msg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return msg
            
        all_family_data = []
        all_param_names = []
        
        for rfa_path in rfa_files:
            try:
                model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(rfa_path)
                open_options = OpenOptions()
                fam_doc = app.OpenDocumentFile(model_path, open_options)
                
                if fam_doc.IsFamilyDocument:
                    fam_name = fam_doc.Title
                    display_path = os.path.dirname(rfa_path)
                    fam_data = {"File Path": display_path, "Family Name": fam_name}
                    
                    fam_manager = fam_doc.FamilyManager
                    current_type = fam_manager.CurrentType
                    
                    for param in fam_manager.Parameters:
                        base_param_name = param.Definition.Name
                        if param.IsShared: p_type = "Shared"
                        elif param.Id.IntegerValue < 0: p_type = "BuiltIn"
                        else: p_type = "Family"
                        param_name = "{} [{}]".format(base_param_name, p_type)
                        
                        val = ""
                        if current_type is not None and current_type.HasValue(param):
                            storage_type = param.StorageType
                            try: val_string = current_type.AsValueString(param)
                            except: val_string = None
                            
                            if val_string: val = val_string
                            else:
                                if storage_type == StorageType.Double: val = current_type.AsDouble(param)
                                elif storage_type == StorageType.Integer: val = current_type.AsInteger(param)
                                elif storage_type == StorageType.String: val = current_type.AsString(param)
                                elif storage_type == StorageType.ElementId: val = current_type.AsElementId(param).IntegerValue
                        fam_data[param_name] = str(val) if val is not None else ""
                        if param_name not in all_param_names: all_param_names.append(param_name)
                    all_family_data.append(fam_data)
                fam_doc.Close(False)
            except Exception as e:
                log.append("Error opening file: {} | Msg: {}".format(rfa_path, str(e)))

        try:
            excel = Excel.ApplicationClass()
            excel.Visible = True 
            excel.ScreenUpdating = False
            excel.Interactive = False
            workbook = excel.Workbooks.Add()
            worksheet = workbook.Worksheets[1]
            worksheet.Name = "Family Parameters"
            
            all_param_names.sort(key=lambda x: (x.split(' [')[-1], x))
            final_param_names = []
            current_type = None
            for name in all_param_names:
                p_type = name.split(' [')[-1]
                if current_type is not None and p_type != current_type: final_param_names.append("")
                final_param_names.append(name)
                current_type = p_type
            
            headers = ["File Path", "Family Name"] + final_param_names
            
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
            
            all_family_data.sort(key=lambda x: x.get("File Path", ""))
            
            new_all_family_data = []
            for i in range(len(all_family_data)):
                if i > 0:
                    prev_path = all_family_data[i-1].get("File Path", "")
                    curr_path = all_family_data[i].get("File Path", "")
                    if prev_path != curr_path:
                        parts1 = prev_path.split('\\')
                        parts2 = curr_path.split('\\')
                        common = []
                        for a, b in zip(parts1, parts2):
                            if a == b: common.append(a)
                            else: break
                        common_path = "\\".join(common)
                        new_all_family_data.append({"File Path": common_path, "IsSeparator": True})
                new_all_family_data.append(all_family_data[i])
            all_family_data = new_all_family_data
            
            rows_count = len(all_family_data)
            cols_count = len(headers)
            if rows_count > 0:
                data_arr = System.Array.CreateInstance(System.Object, rows_count, cols_count)
                for r in range(rows_count):
                    fam_data = all_family_data[r]
                    if fam_data.get("IsSeparator"): continue
                    for c in range(cols_count):
                        header = headers[c]
                        if c == 0: data_arr[r, c] = fam_data.get("File Path", "")
                        elif c == 1: data_arr[r, c] = fam_data.get("Family Name", "")
                        elif header != "" and header in fam_data: data_arr[r, c] = str(fam_data[header])
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
            
            if len(all_family_data) > 0:
                max_depth = min(7, max([len(d.get("File Path", "").split('\\')) for d in all_family_data]))
                for lvl in range(max_depth):
                    start_row = None
                    current_val = None
                    for row_offset in range(len(all_family_data)):
                        rfa_path = all_family_data[row_offset].get("File Path", "")
                        if not rfa_path: continue
                        parts = rfa_path.split('\\')
                        val = "\\".join(parts[:lvl+1]) if lvl < len(parts) else None
                        if val != current_val:
                            if current_val is not None and start_row is not None:
                                if start_row <= row_offset + 1: worksheet.Range(worksheet.Cells(start_row, 1), worksheet.Cells(row_offset + 1, 1)).EntireRow.Group()
                            current_val = val
                            start_row = row_offset + 2 if val is not None else None
                    if current_val is not None and start_row is not None:
                        if start_row <= len(all_family_data) + 1: worksheet.Range(worksheet.Cells(start_row, 1), worksheet.Cells(len(all_family_data) + 1, 1)).EntireRow.Group()
            
            excel.ScreenUpdating = True
            excel.Interactive = True
            try:
                excel.WindowState = Excel.XlWindowState.xlMinimized
                excel.WindowState = Excel.XlWindowState.xlMaximized
            except: pass
            
            result_msg = "Success! Processed {} families and wrote to Excel.".format(len(all_family_data))
            if log: result_msg += "\n(With {} errors)".format(len(log))
            MessageBox.Show(result_msg, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            return [result_msg, log]
        except Exception as e:
            try:
                if 'excel' in locals():
                    excel.ScreenUpdating = True
                    excel.Interactive = True
            except: pass
            MessageBox.Show(str(e), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return str(e)
    else:
        return "User canceled folder selection."


# ==========================================
# MODE 4: Process Single File
# ==========================================
def process_single_file(doc, app, separate_sheets):
    dialog = OpenFileDialog()
    dialog.Title = "Select a Revit Family (.rfa) file"
    dialog.Filter = "Revit Family Files (*.rfa)|*.rfa"
    dialog.Multiselect = True
    
    if dialog.ShowDialog() == DialogResult.OK:
        all_type_data = []
        all_param_names = []
        
        for rfa_path in dialog.FileNames:
            try:
                model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(rfa_path)
                open_options = OpenOptions()
                fam_doc = app.OpenDocumentFile(model_path, open_options)
                
                if fam_doc.IsFamilyDocument:
                    fam_name = fam_doc.Title
                    fam_manager = fam_doc.FamilyManager
                    
                    for fam_type in fam_manager.Types:
                        type_name = fam_type.Name
                        type_data = {"Family Name": fam_name, "Type Name": type_name}
                        
                        for param in fam_manager.Parameters:
                            base_param_name = param.Definition.Name
                            if param.IsShared: p_type = "Shared"
                            elif param.Id.IntegerValue < 0: p_type = "BuiltIn"
                            else: p_type = "Family"
                            param_name = "{} [{}]".format(base_param_name, p_type)
                            
                            val = ""
                            if fam_type.HasValue(param):
                                storage_type = param.StorageType
                                try: val_string = fam_type.AsValueString(param)
                                except: val_string = None
                                if val_string: val = val_string
                                else:
                                    if storage_type == StorageType.Double: val = fam_type.AsDouble(param)
                                    elif storage_type == StorageType.Integer: val = fam_type.AsInteger(param)
                                    elif storage_type == StorageType.String: val = fam_type.AsString(param)
                                    elif storage_type == StorageType.ElementId: val = fam_type.AsElementId(param).IntegerValue
                            type_data[param_name] = str(val) if val is not None else ""
                            if param_name not in all_param_names: all_param_names.append(param_name)
                        all_type_data.append(type_data)
                fam_doc.Close(False)
            except Exception as e:
                log.append("Error processing file: {} | Msg: {}".format(rfa_path, str(e)))
        
        if all_type_data:
            all_type_data.sort(key=lambda x: (x.get("Family Name", ""), x.get("Type Name", "")))
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
                    
                    # Gather params for THIS group
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
                
                excel.ScreenUpdating = True
                excel.Interactive = True
                try:
                    excel.WindowState = Excel.XlWindowState.xlMinimized
                    excel.WindowState = Excel.XlWindowState.xlMaximized
                except: pass
                
                result_msg = "Success! Processed {} types and wrote to Excel.".format(len(all_type_data))
                if log: result_msg += "\n(With {} errors)".format(len(log))
                MessageBox.Show(result_msg, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                return [result_msg, log]
            except Exception as e:
                try:
                    if 'excel' in locals():
                        excel.ScreenUpdating = True
                        excel.Interactive = True
                except: pass
                MessageBox.Show(str(e), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                return str(e)
        else:
            msg = "No types found or unable to read."
            MessageBox.Show(msg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return msg
    else:
        return "User canceled file selection."


# ==========================================
# Main Execution
# ==========================================
def main_process():
    doc = DocumentManager.Instance.CurrentDBDocument
    app = DocumentManager.Instance.CurrentUIApplication.Application
    
    ui = MainSelectionUI()
    if ui.window.ShowDialog():
        mode = ui.selected_mode
        if mode == 1:
            return process_active_project_families(doc, app)
        elif mode == 2:
            return process_active_project_categories(doc, app)
        elif mode == 3:
            return process_batch_folder(doc, app)
        elif mode == 4:
            return process_single_file(doc, app, ui.separate_sheets)
    else:
        return "User canceled mode selection."

if run_it:
    OUT = main_process()
else:
    OUT = "Waiting... (Set IN[0] to True to run the script)"
