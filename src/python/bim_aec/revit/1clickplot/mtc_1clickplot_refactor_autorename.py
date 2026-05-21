# -*- coding: utf-8 -*-
"""
Script: MTC Change Object & Line Style Color 1clickplot (Refactored + Auto-Rename)
Architecture: Single-file with Separation of Concerns (SoC) + TransactionGroup Rollback
"""
import clr
import System
import re
import os
import time
import shutil

# --- Load Revit API ---
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    Color as RevitColor, Transaction, TransactionGroup, FilteredElementCollector,
    ViewSheetSet, PrintSetting, View, ViewSheet, ElementId, ViewSet
)

# --- Load Windows / WPF API ---
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("System.Xml")
clr.AddReference("System.Printing")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Markup import XamlReader
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows.Controls import CheckBox
from System.Drawing.Printing import PrinterSettings
from System.Printing import LocalPrintServer
from System.Windows.Forms import FolderBrowserDialog, DialogResult

# ==========================================
# 1. Config & Utilities
# ==========================================
class Config:
    IS_ACTIVE = IN[0] if isinstance(IN[0], bool) else True
    DEFAULT_EXCLUSIONS = ["Revision Clouds", "Revision Cloud Tags"]
    BLACK_COLOR = RevitColor(0, 0, 0)
    TARGET_PRINTER = "PDF24"
    TEMP_PDF_FOLDER = r"C:\Revit_PDF_Temp"
    TARGET_PARAM_NAME = "Sheet Number" #"FileName" ของเดิม

    XAML_UI = """
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="Itti Drawing Plot Manager (Pro)" Width="560" Height="920"
            Background="#F5F5F5" Topmost="True" WindowStartupLocation="CenterScreen">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
        <StackPanel Margin="20">
            <Label Content="1. เลือกรายการที่ต้องการ Exclude (ไม่เปลี่ยนเป็นสีดำ):" FontWeight="Bold" Background="#E0E0E0" Padding="5"/>
            <TextBox Name="txtSearch" Height="25" Margin="0,5,0,5"/>
            <ListBox Name="lbCategories" Height="180" Margin="0,0,0,10"/>

            <Label Content="2. ตั้งค่า Printer และ Print Setting:" FontWeight="Bold" Background="#E0E0E0" Padding="5" Margin="0,5,0,0"/>
            <TextBlock Text="เลือก Printer:" Margin="5,8,0,2"/>
            <ComboBox Name="cbPrinter" Margin="0,0,0,8"/>
            <TextBlock Text="เลือก Print Setting:" Margin="5,0,0,2"/>
            <ComboBox Name="cbPrintSetting" Margin="0,0,0,10"/>

            <Label Content="3. เลือกโหมดการระบุ View/Sheet ที่จะ Plot:" FontWeight="Bold" Background="#E0E0E0" Padding="5" Margin="0,5,0,0"/>
            <StackPanel Orientation="Horizontal" Margin="5,8,0,5">
                <RadioButton Name="rbModeSet" Content="ใช้ View/Sheet Set (แบบเดิม)" IsChecked="True" GroupName="PlotMode" Margin="0,0,20,0"/>
                <RadioButton Name="rbModeCustom" Content="เลือก View/Sheet เอง (แบบใหม่)" GroupName="PlotMode"/>
            </StackPanel>

            <StackPanel Name="panelSetMode" Margin="0,0,0,10">
                <TextBlock Text="เลือก View / Sheet Set:" Margin="5,0,0,2"/>
                <ComboBox Name="cbViewSet" Margin="0,0,0,5"/>
            </StackPanel>

            <StackPanel Name="panelCustomMode" Visibility="Collapsed" Margin="0,0,0,10">
                <TextBlock Text="ค้นหา Sheet/View:" Margin="5,0,0,2"/>
                <TextBox Name="txtSheetSearch" Height="25" Margin="0,2,0,5"/>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,5">
                    <Button Name="btnSelectAll" Content="Select All" Width="90" Height="22" Margin="0,0,5,0"/>
                    <Button Name="btnClearAll"  Content="Clear All"  Width="90" Height="22"/>
                </StackPanel>
                <ListBox Name="lbSheets" Height="200" Margin="0,0,0,5"/>
            </StackPanel>

            <Label Content="4. ตั้งค่าไฟล์ Output:" FontWeight="Bold" Background="#E0E0E0" Padding="5" Margin="0,5,0,0"/>
            <TextBlock Text="รูปแบบการเซฟไฟล์:" Margin="5,8,0,2"/>
            <RadioButton Name="rbCombine"  Content="Combine — รวมเป็นไฟล์เดียว" IsChecked="True" Margin="5,0,0,3"/>
            <RadioButton Name="rbSeparate" Content="Separate — แยกแต่ละ Sheet/View" Margin="5,0,0,10"/>

            <TextBlock Text="โฟลเดอร์สำหรับเซฟไฟล์ (Output Path):" Margin="5,0,0,2"/>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                <TextBox Name="txtPath" Width="370" Height="25" Margin="5,0,5,0" Text="C:\Temp"/>
                <Button Name="btnBrowse" Content="Browse..." Width="80" Height="25"/>
            </StackPanel>

            <CheckBox Name="chkAutoRename" Content="เปิดใช้ระบบ Auto-Rename (ดึงชื่อไฟล์จาก Parameter 'FileName')" IsChecked="True" Margin="0,0,0,5" Foreground="#0055FF" FontWeight="Bold"/>
            <CheckBox Name="chkDoPrint" Content="สั่ง Plot ทันที (พร้อมระบบเช็คคิวอัจฉริยะ)" IsChecked="True" Margin="0,0,0,10"/>
            <Button Name="btnRun" Content="RUN PROCESS" Height="50" Background="#212121" Foreground="#FFFFFF" FontWeight="Bold"/>
        </StackPanel>
        </ScrollViewer>
    </Window>
    """

def natural_sort_key(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split('([0-9]+)', text)]

def get_elem(eid):
    try: return doc.GetElement(eid)
    except Exception: return None

# ==========================================
# 2. Managers (Revit & System Print)
# ==========================================
class PrintSystemManager:
    @staticmethod
    def get_installed_printers():
        return sorted([p for p in PrinterSettings.InstalledPrinters], key=natural_sort_key)

    @staticmethod
    def wait_for_spooler(printer_name, timeout_limit=120):
        try:
            print_server = LocalPrintServer()
            print_queue  = print_server.GetPrintQueue(printer_name)
            time.sleep(1)
            start_time = time.time()
            
            while True:
                print_queue.Refresh()
                if print_queue.NumberOfJobs == 0:
                    return u"✅ พิมพ์เสร็จสมบูรณ์! ระบบกำลังคืนค่าสี..."
                if (time.time() - start_time) > timeout_limit:
                    return u"⚠️ Timeout 2 นาที ข้ามการรอและคืนค่าสี..."
                time.sleep(1)
        except Exception:
            time.sleep(5) # Fallback หากเข้าถึง Spooler ไม่ได้
            return u"✅ ส่งคำสั่ง Plot สำเร็จ (Fallback 5s)"
            
    @staticmethod
    def watch_and_rename_files(rename_map, out_dir):
        try:
            processed = 0
            expected = len(rename_map)
            timeout = 600
            start_t = time.time()
            
            while processed < expected:
                if time.time() - start_t > timeout:
                    return u"\n⚠️ หมดเวลารอไฟล์จาก PDF24 (Timeout)"
                    
                if os.path.exists(Config.TEMP_PDF_FOLDER):
                    for filename in os.listdir(Config.TEMP_PDF_FOLDER):
                        if filename.endswith(".pdf"):
                            match = re.search(r" - Sheet - (.*?) - ", filename)
                            if match:
                                sheet_num_extracted = match.group(1).strip()
                                if sheet_num_extracted in rename_map:
                                    old_path = os.path.join(Config.TEMP_PDF_FOLDER, filename)
                                    desired_name = rename_map[sheet_num_extracted]
                                    new_path = os.path.join(out_dir, desired_name)
                                    
                                    if os.path.exists(new_path):
                                        try: os.remove(new_path)
                                        except: pass
                                    
                                    try:
                                        shutil.move(old_path, new_path)
                                        processed += 1
                                        del rename_map[sheet_num_extracted]
                                    except:
                                        pass
                time.sleep(1)
            
            if processed > 0:
                return u"\n✅ เปลี่ยนชื่อและย้ายไฟล์สำเร็จ {} แผ่น".format(processed)
            return u""
        except Exception as e:
            return u"\n❌ Rename Error: " + str(e)

class RevitDataManager:
    def __init__(self, doc):
        self.doc = doc
        self.cat_obj_map = {}

    def get_print_settings(self):
        ps_elems = FilteredElementCollector(self.doc).OfClass(PrintSetting).ToElements()
        ps_id_map = {ps.Name: ps.Id for ps in ps_elems if ps.Name}
        return sorted(ps_id_map.keys(), key=natural_sort_key), ps_id_map

    def get_view_sheet_sets(self):
        vs_elems = FilteredElementCollector(self.doc).OfClass(ViewSheetSet).ToElements()
        vs_id_map = {vs.Name: vs.Id for vs in vs_elems if vs.Name and not vs.Name.startswith("<")}
        return sorted(vs_id_map.keys(), key=natural_sort_key), vs_id_map

    def get_printable_items(self):
        all_sheets = FilteredElementCollector(self.doc).OfClass(ViewSheet).ToElements()
        sheets_data = sorted(
            [{"name": u"[Sheet] {}: {}".format(s.SheetNumber, s.Name), "id": s.Id} for s in all_sheets],
            key=lambda x: natural_sort_key(x["name"])
        )
        all_views = FilteredElementCollector(self.doc).OfClass(View).ToElements()
        views_data = sorted(
            [{"name": u"[View] {}".format(v.Name), "id": v.Id} for v in all_views 
             if not v.IsTemplate and v.CanBePrinted and not hasattr(v, 'SheetNumber')],
            key=lambda x: natural_sort_key(x["name"])
        )
        return sheets_data + views_data

    def get_categories(self):
        cat_names = {}
        for cat in self.doc.Settings.Categories:
            try:
                if ".dwg" in cat.Name.lower(): continue
                cat_names[cat.Name] = cat.Name
                self.cat_obj_map[cat.Name] = cat
                for subcat in cat.SubCategories:
                    if ".dwg" not in subcat.Name.lower():
                        key = "{}  >  {}".format(cat.Name, subcat.Name)
                        cat_names[key] = key
                        self.cat_obj_map[key] = subcat
            except AttributeError: pass
        return cat_names

    def change_colors_to_black(self, excluded_names):
        # ใช้วิธีลุยเปลี่ยนเป็นสีดำทั้งหมดโดยไม่ต้องเก็บค่าเดิม (เพราะเราใช้ Rollback)
        for name, cat in self.cat_obj_map.items():
            if name not in excluded_names:
                try: cat.LineColor = Config.BLACK_COLOR
                except AttributeError: pass
        self.doc.Regenerate()

    def apply_print_settings_and_submit(self, ui_data, ps_id_map, vs_id_map):
        pm = self.doc.PrintManager
        pm.SelectNewPrintDriver(ui_data['printer'])
        pm.PrintRange = pm.PrintRange.Select
        pm.PrintToFile = True

        out_dir = ui_data['out_path']
        if not os.path.exists(out_dir): os.makedirs(out_dir)

        pm.CombinedFile = ui_data['is_combine']
        
        # --- Auto Rename Prep ---
        rename_map = {}
        if not pm.CombinedFile and ui_data.get('auto_rename', False):
            sheets_to_print = []
            if ui_data['use_custom_mode']:
                for vid in ui_data['selected_view_ids']:
                    v = get_elem(vid)
                    if v and isinstance(v, ViewSheet):
                        sheets_to_print.append(v)
            else:
                vs_id = vs_id_map.get(ui_data['view_set'])
                if vs_id:
                    vs_elem = get_elem(vs_id)
                    if vs_elem:
                        for v in vs_elem.Views:
                            if isinstance(v, ViewSheet):
                                sheets_to_print.append(v)
            
            for sheet in sheets_to_print:
                sheet_num = sheet.SheetNumber
                sheet_name = sheet.Name
                param = sheet.LookupParameter(Config.TARGET_PARAM_NAME)
                new_filename = u"{} - {}.pdf".format(sheet_num, sheet_name)
                if param and param.HasValue:
                    val = param.AsString()
                    if val and val.strip() != "":
                        clean_val = re.sub(r'[\\/*?:"<>|]', "", val)
                        new_filename = u"{}.pdf".format(clean_val)
                rename_map[sheet_num] = new_filename

        if not pm.CombinedFile:
            if ui_data.get('auto_rename', False):
                if not os.path.exists(Config.TEMP_PDF_FOLDER):
                    os.makedirs(Config.TEMP_PDF_FOLDER)
                pm.PrintToFileName = os.path.join(Config.TEMP_PDF_FOLDER, "Sheet.pdf")
            else:
                pm.PrintToFileName = os.path.join(out_dir, "Sheet.pdf")
        else:
            pm.PrintToFileName = os.path.join(out_dir, "Combined_Print.pdf")

        # Apply Print Setting
        ps_id = ps_id_map.get(ui_data['print_setting'])
        if ps_id:
            ps_elem = get_elem(ps_id)
            if ps_elem: pm.PrintSetup.CurrentPrintSetting = ps_elem

        # Apply ViewSheetSet (Custom หรือ Set เดิม)
        if not ui_data['use_custom_mode']:
            vs_id = vs_id_map.get(ui_data['view_set'])
            if vs_id:
                vs_elem = get_elem(vs_id)
                if vs_elem: pm.ViewSheetSetting.CurrentViewSheetSet = vs_elem
        else:
            if ui_data['selected_view_ids']:
                vss = pm.ViewSheetSetting
                view_set = ViewSet()
                for vid in ui_data['selected_view_ids']:
                    v = get_elem(vid)
                    if v: view_set.Insert(v)
                vss.CurrentViewSheetSet.Views = view_set
                try: vss.SaveAs("_MTC_Temp_Set")
                except Exception: vss.Save()
            else:
                raise Exception(u"⚠️ ไม่ได้เลือก Sheet/View ใดเลยในโหมด Custom")

        pm.SubmitPrint()
        return rename_map

# ==========================================
# 3. UI Controller (จัดการหน้าจออย่างเดียว)
# ==========================================
class PrintUIController(object):
    def __init__(self, printer_list):
        self.checkbox_map = {}
        self.sheet_checkbox_map = {}
        self.sorted_names = []
        self.printable_items = []
        
        self.ui_data = {} # เก็บค่าสรุปตอนกดยืนยัน

        self.window = XamlReader.Load(XmlReader.Create(StringReader(Config.XAML_UI)))

        # Bind Controls
        self.txtSearch = self.window.FindName("txtSearch")
        self.lbCategories = self.window.FindName("lbCategories")
        self.cbPrinter = self.window.FindName("cbPrinter")
        self.cbPrintSetting = self.window.FindName("cbPrintSetting")
        self.cbViewSet = self.window.FindName("cbViewSet")
        self.rbModeSet = self.window.FindName("rbModeSet")
        self.rbModeCustom = self.window.FindName("rbModeCustom")
        self.panelSetMode = self.window.FindName("panelSetMode")
        self.panelCustomMode = self.window.FindName("panelCustomMode")
        self.txtSheetSearch = self.window.FindName("txtSheetSearch")
        self.lbSheets = self.window.FindName("lbSheets")
        self.btnSelectAll = self.window.FindName("btnSelectAll")
        self.btnClearAll = self.window.FindName("btnClearAll")
        self.rbCombine = self.window.FindName("rbCombine")
        self.txtPath = self.window.FindName("txtPath")
        self.btnBrowse = self.window.FindName("btnBrowse")
        self.chkAutoRename = self.window.FindName("chkAutoRename")
        self.chkDoPrint = self.window.FindName("chkDoPrint")
        self.btnRun = self.window.FindName("btnRun")

        # Setup Printer List
        self.cbPrinter.ItemsSource = printer_list
        if Config.TARGET_PRINTER in printer_list: self.cbPrinter.SelectedItem = Config.TARGET_PRINTER
        elif printer_list: self.cbPrinter.SelectedIndex = 0

        # Events
        self.txtSearch.TextChanged += self.on_search_changed
        self.txtSheetSearch.TextChanged += self.on_sheet_search_changed
        self.btnBrowse.Click += self.on_browse
        self.btnRun.Click += self.on_submit
        self.btnSelectAll.Click += self.on_select_all_sheets
        self.btnClearAll.Click += self.on_clear_all_sheets
        self.rbModeSet.Checked += self.on_mode_changed
        self.rbModeCustom.Checked += self.on_mode_changed

    def inject_revit_data(self, print_settings, view_sets, cat_names, printable_items):
        self.sorted_names = sorted(cat_names.keys(), key=natural_sort_key)
        self.printable_items = printable_items
        self.cbPrintSetting.ItemsSource = print_settings
        if print_settings: self.cbPrintSetting.SelectedIndex = 0
        self.cbViewSet.ItemsSource = view_sets
        if view_sets: self.cbViewSet.SelectedIndex = 0

        self.refresh_list("")
        self.refresh_sheet_list("")

    def on_mode_changed(self, sender, e):
        from System.Windows import Visibility
        if self.rbModeCustom.IsChecked:
            self.panelSetMode.Visibility = Visibility.Collapsed
            self.panelCustomMode.Visibility = Visibility.Visible
        else:
            self.panelSetMode.Visibility = Visibility.Visible
            self.panelCustomMode.Visibility = Visibility.Collapsed

    def refresh_list(self, txt):
        self.lbCategories.Items.Clear()
        search = txt.lower()
        for name in self.sorted_names:
            if search in name.lower():
                prev_checked = self.checkbox_map[name].IsChecked if name in self.checkbox_map else any(d_ex == name or d_ex in name.split("  >  ") for d_ex in Config.DEFAULT_EXCLUSIONS)
                cb = CheckBox()
                cb.Content, cb.IsChecked = name, prev_checked
                self.checkbox_map[name] = cb
                self.lbCategories.Items.Add(cb)

    def on_search_changed(self, sender, e): self.refresh_list(self.txtSearch.Text)

    def refresh_sheet_list(self, txt):
        self.lbSheets.Items.Clear()
        search = txt.lower()
        for item in self.printable_items:
            name = item["name"]
            if search in name.lower():
                prev_checked = self.sheet_checkbox_map[name].IsChecked if name in self.sheet_checkbox_map else False
                cb = CheckBox()
                cb.Content, cb.IsChecked = name, prev_checked
                self.sheet_checkbox_map[name] = cb
                self.lbSheets.Items.Add(cb)

    def on_sheet_search_changed(self, sender, e): self.refresh_sheet_list(self.txtSheetSearch.Text)
    def on_select_all_sheets(self, sender, e):
        for cb in self.sheet_checkbox_map.values(): cb.IsChecked = True
    def on_clear_all_sheets(self, sender, e):
        for cb in self.sheet_checkbox_map.values(): cb.IsChecked = False

    def on_browse(self, sender, e):
        dialog = FolderBrowserDialog()
        dialog.Description = u"เลือกโฟลเดอร์สำหรับบันทึกไฟล์ PDF"
        if dialog.ShowDialog() == DialogResult.OK: self.txtPath.Text = dialog.SelectedPath

    def on_submit(self, sender, e):
        self.ui_data = {
            'excluded_names': [n for n, cb in self.checkbox_map.items() if cb.IsChecked],
            'selected_view_ids': [item["id"] for item in self.printable_items if item["name"] in [n for n, cb in self.sheet_checkbox_map.items() if cb.IsChecked]],
            'use_custom_mode': bool(self.rbModeCustom.IsChecked),
            'printer': self.cbPrinter.SelectedItem,
            'print_setting': self.cbPrintSetting.SelectedItem,
            'view_set': self.cbViewSet.SelectedItem,
            'is_combine': bool(self.rbCombine.IsChecked),
            'out_path': self.txtPath.Text,
            'do_print': bool(self.chkDoPrint.IsChecked),
            'auto_rename': bool(self.chkAutoRename.IsChecked)
        }
        self.window.DialogResult = True
        self.window.Close()

# ==========================================
# 4. Main Workflow (ตัวควบคุมหลัก)
# ==========================================
def main_process():
    if not Config.IS_ACTIVE:
        return u"Script is disabled."

    # บังคับปิด Transaction ที่อาจค้างอยู่จาก Node อื่น
    TransactionManager.Instance.ForceCloseTransaction()
    
    out_message = u"UI ถูกยกเลิก"
    rename_map = {}

    try:
        # --- 1. Data Prep ---
        revit_mgr = RevitDataManager(doc)
        ps_list, ps_id_map = revit_mgr.get_print_settings()
        vs_list, vs_id_map = revit_mgr.get_view_sheet_sets()
        printable_items = revit_mgr.get_printable_items()
        cat_names = revit_mgr.get_categories()
        printers_list = PrintSystemManager.get_installed_printers()

        # --- 2. Show UI ---
        ui = PrintUIController(printers_list)
        ui.inject_revit_data(ps_list, vs_list, cat_names, printable_items)

        if ui.window.ShowDialog():
            
            # --- 3. EXECUTION WITH TRANSACTION GROUP ---
            # ใช้ TransactionGroup เป็นเกราะป้องกัน เมื่อสั่ง Rollback ทุกอย่างในนี้จะคืนค่าทั้งหมด
            tg = TransactionGroup(doc, "MTC: Smart Print Process")
            tg.Start()

            try:
                # 3.1 เปลี่ยนสี
                t_color = Transaction(doc, "MTC: Temp Color")
                t_color.Start()
                revit_mgr.change_colors_to_black(ui.ui_data['excluded_names'])
                t_color.Commit()

                # 3.2 สั่ง Print
                if ui.ui_data['do_print']:
                    t_print = Transaction(doc, "MTC: Setup Print")
                    t_print.Start()
                    rename_map = revit_mgr.apply_print_settings_and_submit(ui.ui_data, ps_id_map, vs_id_map)
                    t_print.Commit()

                    # 3.3 รอคิวปริ้น (ขณะที่ยังอยู่ใน TransactionGroup)
                    spool_status = PrintSystemManager.wait_for_spooler(ui.ui_data['printer'])
                    out_message = spool_status
                else:
                    out_message = u"✅ เปลี่ยนสีเสร็จสิ้น (ไม่ได้สั่ง Print)"

            except Exception as inner_e:
                out_message = u"❌ Print Error: " + str(inner_e)

            finally:
                # 3.4 พระเอกของงานนี้! ยกเลิกการกระทำทั้งหมดใน Revit (สีกลับมาเดิม, ลบ Temp Set)
                tg.RollBack()

            # --- 4. Auto-Rename & Move Files ---
            if ui.ui_data.get('do_print') and not ui.ui_data.get('is_combine') and ui.ui_data.get('auto_rename') and rename_map:
                out_message += u"\n⏳ กำลังรอไฟล์จาก PDF24 เพื่อเปลี่ยนชื่อและย้ายโฟลเดอร์..."
                rename_status = PrintSystemManager.watch_and_rename_files(rename_map, ui.ui_data['out_path'])
                out_message += rename_status

    except Exception as e:
        out_message = u"❌ System Error: " + str(e)

    return out_message

# ==========================================
# Run Script
# ==========================================
OUT = main_process()