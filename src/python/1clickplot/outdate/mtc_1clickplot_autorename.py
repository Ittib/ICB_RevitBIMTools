# -*- coding: utf-8 -*-
"""
Script: MTC Change Object & Line Style Color 1clickplot (with Custom Sheet Selection)
"""
import clr
import System
import re
import os
import time
import shutil

clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    Color as RevitColor, Transaction, FilteredElementCollector,
    ViewSheetSet, PrintSetting, View, ViewSheet, ElementId, ViewSet,
    TransactionStatus
)

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
# 1. Config & Initial Setup
# ==========================================

is_active = IN[0] if isinstance(IN[0], bool) else True
default_exclusions = ["Revision Clouds", "Revision Cloud Tags"]
black_color = RevitColor(0, 0, 0)
target_printer_name = "PDF24"

# --- New Config ---
TEMP_PDF_FOLDER = r"C:\Revit_PDF_Temp"
TARGET_PARAM_NAME = "FileName"


def natural_sort_key(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split('([0-9]+)', text)]

def get_elem(eid):
    try:
        return doc.GetElement(eid)
    except Exception:
        return None

# ==========================================
# 2. Data Preparation — แตะแค่ Printer (ไม่ใช่ Revit API)
# ==========================================

printers_list = sorted(
    [p for p in PrinterSettings.InstalledPrinters],
    key=natural_sort_key
)

# ==========================================
# 3. XAML UI
# ==========================================

XAML_UI = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Itti Drawing Plot Manager" Width="560" Height="920"
        Background="#F5F5F5" Topmost="True" WindowStartupLocation="CenterScreen">
    <ScrollViewer VerticalScrollBarVisibility="Auto">
    <StackPanel Margin="20">

        <Label Content="1. เลือกรายการที่ต้องการ Exclude (ไม่เปลี่ยนเป็นสีดำ):"
               FontWeight="Bold" Background="#E0E0E0" Padding="5"/>
        <TextBox Name="txtSearch" Height="25" Margin="0,5,0,5"/>
        <ListBox Name="lbCategories" Height="180" Margin="0,0,0,10"/>

        <Label Content="2. ตั้งค่า Printer และ Print Setting:"
               FontWeight="Bold" Background="#E0E0E0" Padding="5" Margin="0,5,0,0"/>
        <TextBlock Text="เลือก Printer:" Margin="5,8,0,2"/>
        <ComboBox Name="cbPrinter" Margin="0,0,0,8"/>
        <TextBlock Text="เลือก Print Setting:" Margin="5,0,0,2"/>
        <ComboBox Name="cbPrintSetting" Margin="0,0,0,10"/>

        <Label Content="3. เลือกโหมดการระบุ View/Sheet ที่จะ Plot:"
               FontWeight="Bold" Background="#E0E0E0" Padding="5" Margin="0,5,0,0"/>
        <StackPanel Orientation="Horizontal" Margin="5,8,0,5">
            <RadioButton Name="rbModeSet" Content="ใช้ View/Sheet Set (แบบเดิม)"
                         IsChecked="True" GroupName="PlotMode" Margin="0,0,20,0"/>
            <RadioButton Name="rbModeCustom" Content="เลือก View/Sheet เอง (แบบใหม่)"
                         GroupName="PlotMode"/>
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

        <Label Content="4. ตั้งค่าไฟล์ Output:"
               FontWeight="Bold" Background="#E0E0E0" Padding="5" Margin="0,5,0,0"/>
        <TextBlock Text="รูปแบบการเซฟไฟล์:" Margin="5,8,0,2"/>
        <RadioButton Name="rbCombine"  Content="Combine — รวมเป็นไฟล์เดียว"
                     IsChecked="True" Margin="5,0,0,3"/>
        <RadioButton Name="rbSeparate" Content="Separate — แยกแต่ละ Sheet/View"
                     Margin="5,0,0,10"/>

        <TextBlock Text="โฟลเดอร์สำหรับเซฟไฟล์ (Output Path):" Margin="5,0,0,2"/>
        <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
            <TextBox Name="txtPath" Width="370" Height="25" Margin="5,0,5,0" Text="C:\Temp"/>
            <Button Name="btnBrowse" Content="Browse..." Width="80" Height="25"/>
        </StackPanel>

        <CheckBox Name="chkAutoRename" Content="เปิดใช้ระบบ Auto-Rename (ดึงชื่อไฟล์จาก Parameter 'FileName')"
                  IsChecked="True" Margin="0,0,0,5" Foreground="#0055FF" FontWeight="Bold"/>
        <CheckBox Name="chkDoPrint" Content="สั่ง Plot ทันที (พร้อมระบบเช็คคิวอัจฉริยะ)"
                  IsChecked="True" Margin="0,0,0,10"/>
        <Button Name="btnRun" Content="RUN PROCESS" Height="50"
                Background="#212121" Foreground="#FFFFFF" FontWeight="Bold"/>

    </StackPanel>
    </ScrollViewer>
</Window>
"""

# ==========================================
# 4. UI Class
# ==========================================

class FullProcessSelector(object):

    def __init__(self, printer_list):
        self.checkbox_map       = {} # Dictionary สำหรับเก็บ CheckBox ของ Categories (key: ชื่อ Category, value: CheckBox object)
        self.sheet_checkbox_map = {} # Dictionary สำหรับเก็บ CheckBox ของ Sheets/Views (key: ชื่อ Sheet/View, value: CheckBox object)
        self.excluded_names     = [] # ลิสต์สำหรับเก็บชื่อ Category ที่ผู้ใช้เลือก Exclude
        self.selected_view_ids  = [] # ลิสต์สำหรับเก็บ ElementId ของ View/Sheet ที่ผู้ใช้เลือกในโหมด Custom
        self.use_custom_mode    = False # Flag เพื่อระบุว่าผู้ใช้เลือกโหมด Custom หรือไม่
        self.printable_items    = [] # ลิสต์ของ Dictionary ที่เก็บข้อมูล View/Sheet ที่สามารถพิมพ์ได้ (name, id)
        self.sorted_names       = [] # ลิสต์ของชื่อ Category ที่เรียงลำดับแล้ว

        self.window = XamlReader.Load(XmlReader.Create(StringReader(XAML_UI)))

        self.txtSearch       = self.window.FindName("txtSearch")
        self.lbCategories    = self.window.FindName("lbCategories")
        self.cbPrinter       = self.window.FindName("cbPrinter")
        self.cbPrintSetting  = self.window.FindName("cbPrintSetting")
        self.cbViewSet       = self.window.FindName("cbViewSet")
        self.rbModeSet       = self.window.FindName("rbModeSet")
        self.rbModeCustom    = self.window.FindName("rbModeCustom")
        self.panelSetMode    = self.window.FindName("panelSetMode")
        self.panelCustomMode = self.window.FindName("panelCustomMode")
        self.txtSheetSearch  = self.window.FindName("txtSheetSearch")
        self.lbSheets        = self.window.FindName("lbSheets")
        self.btnSelectAll    = self.window.FindName("btnSelectAll")
        self.btnClearAll     = self.window.FindName("btnClearAll")
        self.rbCombine       = self.window.FindName("rbCombine")
        self.txtPath         = self.window.FindName("txtPath")
        self.btnBrowse       = self.window.FindName("btnBrowse")
        self.chkAutoRename   = self.window.FindName("chkAutoRename")
        self.chkDoPrint      = self.window.FindName("chkDoPrint")
        self.btnRun          = self.window.FindName("btnRun")

        self.cbPrinter.ItemsSource = printer_list
        if target_printer_name in printer_list:
            self.cbPrinter.SelectedItem = target_printer_name
        elif printer_list:
            self.cbPrinter.SelectedIndex = 0

        self.txtSearch.TextChanged      += self.on_search_changed
        self.txtSheetSearch.TextChanged += self.on_sheet_search_changed
        self.btnBrowse.Click            += self.on_browse
        self.btnRun.Click               += self.on_submit
        self.btnSelectAll.Click         += self.on_select_all_sheets
        self.btnClearAll.Click          += self.on_clear_all_sheets
        self.rbModeSet.Checked          += self.on_mode_changed
        self.rbModeCustom.Checked       += self.on_mode_changed

    def inject_revit_data(self, print_settings, view_sets, cat_names, printable_items):
        self.sorted_names    = sorted(cat_names.keys(), key=natural_sort_key)
        self.printable_items = printable_items

        self.cbPrintSetting.ItemsSource = print_settings
        if print_settings: self.cbPrintSetting.SelectedIndex = 0

        self.cbViewSet.ItemsSource = view_sets
        if view_sets: self.cbViewSet.SelectedIndex = 0

        self.refresh_list("")
        self.refresh_sheet_list("")

    def ShowDialog(self):
        return self.window.ShowDialog()

    def on_mode_changed(self, sender, e):
        from System.Windows import Visibility
        if self.rbModeCustom.IsChecked:
            self.panelSetMode.Visibility    = Visibility.Collapsed
            self.panelCustomMode.Visibility = Visibility.Visible
        else:
            self.panelSetMode.Visibility    = Visibility.Visible
            self.panelCustomMode.Visibility = Visibility.Collapsed

    def refresh_list(self, txt):
        self.lbCategories.Items.Clear()
        search = txt.lower()
        for name in self.sorted_names:
            if search in name.lower():
                prev_checked = False
                if name in self.checkbox_map:
                    prev_checked = self.checkbox_map[name].IsChecked
                else:
                    prev_checked = any(
                        d_ex == name or d_ex in name.split("  >  ")
                        for d_ex in default_exclusions
                    )
                cb = CheckBox()
                cb.Content   = name
                cb.IsChecked = prev_checked
                self.checkbox_map[name] = cb
                self.lbCategories.Items.Add(cb)

    def on_search_changed(self, sender, e):
        self.refresh_list(self.txtSearch.Text)

    def refresh_sheet_list(self, txt):
        self.lbSheets.Items.Clear()
        search = txt.lower()
        for item in self.printable_items:
            name = item["name"]
            if search in name.lower():
                prev_checked = False
                if name in self.sheet_checkbox_map:
                    prev_checked = self.sheet_checkbox_map[name].IsChecked
                cb = CheckBox()
                cb.Content   = name
                cb.IsChecked = prev_checked
                self.sheet_checkbox_map[name] = cb
                self.lbSheets.Items.Add(cb)

    def on_sheet_search_changed(self, sender, e):
        self.refresh_sheet_list(self.txtSheetSearch.Text)

    def on_select_all_sheets(self, sender, e):
        for cb in self.sheet_checkbox_map.values():
            cb.IsChecked = True

    def on_clear_all_sheets(self, sender, e):
        for cb in self.sheet_checkbox_map.values():
            cb.IsChecked = False

    def on_browse(self, sender, e):
        dialog = FolderBrowserDialog()
        dialog.Description = u"เลือกโฟลเดอร์สำหรับบันทึกไฟล์ PDF"
        if dialog.ShowDialog() == DialogResult.OK:
            self.txtPath.Text = dialog.SelectedPath

    def on_submit(self, sender, e):
        self.excluded_names = [
            name for name, cb in self.checkbox_map.items() if cb.IsChecked
        ]
        self.selected_view_ids = [
            item["id"] for item in self.printable_items
            if item["name"] in [
                n for n, cb in self.sheet_checkbox_map.items() if cb.IsChecked
            ]
        ]
        self.use_custom_mode     = bool(self.rbModeCustom.IsChecked)
        self.window.DialogResult = True
        self.window.Close()

# ==========================================
# 5. Main Process Logic
# ==========================================

OUT_MSG = u"UI ถูกยกเลิก"

if is_active:
    try:
        form = FullProcessSelector(printers_list)

        TransactionManager.Instance.ForceCloseTransaction()

        ps_elems  = FilteredElementCollector(doc).OfClass(PrintSetting).ToElements()
        ps_id_map = {ps.Name: ps.Id for ps in ps_elems if ps.Name}
        print_settings_list = sorted(ps_id_map.keys(), key=natural_sort_key)

        vs_elems  = FilteredElementCollector(doc).OfClass(ViewSheetSet).ToElements()
        vs_id_map = {vs.Name: vs.Id for vs in vs_elems if vs.Name and not vs.Name.startswith("<")}
        view_sets_list = sorted(vs_id_map.keys(), key=natural_sort_key)

        all_sheets  = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
        sheets_data = sorted(
            [{"name": u"[Sheet] {}: {}".format(s.SheetNumber, s.Name), "id": s.Id}
             for s in all_sheets],
            key=lambda x: natural_sort_key(x["name"])
        )

        all_views  = FilteredElementCollector(doc).OfClass(View).ToElements()
        views_data = sorted(
            [{"name": u"[View] {}".format(v.Name), "id": v.Id}
             for v in all_views
             if not v.IsTemplate and v.CanBePrinted and not hasattr(v, 'SheetNumber')],
            key=lambda x: natural_sort_key(x["name"])
        )
        all_printable = sheets_data + views_data

        cat_names   = {}
        cat_obj_map = {}
        for cat in doc.Settings.Categories:
            try:
                if ".dwg" in cat.Name.lower(): continue
                cat_names[cat.Name]   = cat.Name
                cat_obj_map[cat.Name] = cat
                for subcat in cat.SubCategories:
                    if ".dwg" not in subcat.Name.lower():
                        key = "{}  >  {}".format(cat.Name, subcat.Name)
                        cat_names[key]   = key
                        cat_obj_map[key] = subcat
            except AttributeError:
                pass

        form.inject_revit_data(print_settings_list, view_sets_list, cat_names, all_printable)

        if form.ShowDialog():

            TransactionManager.Instance.ForceCloseTransaction()

            original_colors = {}
            try:
                # --- 1. Change Color to Black ---
                t_black = Transaction(doc, "MTC: Temp Black")
                t_black.Start()
                for name, cat in cat_obj_map.items():
                    if name not in form.excluded_names:
                        try:
                            curr = cat.LineColor
                            if curr.IsValid:
                                original_colors[name] = curr
                            cat.LineColor = black_color
                        except AttributeError:
                            pass
                doc.Regenerate()
                t_black.Commit()

                # --- 2. Print Process ---
                if form.chkDoPrint.IsChecked:
                    try:
                        pm = doc.PrintManager
                        selected_printer = form.cbPrinter.SelectedItem

                        pm.SelectNewPrintDriver(selected_printer)
                        pm.PrintRange = pm.PrintRange.Select
                        pm.PrintToFile = True

                        out_dir = form.txtPath.Text
                        if not os.path.exists(out_dir):
                            os.makedirs(out_dir)

                        pm.CombinedFile = form.rbCombine.IsChecked
                        
                        # --- Auto-Rename Prep ---
                        rename_map = {}
                        if not pm.CombinedFile and form.chkAutoRename.IsChecked:
                            sheets_to_print = []
                            if form.use_custom_mode:
                                for vid in form.selected_view_ids:
                                    v = get_elem(vid)
                                    if v and isinstance(v, ViewSheet):
                                        sheets_to_print.append(v)
                            else:
                                vs_id = vs_id_map.get(form.cbViewSet.SelectedItem)
                                if vs_id:
                                    vs_elem = get_elem(vs_id)
                                    if vs_elem:
                                        for v in vs_elem.Views:
                                            if isinstance(v, ViewSheet):
                                                sheets_to_print.append(v)
                            
                            for sheet in sheets_to_print:
                                sheet_num = sheet.SheetNumber
                                sheet_name = sheet.Name
                                param = sheet.LookupParameter(TARGET_PARAM_NAME)
                                new_filename = u"{} - {}.pdf".format(sheet_num, sheet_name)
                                if param and param.HasValue:
                                    val = param.AsString()
                                    if val and val.strip() != "":
                                        clean_val = __import__('re').sub(r'[\\/*?:"<>|]', "", val)
                                        new_filename = u"{}.pdf".format(clean_val)
                                rename_map[sheet_num] = new_filename

                        if not pm.CombinedFile:
                            if form.chkAutoRename.IsChecked:
                                if not os.path.exists(TEMP_PDF_FOLDER):
                                    os.makedirs(TEMP_PDF_FOLDER)
                                pm.PrintToFileName = os.path.join(TEMP_PDF_FOLDER, "Sheet.pdf")
                            else:
                                pm.PrintToFileName = os.path.join(out_dir, "Sheet.pdf")
                        else:
                            pm.PrintToFileName = os.path.join(out_dir, "Combined_Print.pdf")

                        # Apply Print Setting
                        t_print = Transaction(doc, "MTC: Setup Print")
                        t_print.Start()
                        ps_id = ps_id_map.get(form.cbPrintSetting.SelectedItem)
                        if ps_id:
                            ps_elem = get_elem(ps_id)
                            if ps_elem:
                                pm.PrintSetup.CurrentPrintSetting = ps_elem
                        t_print.Commit()

                        # Apply ViewSheetSet
                        t_vs = Transaction(doc, "MTC: Setup ViewSet")
                        t_vs.Start()

                        if not form.use_custom_mode:
                            vs_id = vs_id_map.get(form.cbViewSet.SelectedItem)
                            if vs_id:
                                vs_elem = get_elem(vs_id)
                                if vs_elem:
                                    pm.ViewSheetSetting.CurrentViewSheetSet = vs_elem
                        else:
                            if form.selected_view_ids:
                                vss      = pm.ViewSheetSetting
                                view_set = ViewSet()
                                for vid in form.selected_view_ids:
                                    v = get_elem(vid)
                                    if v:
                                        view_set.Insert(v)
                                vss.CurrentViewSheetSet.Views = view_set
                                try:
                                    vss.SaveAs("_MTC_Temp_Set")
                                except Exception:
                                    vss.Save()
                            else:
                                t_vs.RollBack()
                                raise Exception(u"⚠️ ไม่ได้เลือก Sheet/View ใดเลยในโหมด Custom")

                        t_vs.Commit()
                        pm.SubmitPrint()

                        # ลบ Temp Set หลัง Submit
                        if form.use_custom_mode:
                            try:
                                t_clean = Transaction(doc, "MTC: Clean Temp Set")
                                t_clean.Start()
                                temp_vs = next(
                                    (vs for vs in FilteredElementCollector(doc)
                                     .OfClass(ViewSheetSet).ToElements()
                                     if vs.Name == "_MTC_Temp_Set"),
                                    None
                                )
                                if temp_vs:
                                    doc.Delete(temp_vs.Id)
                                t_clean.Commit()
                            except Exception:
                                pass

                        # Smart Queue Polling
                        try:
                            print_server = LocalPrintServer()
                            print_queue  = print_server.GetPrintQueue(selected_printer)
                            time.sleep(1)
                            timeout_limit = 120
                            start_time    = time.time()
                            while True:
                                print_queue.Refresh()
                                if print_queue.NumberOfJobs == 0:
                                    OUT_MSG = u"✅ พิมพ์เสร็จสมบูรณ์! ทำการคืนค่าสี..."
                                    break
                                if (time.time() - start_time) > timeout_limit:
                                    OUT_MSG = u"⚠️ Timeout 2 นาที ทำการคืนค่าสี..."
                                    break
                                time.sleep(1)
                        except Exception:
                            time.sleep(5)
                            OUT_MSG = u"✅ ส่งคำสั่ง Plot สำเร็จ (Fallback 5s)"

                    except Exception as pe:
                        OUT_MSG = u"❌ Plot Error: " + str(pe)

            finally:
                # --- 3. Restore Color (Always Run) ---
                if original_colors:
                    t_restore = Transaction(doc, "MTC: Restore Color")
                    try:
                        t_restore.Start()
                        for name, color in original_colors.items():
                            try:
                                cat_obj_map[name].LineColor = color
                            except Exception:
                                pass
                        doc.Regenerate()
                        t_restore.Commit()
                    except Exception:
                        if t_restore.GetStatus() == TransactionStatus.Started:
                            t_restore.RollBack()

            # --- 4. Auto-Rename & Move Files ---
            if form.chkDoPrint.IsChecked and not form.rbCombine.IsChecked and 'rename_map' in locals() and rename_map:
                try:
                    OUT_MSG += u"\n⏳ กำลังรอไฟล์จาก PDF24 เพื่อเปลี่ยนชื่อและย้ายโฟลเดอร์..."
                    processed = 0
                    expected = len(rename_map)
                    timeout = 60 * 10
                    start_t = time.time()
                    
                    while processed < expected:
                        if time.time() - start_t > timeout:
                            OUT_MSG += u"\n⚠️ หมดเวลารอไฟล์จาก PDF24 (Timeout)"
                            break
                            
                        if os.path.exists(TEMP_PDF_FOLDER):
                            for filename in os.listdir(TEMP_PDF_FOLDER):
                                if filename.endswith(".pdf"):
                                    match = __import__('re').search(r" - Sheet - (.*?) - ", filename)
                                    if match:
                                        sheet_num_extracted = match.group(1).strip()
                                        if sheet_num_extracted in rename_map:
                                            old_path = os.path.join(TEMP_PDF_FOLDER, filename)
                                            desired_name = rename_map[sheet_num_extracted]
                                            new_path = os.path.join(out_dir, desired_name)
                                            
                                            if os.path.exists(new_path):
                                                try:
                                                    os.remove(new_path)
                                                except:
                                                    pass
                                            
                                            try:
                                                shutil.move(old_path, new_path)
                                                processed += 1
                                                del rename_map[sheet_num_extracted]
                                            except:
                                                pass
                        time.sleep(1)
                    
                    if processed > 0:
                        OUT_MSG += u"\n✅ เปลี่ยนชื่อและย้ายไฟล์สำเร็จ {} แผ่น".format(processed)
                except Exception as e:
                    OUT_MSG += u"\n❌ Rename Error: " + str(e)

    except Exception as e:
        OUT_MSG = u"❌ Error: " + str(e)

OUT = OUT_MSG