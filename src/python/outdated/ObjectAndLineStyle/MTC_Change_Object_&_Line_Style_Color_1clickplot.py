# -*- coding: utf-8 -*-
"""
Script: MTC Change Object & Line Style Color 1clickplot 
Purpose: Change Object & Line Style color to black, then Plot to PDF via selected Printer
"""
# Line 5-6: Load CLR to use .NET framework in Python
import clr
# Line 7: System environment management
import sys
# Line 8: .NET System namespace access
import System
# Line 9: Regular Expression for text pattern matching/splitting (used in natural_sort_key)
import re
# Line 10: File/directory path management (used for output folder operations)
import os
# Line 11: Time management for print queue timeout checking (used in polling loop)
import time

# ========== REVIT API IMPORTS ==========
# Line 13-15: Load Revit Services to get current document
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
# Line 15: Get current Revit document (doc variable used throughout script)
doc = DocumentManager.Instance.CurrentDBDocument

# Line 17-18: Load Revit DB API for Color, Transaction, FilteredElementCollector, ViewSheetSet, PrintSetting
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import Color as RevitColor, Transaction, FilteredElementCollector, ViewSheetSet, PrintSetting

# ========== WPF & WINDOWS UI IMPORTS ==========
# Line 20-24: Load WPF assemblies for UI dialog creation
clr.AddReference("PresentationFramework")  # WPF Framework base
clr.AddReference("PresentationCore")        # WPF Core functionality
clr.AddReference("System.Xml")              # XAML parsing support
clr.AddReference("System.Printing")         # Printer queue management
clr.AddReference("System.Windows.Forms")    # Windows Forms components (FolderBrowserDialog)

# Line 26-32: Import specific UI components and helpers
from System.Windows.Markup import XamlReader            # Parse XAML string to create UI window
from System.IO import StringReader                      # Convert XAML to stream
from System.Xml import XmlReader                        # XML reader for XAML parsing
from System.Windows.Controls import CheckBox            # Checkbox UI element for categories
from System.Drawing.Printing import PrinterSettings     # Get list of installed printers
from System.Printing import LocalPrintServer            # Monitor printer queue for job completion
from System.Windows.Forms import FolderBrowserDialog, DialogResult  # Folder selection UI

# ========== CONFIGURATION & SETUP ==========
# Line 36: Check if input IN[0] is boolean to activate/deactivate script
is_active = IN[0] if isinstance(IN[0], bool) else True
# Line 37: Default categories to exclude from color change (checked in UI default state)
default_exclusions = ["Revision Clouds", "Revision Cloud Tags"]
# Line 38: Black color object (RGB: 0,0,0) used for changing line colors
black_color = RevitColor(0, 0, 0)
# Line 39: Default printer name for auto-selection when loading UI
target_printer_name = "PDF24"

# ========== HELPER FUNCTION: NATURAL SORT ==========
# Line 42-43: Function for natural sorting - sorts numbers numerically (1,2,10 not 1,10,2)
def natural_sort_key(text):
    # Line 43: Split by digits, convert numeric strings to int for proper sorting
    return [int(c) if c.isdigit() else c.lower() for c in re.split('([0-9]+)', text)]

# ========== DATA PREPARATION ==========
# Line 49: Get all installed printers sorted naturally (used in UI dropdown at Line 291)
printers_list = sorted([p for p in PrinterSettings.InstalledPrinters], key=natural_sort_key)

# Line 50-51: Get all PrintSetting objects from document and extract names
ps_elems = FilteredElementCollector(doc).OfClass(PrintSetting).ToElements()
# Line 51: Extract PrintSetting names for UI dropdown (used Line 293, 407)
print_settings_list = sorted([ps.Name for ps in ps_elems if ps.Name], key=natural_sort_key)

# Line 53-54: Get all ViewSheetSet objects from document
vs_elems = FilteredElementCollector(doc).OfClass(ViewSheetSet).ToElements()
# Line 54: Extract ViewSheetSet names excluding system sets (starting with "<")
# Used in UI Line 294 and print process Line 409
view_sets_list = sorted([vs.Name for vs in vs_elems if vs.Name and not vs.Name.startswith("<")], key=natural_sort_key)

# Line 56-65: Build dictionary mapping category names to category objects
# Structure: {category_name: category_object, "Parent  >  Child": subcategory_object}
item_data_map = {}
# Line 57: Loop through all categories in Revit document
for cat in doc.Settings.Categories:
    # Line 58: Try-except to handle categories without expected properties
    try:
        # Line 59: Skip DWG-related categories (they don't have SubCategories)
        if ".dwg" in cat.Name.lower(): continue
        # Line 60: Add main category to dictionary (used Line 320 for checkbox creation)
        item_data_map[cat.Name] = cat
        # Line 61: Loop through subcategories of this category
        for subcat in cat.SubCategories:
            # Line 62: Skip DWG subcategories
            if ".dwg" not in subcat.Name.lower():
                # Line 63: Add subcategory with hierarchical naming format
                item_data_map["{}  >  {}".format(cat.Name, subcat.Name)] = subcat
    # Line 64: Catch AttributeError when category missing SubCategories property
    except AttributeError:
        # Line 65: Continue to next category
        pass

# ========== UI XAML DEFINITION ==========
# Line 68-99: Define entire UI layout in XAML string (loaded at Line 282)
XAML_UI = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Itti Drawing Plot Manager" Width="500" Height="800" Background="#F5F5F5" Topmost="True" WindowStartupLocation="CenterScreen">
    <StackPanel Margin="20">
        <!-- Section 1: Category Selection -->
        <Label Content="1. Select items to Exclude:" FontWeight="Bold"/>
        <TextBox Name="txtSearch" Height="25" Margin="0,5,0,5"/>          <!-- Search box for filtering categories -->
        <ListBox Name="lbCategories" Height="200"/>                        <!-- Checkbox list display -->
        
        <!-- Section 2: Plot Settings -->
        <Label Content="2. Configure Plot Settings:" FontWeight="Bold" Margin="0,15,0,0"/>
        <TextBlock Text="Select Printer:" Margin="5,5,0,0"/>
        <ComboBox Name="cbPrinter" Margin="0,2,0,10"/>                     <!-- Printer dropdown -->
        <TextBlock Text="Select Print Setting:" Margin="5,5,0,0"/>
        <ComboBox Name="cbPrintSetting" Margin="0,2,0,10"/>                <!-- Print setting dropdown -->
        <TextBlock Text="Select View / Sheet Set:" Margin="5,5,0,0"/>
        <ComboBox Name="cbViewSet" Margin="0,2,0,10"/>                     <!-- View/Sheet set dropdown -->
        
        <!-- File Format Options -->
        <TextBlock Text="File Format:" FontWeight="Bold" Margin="5,0,0,5"/>
        <RadioButton Name="rbCombine" Content="Combine views/sheets into one file" IsChecked="True" Margin="5,0,0,5"/>
        <RadioButton Name="rbSeparate" Content="Create separate files" Margin="5,0,0,10"/>
        
        <!-- Output Folder Selection -->
        <TextBlock Text="Output Folder (for PDF files):" Margin="5,0,0,2"/>
        <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
            <TextBox Name="txtPath" Width="350" Height="25" Margin="5,0,5,0" Text="C:\Temp"/>
            <Button Name="btnBrowse" Content="Browse..." Width="80" Height="25"/>
        </StackPanel>
        
        <!-- Auto Print Checkbox -->
        <CheckBox Name="chkDoPrint" Content="Execute Plot immediately (with smart queue check)" IsChecked="True" Margin="0,0,0,10"/>
        
        <!-- Run Button -->
        <Button Name="btnRun" Content="RUN PROCESS" Height="50" Background="#212121" Foreground="#FFFFFF" FontWeight="Bold"/>
    </StackPanel>
</Window>
"""

# ========== UI CLASS DEFINITION ==========
# Line 101-175: FullProcessSelector class manages UI dialog and user interactions
class FullProcessSelector(object):
    # Line 102: Constructor initialization
    def __init__(self, item_map, p_list, v_list, printer_list):
        # Line 102: Store reference to category dictionary (used throughout class)
        self.item_map = item_map
        # Line 103: Create sorted list of all category names (used Line 320 for checkbox list)
        self.sorted_names = sorted(self.item_map.keys(), key=natural_sort_key)
        # Line 104: Dictionary to store checkbox controls by name (used Line 319, 358)
        self.checkbox_map = {}
        # Line 105: List to store user-selected exclusions (used Line 357 in on_submit)
        self.excluded_list = []
        
        # Line 107: Load XAML string and create window (uses XamlReader, XmlReader, StringReader)
        self.window = XamlReader.Load(XmlReader.Create(StringReader(XAML_UI)))
        
        # ========== CONTROL BINDING ==========
        # Line 109-118: Bind UI controls from XAML to Python variables
        # Line 110: Search textbox for filtering categories (connected to Line 316 on_search_changed)
        self.txtSearch = self.window.FindName("txtSearch")
        # Line 111: Categories listbox (updated Line 320 in refresh_list)
        self.lbCategories = self.window.FindName("lbCategories")
        # Line 112: Printer dropdown (populated Line 291 with printers_list)
        self.cbPrinter = self.window.FindName("cbPrinter")
        # Line 113: Print setting dropdown (populated Line 293 with print_settings_list)
        self.cbPrintSetting = self.window.FindName("cbPrintSetting")
        # Line 114: View/Sheet set dropdown (populated Line 294 with view_sets_list)
        self.cbViewSet = self.window.FindName("cbViewSet")
        # Line 115: Combine files radio button (checked Line 404 for file combining)
        self.rbCombine = self.window.FindName("rbCombine")
        # Line 116: Output path textbox (read Line 415 for output directory)
        self.txtPath = self.window.FindName("txtPath")
        # Line 117: Browse folder button (event handler Line 314 on_browse)
        self.btnBrowse = self.window.FindName("btnBrowse")
        # Line 118: Auto-print checkbox (checked Line 389 for print execution)
        self.chkDoPrint = self.window.FindName("chkDoPrint")
        # Line 119: Main RUN button (event handler Line 335 on_submit)
        self.btnRun = self.window.FindName("btnRun")
        
        # ========== POPULATE CONTROLS ==========
        # Line 121-127: Set data and default selections for combo boxes
        # Line 121: Set printer list as dropdown items
        self.cbPrinter.ItemsSource = printer_list
        # Line 122: Try to select default printer (PDF24) if available
        if target_printer_name in printer_list: self.cbPrinter.SelectedItem = target_printer_name
        # Line 123: Otherwise select first printer
        elif printer_list: self.cbPrinter.SelectedIndex = 0
        
        # Line 125: Set print settings list
        self.cbPrintSetting.ItemsSource = p_list
        # Line 126: Select first item if available
        if p_list: self.cbPrintSetting.SelectedIndex = 0
        
        # Line 128: Set view/sheet sets list
        self.cbViewSet.ItemsSource = v_list
        # Line 129: Select first item if available
        if v_list: self.cbViewSet.SelectedIndex = 0
        
        # Line 131: Populate category checkbox list (calls refresh_list method)
        self.refresh_list("")
        
        # ========== EVENT HANDLERS ==========
        # Line 133-135: Connect UI events to handler methods
        # Line 133: Search textbox text change event triggers on_search_changed
        self.txtSearch.TextChanged += self.on_search_changed
        # Line 134: Browse button click event triggers on_browse
        self.btnBrowse.Click += self.on_browse
        # Line 135: RUN button click event triggers on_submit
        self.btnRun.Click += self.on_submit
        
    # Line 137: Show dialog and return result (True=OK, False=Cancel)
    def ShowDialog(self): 
        return self.window.ShowDialog()
        
    # ========== BROWSE FOLDER EVENT HANDLER ==========
    # Line 140-142: Handle Browse button click
    def on_browse(self, sender, e):
        # Line 141: Create folder browser dialog instance
        dialog = FolderBrowserDialog()
        # Line 142: Set dialog description text
        dialog.Description = u"Select folder to save PDF files"
        # Line 143: If user clicks OK, update txtPath with selected folder
        if dialog.ShowDialog() == DialogResult.OK: self.txtPath.Text = dialog.SelectedPath

    # ========== CATEGORY LIST REFRESH ==========
    # Line 145-155: Refresh category list based on search text filter
    def refresh_list(self, txt):
        # Line 146: Clear current list items
        self.lbCategories.Items.Clear()
        # Line 147: Convert search text to lowercase for case-insensitive search
        search = txt.lower()
        # Line 148: Loop through all category names from item_data_map
        for name in self.sorted_names:
            # Line 149: Check if search text appears in category name
            if search in name.lower():
                # Line 150: If checkbox doesn't exist for this category, create it
                if name not in self.checkbox_map:
                    # Line 151: Create new CheckBox control
                    cb = CheckBox()
                    # Line 152: Set checkbox label to category name
                    cb.Content = name
                    # Line 153-154: Check if category is in default_exclusions list
                    # If yes, mark checkbox as checked (pre-selected for exclusion)
                    if any(d_ex == name or d_ex in name.split("  >  ") for d_ex in default_exclusions):
                        cb.IsChecked = True
                    # Line 155: Store checkbox in dictionary for later use
                    self.checkbox_map[name] = cb
                # Line 156: Add checkbox to listbox display
                self.lbCategories.Items.Add(self.checkbox_map[name])

    # ========== SEARCH TEXT CHANGED EVENT ==========
    # Line 158: Handle search textbox text change
    def on_search_changed(self, sender, e): 
        # Line 158: Re-populate list with filtered results based on new search text
        self.refresh_list(self.txtSearch.Text)

    # ========== SUBMIT BUTTON HANDLER ==========
    # Line 160-165: Handle RUN button click - collect selected exclusions and close dialog
    def on_submit(self, sender, e):
        # Line 161-162: List comprehension to collect all checked categories
        # Get category objects from item_map for all checked checkboxes
        self.excluded_list = [self.item_map[name] for name, cb in self.checkbox_map.items() if cb.IsChecked]
        # Line 163: Set dialog result to indicate successful completion (True = OK clicked)
        self.window.DialogResult = True
        # Line 164: Close the dialog window
        self.window.Close()

# ========== MAIN PROCESS LOGIC ==========
# Line 168: Initialize output message variable
OUT_MSG = u"UI was cancelled"

# Line 170: Check if script is active (from Line 36: is_active)
if is_active:
    # Line 171: Try-except for main process error handling
    try:
        # Line 172: Create UI form instance with prepared data lists
        form = FullProcessSelector(item_data_map, print_settings_list, view_sets_list, printers_list)
        # Line 173: Show dialog and wait for user confirmation
        if form.ShowDialog():
            
            # ========== CHANGE COLOR TO BLACK ==========
            # Line 176: Dictionary to store original colors for restoration later
            original_colors = {}
            # Line 177: Create Revit transaction for temporary black color change
            t_black = Transaction(doc, "MTC: Temp Black")
            # Line 178: Start transaction
            t_black.Start()
            # Line 179: Loop through all categories in item_data_map
            for name, cat in item_data_map.items():
                # Line 180: Check if category is NOT in excluded list
                if cat not in form.excluded_list:
                    # Line 181: Try-except for categories without LineColor property
                    try:
                        # Line 182: Get current line color of category
                        curr = cat.LineColor
                        # Line 183: Check if color is not already black (RGB: 0,0,0)
                        if curr.Red != 0 or curr.Green != 0 or curr.Blue != 0:
                            # Line 184: Store original color in dictionary for later restoration
                            original_colors[name] = curr
                            # Line 185: Change line color to black
                            cat.LineColor = black_color
                    # Line 186: Catch AttributeError if LineColor doesn't exist
                    except AttributeError: pass
            # Line 187: Commit transaction to save color changes to document
            t_black.Commit()
            
            # ========== PRINT PROCESS ==========
            # Line 190: Check if auto-print is enabled (from chkDoPrint checkbox)
            if form.chkDoPrint.IsChecked:
                # Line 191: Try-except for print process error handling
                try:
                    # Line 192: Get PrintManager from document
                    pm = doc.PrintManager
                    # Line 193: Get selected printer name from dropdown
                    selected_printer = form.cbPrinter.SelectedItem
                    
                    # Line 195: Create transaction for print setup
                    t_print = Transaction(doc, "MTC: Setup Print")
                    # Line 196: Start transaction
                    t_print.Start()
                    
                    # Line 198: Select the printer driver by name
                    pm.SelectNewPrintDriver(selected_printer)
                    # Line 199: Set print range to selected items/sheets
                    pm.PrintRange = pm.PrintRange.Select
                    # Line 200: Enable print to file (PDF)
                    pm.PrintToFile = True 
                    # Line 201: Get output directory from UI textbox
                    out_dir = form.txtPath.Text
                    # Line 202: Create output directory if it doesn't exist
                    if not os.path.exists(out_dir): os.makedirs(out_dir)

                    # Line 204: Check "Combine" radio button selection
                    pm.CombinedFile = form.rbCombine.IsChecked
                    # Line 205: Set filename based on combine setting
                    filename = "Combined_Print.pdf" if pm.CombinedFile else "Sheet.pdf"
                    # Line 206: Construct full output file path
                    pm.PrintToFileName = os.path.join(out_dir, filename)
                    
                    # ========== SET PRINT SETTINGS ==========
                    # Line 209: Use Python generator next() to find selected PrintSetting
                    # Returns first match or None if not found (more efficient than loop)
                    selected_ps = next((ps for ps in ps_elems if ps.Name == form.cbPrintSetting.SelectedItem), None)
                    # Line 210: If PrintSetting found, set as current print setting
                    if selected_ps: pm.PrintSetup.CurrentPrintSetting = selected_ps
                    
                    # Line 212: Use generator next() to find selected ViewSheetSet object
                    selected_vs = next((vs for vs in vs_elems if vs.Name == form.cbViewSet.SelectedItem), None)
                    # Line 213: If ViewSheetSet found, set as current view/sheet set
                    if selected_vs: pm.ViewSheetSetting.CurrentViewSheetSet = selected_vs
                            
                    # Line 215: Commit print setup transaction
                    t_print.Commit()
                    # Line 216: Submit print job to printer
                    pm.SubmitPrint()
                    
                    # ========== SMART QUEUE POLLING ==========
                    # Line 219: Try-except for printer queue checking (may not work on all systems)
                    try:
                        # Line 220: Connect to local print server
                        print_server = LocalPrintServer()
                        # Line 221: Get print queue for selected printer
                        print_queue = print_server.GetPrintQueue(selected_printer)
                        # Line 222: Wait 1 second for job to reach printer queue
                        time.sleep(1)
                        # Line 223: Set 2-minute timeout to prevent application freeze
                        timeout_limit = 120
                        # Line 224: Record start time for timeout calculation
                        start_time = time.time()
                        
                        # Line 226: Loop while waiting for print queue to empty
                        while True:
                            # Line 227: Refresh printer queue status
                            print_queue.Refresh()
                            # Line 228: Check if print queue is empty (all jobs completed)
                            if print_queue.NumberOfJobs == 0:
                                # Line 229: Update status message
                                OUT_MSG = u"Completed! Restoring colors..."
                                # Line 230: Exit polling loop
                                break
                            # Line 231: Check if timeout exceeded (2 minutes)
                            if (time.time() - start_time) > timeout_limit:
                                # Line 232: Update status message with timeout warning
                                OUT_MSG = u"Queue check timeout (2 min) - restoring colors..."
                                # Line 233: Exit polling loop
                                break
                            # Line 234: Wait 1 second before checking queue again
                            time.sleep(1)
                            
                    # Line 236: If queue check fails, use fallback delay
                    except Exception:
                        # Line 237: Wait 5 seconds as fallback delay
                        time.sleep(5)
                        # Line 238: Update status message with fallback info
                        OUT_MSG = u"Plot submitted (queue check failed, using 5s fallback)"
                        
                # Line 240: Catch print process exceptions
                except Exception as pe: 
                    # Line 240: Update message with error details
                    OUT_MSG = u"Print Error: " + str(pe)

            # ========== RESTORE ORIGINAL COLORS ==========
            # Line 243: Create transaction to restore colors
            t_restore = Transaction(doc, "MTC: Restore Color")
            # Line 244: Start transaction
            t_restore.Start()
            # Line 245: Loop through stored original colors
            for name, color in original_colors.items():
                # Line 246: Try-except for color restoration
                try: 
                    # Line 246: Restore original color from stored dictionary
                    item_data_map[name].LineColor = color
                # Line 247: Ignore errors if color restoration fails
                except Exception: pass
            # Line 248: Commit color restoration transaction
            t_restore.Commit()
            
    # Line 250: Catch any main process exceptions
    except Exception as e:
        # Line 251: Update message with error details
        OUT_MSG = u"Error: " + str(e)

# Line 253: Return final output message to Dynamo
OUT = OUT_MSG
