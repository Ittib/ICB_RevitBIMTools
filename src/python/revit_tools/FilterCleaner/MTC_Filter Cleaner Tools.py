import clr
import sys
import System
import re 

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Window, Thickness, WindowStartupLocation
from System.Windows.Controls import Label, Button, StackPanel, ListBox, SelectionMode, Orientation, TextBox, CheckBox
from System.Windows.Media import SolidColorBrush, Color, ColorConverter

# --- 1. รับค่าจาก Inputs ---
input_data = IN[0] if isinstance(IN[0], list) else [IN[0]]
is_active = IN[1]

# --- 2. ฟังก์ชันช่วยเรียงลำดับแบบธรรมชาติ (Natural Sort) ---
def natural_sort_key(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split('([0-9]+)', text)]

def get_brush(hex_code):
    return SolidColorBrush(ColorConverter.ConvertFromString(hex_code))

# --- 3. UI Class ---
class ProElementSelector(Window):
    def __init__(self, items):
        self.Title = "Element Delete Selector"
        self.Width = 450
        self.Height = 750
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = get_brush("#F5F5F5")
        self.Topmost = True
        
        self.item_map = {}
        for item in items:
            try:
                name = item.Name if hasattr(item, "Name") else str(item)
                
                # [แก้ไข] เพิ่ม ID เข้าไปในชื่อ เพื่อป้องกันกรณี Element ชื่อเหมือนกันโดนทับใน Dictionary
                if hasattr(item, "Id"):
                    name = "{} [{}]".format(name, item.Id)
                
                # เลี่ยงชื่อซ้ำเผื่อกรณีไม่มี ID
                original_name = name
                counter = 1
                while name in self.item_map:
                    name = "{} ({})".format(original_name, counter)
                    counter += 1

                self.item_map[name] = item
            except:
                continue
        
        # ใช้ Natural Sort ในการจัดเรียงรายชื่อ
        self.sorted_names = sorted(self.item_map.keys(), key=natural_sort_key)
        self.checkbox_map = {} 
        self.final_output_list = []

        main_layout = StackPanel()
        main_layout.Margin = Thickness(20)

        header = Label()
        header.Content = "เลือกรายการที่ต้องการลบ (เรียงลำดับใหม่)"
        header.FontSize = 16
        header.FontWeight = System.Windows.FontWeights.Bold
        main_layout.Children.Add(header)

        self.search_box = TextBox()
        self.search_box.Height = 25
        self.search_box.Margin = Thickness(0, 10, 0, 10)
        self.search_box.TextChanged += self.on_search_changed
        main_layout.Children.Add(self.search_box)

        btn_row = StackPanel()
        btn_row.Orientation = Orientation.Horizontal
        btn_row.Margin = Thickness(0, 0, 0, 10)

        btn_all = Button()
        btn_all.Content = "Select All"
        btn_all.Width = 100
        btn_all.Click += self.mass_all

        btn_none = Button()
        btn_none.Content = "Select None"
        btn_none.Width = 100
        btn_none.Margin = Thickness(10, 0, 0, 0)
        btn_none.Click += self.mass_none

        btn_row.Children.Add(btn_all)
        btn_row.Children.Add(btn_none)
        main_layout.Children.Add(btn_row)

        self.listbox = ListBox()
        self.listbox.Height = 350
        self.listbox.SelectionMode = SelectionMode.Extended
        self.listbox.Background = get_brush("#FFFFFF")
        
        self.refresh_list("")
        main_layout.Children.Add(self.listbox)

        btn_submit = Button()
        btn_submit.Content = "CONFIRM DELETE (OUT TO LIST)"
        btn_submit.Height = 50
        btn_submit.Background = get_brush("#EC407A")
        btn_submit.Foreground = get_brush("#FFFFFF")
        btn_submit.Margin = Thickness(0, 15, 0, 0)
        btn_submit.Click += self.on_submit
        main_layout.Children.Add(btn_submit)

        self.Content = main_layout

    def refresh_list(self, txt):
        self.listbox.Items.Clear()
        search = txt.lower()
        for name in self.sorted_names:
            if search in name.lower():
                if name not in self.checkbox_map:
                    cb = CheckBox()
                    cb.Content = name
                    cb.Click += self.sync_checks
                    self.checkbox_map[name] = cb
                self.listbox.Items.Add(self.checkbox_map[name])

    def sync_checks(self, sender, e):
        state = sender.IsChecked
        if sender in self.listbox.SelectedItems:
            for item in self.listbox.SelectedItems:
                if isinstance(item, CheckBox):
                    item.IsChecked = state

    def on_search_changed(self, sender, e):
        self.refresh_list(self.search_box.Text)

    def mass_all(self, s, e):
        for name in self.checkbox_map:
            self.checkbox_map[name].IsChecked = True

    def mass_none(self, s, e):
        for name in self.checkbox_map:
            self.checkbox_map[name].IsChecked = False

    def on_submit(self, s, e):
        for name, cb in self.checkbox_map.items():
            if cb.IsChecked:
                self.final_output_list.append(self.item_map[name])
        self.DialogResult = True
        self.Close()

# --- 4. การรัน ---
output_to_dynamo = []

if is_active:
    try:
        form = ProElementSelector(input_data)
        if form.ShowDialog():
            output_to_dynamo = form.final_output_list
        else:
            # [แก้ไข] เปลี่ยนจาก String เป็นลิสต์ว่าง
            output_to_dynamo = [] 
    except Exception as e:
        # [แก้ไข] ถ้าเกิด Error ให้พ่นลิสต์ว่าง จะได้ไม่ไปทำให้โหนด Delete พังซ้ำซ้อน
        output_to_dynamo = [] 
else:
    # [แก้ไข] เปลี่ยนจาก "UI is waiting" เป็นลิสต์ว่าง
    output_to_dynamo = []

OUT = output_to_dynamo
