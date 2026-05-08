# ไฟล์ dynamic_ui_template.py
import tkinter as tk
from tkinter import ttk

class UniversalUI:
    def __init__(self, title, columns_config, row_data):
        """
        :param title: หัวข้อหน้าต่าง
        :param columns_config: รายชื่อ Column (เช่น ["Name", "Category", "Mark"])
        :param row_data: ข้อมูลที่จะใส่ (เช่น [["Filter A", "Walls", "01"], ["Filter B", "Floors", "02"]])
        """
        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry("600x400")

        # 1. ส่วนหัวข้อ (Dynamic Label)
        label = tk.Label(self.root, text=title, font=("Arial", 12, "bold"))
        label.pack(pady=10)

        # 2. ส่วนตาราง (Dynamic Columns)
        # สร้าง Treeview โดยกำหนดจำนวน Column ตามที่ส่งเข้ามาใน columns_config
        self.tree = ttk.Treeview(self.root, columns=columns_config, show='headings')
        
        for col in columns_config:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 3. ใส่ข้อมูลลงในตาราง
        for row in row_data:
            self.tree.insert("", tk.END, values=row)

        # 4. ปุ่มปิด
        tk.Button(self.root, text="ปิดหน้าต่าง", command=self.root.destroy).pack(pady=10)

    def show(self):
        self.root.mainloop()