import customtkinter as ctk
import pandas as pd
import os
from tkinter import filedialog, messagebox

# ==========================================
# [ CONFIGURATION ] - ปรับแต่งสีและฟอนต์ที่นี่
# ==========================================
PRIMARY_COLOR = "#ff8585"      # สีหลัก (เช่น สีปุ่มเริ่ม)
HOVER_COLOR   = "#fc6262"      # สีตอนเอาเมาส์วางบนปุ่มหลัก
SUB_COLOR     = "#ff8585"      # สีรอง (เช่น สีปุ่มเลือกไฟล์/โฟลเดอร์)
SUB_HOVER     = "#ff8585"      # สีตอนเอาเมาส์วางบนปุ่มรอง
BG_COLOR      = "#494949"      # สีพื้นหลังหน้าต่าง

# ชื่อฟอนต์ (ต้องติดตั้งในเครื่องก่อนนะครับ ถ้าไม่มีจะใช้ Segoe UI แทน)
# แนะนำ: 'FC Minimal', 'Inter', 'Kanit'
MAIN_FONT = "FC Minimal" 
# ==========================================

ctk.set_appearance_mode("dark")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Itti File Renamer")
        self.geometry("600x420")
        self.configure(fg_color=BG_COLOR)
        
        # เตรียม Font Styles
        self.h1 = (MAIN_FONT, 24, "bold")
        self.p  = (MAIN_FONT, 14)
        self.btn_f = (MAIN_FONT, 14, "bold")

        # --- UI ELEMENTS ---
        self.label_header = ctk.CTkLabel(self, text="FILE RENAMER", font=self.h1, text_color="white")
        self.label_header.pack(pady=(40, 30))

        # ส่วน Excel
        self.create_input_group("ไฟล์ Excel (Column A):", self.browse_excel, "entry_excel")
        
        # ส่วน Folder
        self.create_input_group("โฟลเดอร์เป้าหมาย:", self.browse_folder, "entry_folder")

        # ปุ่มเริ่มทำงาน (ใช้ PRIMARY_COLOR)
        self.btn_run = ctk.CTkButton(
            self, text="START PROCESS", 
            font=(MAIN_FONT, 18, "bold"),
            height=55, width=300,
            fg_color=PRIMARY_COLOR, 
            hover_color=HOVER_COLOR,
            corner_radius=10,
            command=self.start_rename
        )
        self.btn_run.pack(pady=(40, 0))

        self.label_footer = ctk.CTkLabel(self, text="v2.0 Minimalist Edition", font=(MAIN_FONT, 10), text_color="#555")
        self.label_footer.pack(side="bottom", pady=10)

    def create_input_group(self, label_text, command, entry_name):
        """สร้างกลุ่ม Input แบบ Minimal"""
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(pady=10, padx=50, fill="x")
        
        label = ctk.CTkLabel(frame, text=label_text, font=self.p, text_color="#bbb")
        label.pack(anchor="w", padx=5)

        inner_frame = ctk.CTkFrame(frame, fg_color="transparent")
        inner_frame.pack(fill="x")

        entry = ctk.CTkEntry(inner_frame, placeholder_text="...", height=35, font=self.p, border_width=1)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        setattr(self, entry_name, entry)

        btn = ctk.CTkButton(inner_frame, text="BROWSE", width=90, height=35, 
                            font=self.btn_f, fg_color=SUB_COLOR, hover_color=SUB_HOVER,
                            command=command)
        btn.pack(side="right")

    # --- FUNCTIONS ---
    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.entry_excel.delete(0, "end")
            self.entry_excel.insert(0, path)

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.entry_folder.delete(0, "end")
            self.entry_folder.insert(0, path)

    def start_rename(self):
        excel_path = self.entry_excel.get()
        folder_path = self.entry_folder.get()

        if not excel_path or not folder_path:
            messagebox.showwarning("Warning", "กรุณาเลือกข้อมูลให้ครบ")
            return

        try:
            files = sorted([f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))])
            df = pd.read_excel(excel_path, header=None)
            new_names = df.iloc[:, 0].dropna().tolist()

            count = 0
            for i in range(min(len(files), len(new_names))):
                old_path = os.path.join(folder_path, files[i])
                ext = os.path.splitext(files[i])[1]
                new_n = str(new_names[i]).strip()
                new_n = new_n if new_n.endswith(ext) else new_n + ext
                
                os.rename(old_path, os.path.join(folder_path, new_n))
                count += 1
            
            messagebox.showinfo("Done", f"เปลี่ยนชื่อสำเร็จ {count} ไฟล์")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = App()
    app.mainloop()