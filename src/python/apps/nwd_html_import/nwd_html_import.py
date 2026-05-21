import customtkinter as ctk
from tkinter import filedialog, messagebox
import openpyxl
import xlwings as xw 
from PIL import Image as PILImage
from bs4 import BeautifulSoup
import os
import glob
import shutil

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

def browse_excel():
    filepath = filedialog.askopenfilename(
        title="เลือกไฟล์ Excel Clash Report",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    if filepath:
        lbl_excel_path.configure(text=filepath)
        excel_path_var.set(filepath)
        
        try:
            wb = openpyxl.load_workbook(filepath, read_only=True)
            sheets = wb.sheetnames
            if sheets:
                sheet_combo.configure(values=sheets)
                sheet_combo.set(sheets[0])
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถอ่านไฟล์ Excel ได้:\n{e}")

def browse_html_folder():
    folderpath = filedialog.askdirectory(title="เลือกโฟลเดอร์ที่เก็บไฟล์ HTML")
    if folderpath:
        lbl_html_folder.configure(text=folderpath)
        html_folder_var.set(folderpath)

def run_process():
    excel_path = excel_path_var.get()
    html_folder = html_folder_var.get()
    target_sheet = sheet_combo.get()

    if not excel_path or not os.path.exists(excel_path):
        messagebox.showwarning("คำเตือน", "กรุณาเลือกไฟล์ Excel ที่ถูกต้อง")
        return
    if not target_sheet or target_sheet == "ยังไม่ได้เลือกไฟล์ Excel":
        messagebox.showwarning("คำเตือน", "กรุณาเลือก Sheet")
        return
    if not html_folder or not os.path.exists(html_folder):
        messagebox.showwarning("คำเตือน", "กรุณาเลือกโฟลเดอร์ HTML")
        return

    try:
        html_files = glob.glob(os.path.join(html_folder, "*.html"))
        if not html_files:
            messagebox.showinfo("ข้อมูล", "ไม่พบไฟล์ .html ในโฟลเดอร์ที่เลือก")
            return

        new_clashes = []
        for html_path in html_files:
            html_dir = os.path.dirname(html_path)
            
            with open(html_path, 'r', encoding='utf-8') as f:
                soup = BeautifulSoup(f, 'html.parser')

            viewpoints = soup.find_all('div', class_='viewpoint')

            for vp in viewpoints:
                title_tag = vp.find('h2')
                title = title_tag.text.strip() if title_tag else ''
                
                img_tag = vp.find('img')
                image_src = img_tag['src'] if img_tag else ''
                
                full_img_path = ""
                if image_src:
                    full_img_path = os.path.join(html_dir, image_src)
                
                if title:
                    new_clashes.append({
                        'title': title,
                        'image_path': full_img_path
                    })

        # 1. คัดลอกไฟล์ต้นฉบับไปเป็นชื่อใหม่
        base, ext = os.path.splitext(excel_path)
        save_path = f"{base}_Updated{ext}"
        shutil.copy(excel_path, save_path)

        # 2. เปิด Excel แบบซ่อนหน้าต่าง
        app = xw.App(visible=False) 
        
        try:
            # Optimize: ปิดการอัปเดตหน้าจอชั่วคราว
            app.screen_updating = False
            
            wb = app.books.open(save_path)
            ws = wb.sheets[target_sheet]

            max_row = ws.range('C' + str(ws.cells.last_cell.row)).end('up').row
            
            existing_titles = set()
            for row in range(1, max_row + 1):
                val = ws.range(f'C{row}').value
                if val:
                    existing_titles.add(str(val).strip())

            item_no = 1
            for row in range(max_row, 0, -1):
                val = ws.range(f'A{row}').value
                if isinstance(val, (int, float)):
                    item_no = int(val) + 1
                    break

            # เตรียมข้อมูลสำหรับ Batch Write
            new_clashes_to_add = [c for c in new_clashes if c['title'] not in existing_titles]
            added_count = len(new_clashes_to_add)
            
            if added_count > 0:
                start_row = max_row + 1
                end_row = max_row + added_count
                
                batch_data = []
                for clash in new_clashes_to_add:
                    if clash['title'].startswith('C'):
                        folder_name = "CLASH CHECK"
                    elif clash['title'].startswith('V'):
                        folder_name = "VISUAL CHECK"
                    else:
                        folder_name = "ETC."
                    
                    # ข้อมูลแต่ละแถว (Column A ถึง H) - ช่องว่างเว้นไว้ด้วย None
                    row_data = [item_no, folder_name, clash['title'], None, None, None, None, "NEW"]
                    batch_data.append(row_data)
                    item_no += 1
                    existing_titles.add(clash['title'])

                # 1. คัดลอก Format รวดเดียว
                if max_row > 1:
                    ws.range(f'{max_row}:{max_row}').copy()
                    ws.range(f'{start_row}:{end_row}').paste(paste='formats')
                
                # 2. วางข้อมูลข้อความรวดเดียว
                ws.range(f'A{start_row}').value = batch_data
                
                # 3. จัดการรูปภาพ (ต้องวนลูปทีละรูป)
                for i, clash in enumerate(new_clashes_to_add):
                    current_row = start_row + i
                    
                    if clash['image_path'] and os.path.exists(clash['image_path']):
                        try:
                            # --- [สำคัญ] แปลง Path เป็น Absolute Path ของ Windows ---
                            abs_image_path = os.path.normpath(os.path.abspath(clash['image_path']))
                            
                            with PILImage.open(abs_image_path) as pil_img:
                                orig_width, orig_height = pil_img.size
                            
                            target_height_px = 250 
                            ratio = orig_width / orig_height
                            target_height_pt = target_height_px * 0.75
                            target_width_pt = target_height_pt * ratio
                            
                            # ขยาย Row Height (ทำก่อนฝังรูป)
                            ws.range(f'{current_row}:{current_row}').row_height = target_height_pt + 10
                            
                            # ขยาย Column Width G ให้กว้างกว่ารูปภาพ
                            needed_width = ((target_height_px * ratio) - 5) / 7
                            current_width = ws.range('G:G').column_width
                            if needed_width > current_width:
                                ws.range('G:G').column_width = needed_width + 2

                            # แทรกลูปภาพลงในเซลล์
                            cell_left = ws.range(f'G{current_row}').left + 5
                            cell_top = ws.range(f'G{current_row}').top + 5
                            
                            pic = ws.pictures.add(
                                abs_image_path,
                                left=cell_left,
                                top=cell_top,
                                width=target_width_pt,
                                height=target_height_pt
                            )
                            # ล็อกรูปเข้ากับเซลล์ (Move and size with cells)
                            pic.api.Placement = 1 

                        except Exception as img_err:
                            filename = os.path.basename(clash['image_path'])
                            ws.range(f'G{current_row}').value = f"{filename} (Err: {str(img_err)})"
                            ws.range(f'{current_row}:{current_row}').row_height = 15
                    else:
                        ws.range(f'{current_row}:{current_row}').row_height = 15

            wb.save()
            wb.close()
        
        finally:
            app.quit()
            
        messagebox.showinfo("สำเร็จ", f"อัปเดตเรียบร้อย!\n\nตรวจสอบและเพิ่ม Clash ใหม่พร้อมฝังรูปลงเซลล์จำนวน {added_count} รายการ\nไฟล์ถูกบันทึกที่:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", f"เกิดข้อผิดพลาดระหว่างทำงาน:\n{e}")

# --- หน้าต่าง GUI ---
root = ctk.CTk()
root.title("Navisworks Clash to Excel")
root.geometry("650x380")
root.resizable(False, False)

excel_path_var = ctk.StringVar()
html_folder_var = ctk.StringVar()

frame = ctk.CTkFrame(root)
frame.pack(pady=20, padx=20, fill="both", expand=True)

lbl_step1 = ctk.CTkLabel(frame, text="1. เลือกไฟล์ Excel Report:", font=ctk.CTkFont(size=14, weight="bold"))
lbl_step1.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")
btn_excel = ctk.CTkButton(frame, text="Browse Excel", command=browse_excel, width=120)
btn_excel.grid(row=0, column=1, padx=20, pady=(20, 5), sticky="w")
lbl_excel_path = ctk.CTkLabel(frame, text="ยังไม่ได้เลือกไฟล์...", text_color="gray")
lbl_excel_path.grid(row=1, column=0, columnspan=2, padx=20, pady=(0, 10), sticky="w")

lbl_step2 = ctk.CTkLabel(frame, text="2. เลือก Sheet เป้าหมาย:", font=ctk.CTkFont(size=14, weight="bold"))
lbl_step2.grid(row=2, column=0, padx=20, pady=5, sticky="w")
sheet_combo = ctk.CTkOptionMenu(frame, values=["ยังไม่ได้เลือกไฟล์ Excel"], width=250)
sheet_combo.grid(row=2, column=1, padx=20, pady=5, sticky="w")

lbl_step3 = ctk.CTkLabel(frame, text="3. เลือกโฟลเดอร์ HTML:", font=ctk.CTkFont(size=14, weight="bold"))
lbl_step3.grid(row=3, column=0, padx=20, pady=5, sticky="w")
btn_html = ctk.CTkButton(frame, text="Browse Folder", command=browse_html_folder, width=120)
btn_html.grid(row=3, column=1, padx=20, pady=5, sticky="w")
lbl_html_folder = ctk.CTkLabel(frame, text="ยังไม่ได้เลือกโฟลเดอร์...", text_color="gray")
lbl_html_folder.grid(row=4, column=0, columnspan=2, padx=20, pady=(0, 15), sticky="w")

btn_run = ctk.CTkButton(frame, text="อัปเดตข้อมูลลง Excel", command=run_process, 
                        fg_color="#28a745", hover_color="#218838", 
                        font=ctk.CTkFont(size=15, weight="bold"), height=40)
btn_run.grid(row=5, column=0, columnspan=2, padx=20, pady=20, sticky="ew")

root.mainloop()