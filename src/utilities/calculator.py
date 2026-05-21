import customtkinter as ctk
import sympy as sp

# ตั้งค่าธีมเริ่มต้น (System = เปลี่ยนตาม OS, Dark = มืดเสมอ, Light = สว่างเสมอ)
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PyQalculateModern(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PyQalculate (Modern Edition)")
        self.geometry("400x600")
        self.minsize(350, 500)

        # ================= UI SETUP =================
        # จัดการ Grid Layout หลัก
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)  # ส่วน History ให้ขยายได้
        self.grid_rowconfigure(1, weight=0)  # ส่วน Input
        self.grid_rowconfigure(2, weight=2)  # ส่วน แป้นพิมพ์

        # 1. ส่วนแสดงประวัติการคำนวณ (History)
        self.history_text = ctk.CTkTextbox(self, font=("Consolas", 16), state='disabled', wrap='word')
        self.history_text.grid(row=0, column=0, sticky="nsew", padx=15, pady=(15, 10))

        # 2. ช่องพิมพ์สมการ (Input Entry)
        self.input_var = ctk.StringVar()
        self.input_entry = ctk.CTkEntry(self, textvariable=self.input_var, font=("Consolas", 20), height=50)
        self.input_entry.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 15))
        
        # ผูกปุ่ม Enter บนคีย์บอร์ดกับการคำนวณ
        self.input_entry.bind("<Return>", self.calculate)
        self.input_entry.focus()

        # 3. แป้นพิมพ์ (Keypad)
        self.keypad_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.keypad_frame.grid(row=2, column=0, sticky="nsew", padx=15, pady=(0, 15))

        # กำหนด Grid ให้ปุ่มขยายเต็มพื้นที่
        for i in range(4):
            self.keypad_frame.grid_columnconfigure(i, weight=1)
        for i in range(5):
            self.keypad_frame.grid_rowconfigure(i, weight=1)

        buttons = [
            ('7', 0, 0), ('8', 0, 1), ('9', 0, 2), ('/', 0, 3),
            ('4', 1, 0), ('5', 1, 1), ('6', 1, 2), ('*', 1, 3),
            ('1', 2, 0), ('2', 2, 1), ('3', 2, 2), ('-', 2, 3),
            ('0', 3, 0), ('.', 3, 1), ('+', 3, 2), ('=', 3, 3),
            ('(', 4, 0), (')', 4, 1), ('C', 4, 2), ('Del', 4, 3)
        ]

        # สร้างปุ่มและตกแต่งสีให้แตกต่างกัน
        for (text, row, col) in buttons:
            # แยกสีปุ่มเพื่อความสวยงาม
            if text in ['/', '*', '-', '+']:
                fg_color = "#E07A5F" # สีส้มอิฐสำหรับเครื่องหมาย
                hover_color = "#F2CC8F"
            elif text == '=':
                fg_color = "#3D5A80" # สีน้ำเงินเข้มสำหรับเครื่องหมายเท่ากับ
                hover_color = "#98C1D9"
            elif text in ['C', 'Del']:
                fg_color = "#E63946" # สีแดงสำหรับลบ
                hover_color = "#F1FAEE"
            else:
                fg_color = None # ใช้สี default ของธีม
                hover_color = None

            btn = ctk.CTkButton(
                self.keypad_frame, 
                text=text, 
                font=("Helvetica", 20, "bold"),
                fg_color=fg_color,
                hover_color=hover_color,
                command=lambda t=text: self.on_button_click(t)
            )
            btn.grid(row=row, column=col, sticky="nsew", padx=4, pady=4)

    # ================= LOGIC & ENGINE =================
    def on_button_click(self, char):
        if char == '=':
            self.calculate()
        elif char == 'C':
            self.input_var.set("")
        elif char == 'Del':
            current = self.input_var.get()
            self.input_var.set(current[:-1])
        else:
            current = self.input_var.get()
            self.input_var.set(current + char)
            # เลื่อนเคอร์เซอร์ไปตำแหน่งขวาสุดเสมอ
            self.input_entry.icursor("end")

    def calculate(self, event=None):
        expr_str = self.input_var.get()
        if not expr_str.strip(): return

        try:
            # ใช้ SymPy ในการประมวลผลสมการ
            parsed_expr = sp.sympify(expr_str)
            result = parsed_expr.evalf()

            # ตัดทศนิยมทิ้งหากผลลัพธ์เป็นจำนวนเต็ม
            if result == int(result):
                result_str = str(int(result))
            else:
                result_str = str(round(result, 6)) # ปัดเศษกันเลขทศนิยมยาวเกินไป

            output = f"= {result_str}"
        except Exception:
            output = "= Error: Invalid Expression"

        self.append_history(expr_str, output)
        self.input_var.set("") # เคลียร์ช่อง Input

    def append_history(self, expr, result):
        self.history_text.configure(state='normal')
        
        # เพิ่มข้อความประวัติลงไป
        self.history_text.insert("end", f"{expr}\n")
        self.history_text.insert("end", f"{result}\n\n")
        
        # เลื่อนหน้าจอลงไปบรรทัดล่างสุด
        self.history_text.see("end")
        self.history_text.configure(state='disabled')

if __name__ == "__main__":
    app = PyQalculateModern()
    app.mainloop()