import ezdxf

def create_rectangle_with_dims(msp, x, y, width, height):
    # กำหนดจุด 4 มุมของสี่เหลี่ยม
    p1 = (x, y)
    p2 = (x + width, y)
    p3 = (x + width, y + height)
    p4 = (x, y + height)
    
    # วาดสี่เหลี่ยมด้วย LWPOLYLINE
    msp.add_lwpolyline([p1, p2, p3, p4], close=True)
    
    # ระยะห่างของเส้นบอกระยะ (Dimension) จากขอบสี่เหลี่ยม
    offset = 10
    
    # เส้นบอกระยะด้านล่าง (แนวนอน)
    dim1 = msp.add_linear_dim(base=(x + width/2, y - offset), p1=p1, p2=p2, dimstyle='EZDXF')
    dim1.render()
    
    # เส้นบอกระยะด้านขวา (แนวตั้ง มุม 90 องศา)
    dim2 = msp.add_linear_dim(base=(x + width + offset, y + height/2), p1=p2, p2=p3, angle=90, dimstyle='EZDXF')
    dim2.render()
    
    # เส้นบอกระยะด้านบน (แนวนอน)
    dim3 = msp.add_linear_dim(base=(x + width/2, y + height + offset), p1=p4, p2=p3, dimstyle='EZDXF')
    dim3.render()
    
    # เส้นบอกระยะด้านซ้าย (แนวตั้ง มุม 90 องศา)
    dim4 = msp.add_linear_dim(base=(x - offset, y + height/2), p1=p1, p2=p4, angle=90, dimstyle='EZDXF')
    dim4.render()

def main():
    # สร้างไฟล์ DXF ใหม่ (ใช้ฟอร์แมต R2010 ซึ่งเข้ากันได้ดีกับ AutoCAD ทุกเวอร์ชัน)
    # setup=True จะช่วยตั้งค่า Dimension Style พื้นฐานให้ชื่อว่า 'EZDXF'
    doc = ezdxf.new('R2010', setup=True)
    msp = doc.modelspace()
    
    # วาดสี่เหลี่ยม 3 ขนาดในตำแหน่งต่างๆ (x, y, กว้าง, สูง)
    create_rectangle_with_dims(msp, 0, 0, 50, 30)
    create_rectangle_with_dims(msp, 100, 0, 80, 80)
    create_rectangle_with_dims(msp, 250, 0, 40, 120)
    
    # บันทึกไฟล์ DXF
    filename = 'rectangles_with_dims.dxf'
    doc.saveas(filename)
    print(f"Created {filename} successfully!")

if __name__ == '__main__':
    main()
