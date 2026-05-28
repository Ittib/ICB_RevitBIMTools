from plyer import notification
import time

def start_focus(minutes):
    # แจ้งเตือนเริ่มต้น
    notification.notify(
        title="เริ่มโหมดโฟกัส! 🚀",
        message=f"ระบบจะแจ้งเตือนคุณในอีก {minutes} นาที",
        app_name="Focus Tool",
        timeout=10
    )
    
    # หน่วงเวลา (แปลงนาทีเป็นวินาที)
    time.sleep(minutes * 60)
    
    # แจ้งเตือนเมื่อครบเวลา
    notification.notify(
        title="หมดเวลาแล้ว! ☕",
        message="พักสายตา และลุกไปยืดเส้นยืดสายหน่อยนะ",
        app_name="Focus Tool",
        timeout=10
    )

if __name__ == "__main__":
    # ทดสอบรันที่ 10 วินาที (0.16 นาที)
    start_focus(0.16)