# 📊 Robot Framework + ExcelSage: Automate Excel Processing

ระบบทดสอบที่ใช้ Robot Framework และไลบรารี [ExcelSage](https://pypi.org/project/robotframework-excelsage/) 
เพื่อ:
- ✅ อ่านข้อมูลจากไฟล์ Excel
- ✅ ตรวจสอบผล "ผ่าน/ไม่ผ่าน"
- ✅ คำนวณเปอร์เซ็นต์จากคะแนน
- ✅ เขียนผลลัพธ์กลับลง Excel แบบอัตโนมัติ

---

## 📌 ความสามารถของสคริปต์นี้

| ความสามารถ | รายละเอียด |
|-------------|------------|
| ✅ อ่านชื่อและคะแนนจาก Excel | จากคอลัมน์ A (Name) และ B (Score) |
| ✅ เช็กผลสอบ | เขียนคำว่า “ผ่าน” หรือ “ไม่ผ่าน” ลงคอลัมน์ C |
| ✅ คำนวณเปอร์เซ็นต์ | ใส่สูตร `=B2/100` ลงคอลัมน์ D |
| ✅ บันทึกผลกลับ | ข้อมูลทั้งหมดถูกเขียนกลับลงไฟล์ Excel |

---

## 📦 วิธีติดตั้ง

1. ติดตั้ง Python 3.7+
2. ติดตั้ง Robot Framework และ ExcelSage:

```bash
pip3 install robotframework
pip3 install robotframework-excelsage

```

[Keywords ExcelSage](https://deekshith-poojary98.github.io/robotframework-excelsage/) 
[Credit by deekshith-poojary98](https://github.com/deekshith-poojary98/robotframework-excelsage)