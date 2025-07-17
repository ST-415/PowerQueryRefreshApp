"""
Installation and Usage Guide
คู่มือการติดตั้งและใช้งาน GUI แบบใหม่
"""

# 🚀 Quick Start Guide - Modern GUI

## 1. ติดตั้ง Dependencies

### ขั้นตอนแรก - ตรวจสอบ Python:
```bash
python --version
```
ต้องเป็น Python 3.7 หรือสูงกว่า

### ติดตั้ง packages:
```bash
pip install -r requirements.txt
```

### Dependencies ที่ติดตั้ง:
- `xlwings>=0.30.0` - สำหรับการทำงานกับ Excel
- `customtkinter>=5.2.0` - สำหรับ Modern GUI
- `pillow>=9.0.0` - สำหรับการจัดการรูปภาพ

## 2. เรียกใช้แอปพลิเคชัน

### 🌟 GUI ใหม่ (แนะนำ):
```bash
python run_modern_gui.py
```
หรือ double-click `start_modern_gui.bat`

### GUI เก่า:
```bash
python run_gui.py
```
หรือ double-click `start_gui.bat`

## 3. การใช้งาน GUI ใหม่

### หน้าหลัก:
1. **File Cards**: แสดงไฟล์ Excel ในรูปแบบการ์ด
2. **Control Buttons**: 
   - ⚙️ Settings - เปิดการตั้งค่า
   - 🔄 Refresh List - อัปเดตรายการไฟล์
   - ✅ Select All - เลือกทั้งหมด
   - ❌ Deselect All - ยกเลิกการเลือกทั้งหมด
3. **Action Buttons**:
   - 🚀 Start Refresh - เริ่มรีเฟช
   - ⏹️ Stop Refresh - หยุดรีเฟช
4. **Progress Bar**: แสดงความคืบหน้า
5. **Status Label**: แสดงสถานะปัจจุบัน

### หน้าการตั้งค่า:
#### 📁 Folders Tab:
- **Data Folder**: โฟลเดอร์ที่เก็บไฟล์ Excel
- **Backup Folder**: โฟลเดอร์สำหรับไฟล์สำรอง
- **Log Folder**: โฟลเดอร์สำหรับไฟล์ log

#### 💾 Backup Tab:
- **Auto Backup**: เปิด/ปิดการสำรองอัตโนมัติ
- **Backup Information**: คำอธิบายเกี่ยวกับการสำรอง

#### 🔄 Refresh Tab:
- **Refresh Timeout**: กำหนดเวลา timeout (วินาที)
- **Performance Tips**: เทคนิคการใช้งาน

## 4. การใช้งานพื้นฐาน

### ขั้นตอนการรีเฟชไฟล์:
1. เปิดแอปพลิเคชัน
2. ตรวจสอบไฟล์ใน File Cards
3. เลือกไฟล์ที่ต้องการรีเฟช (checkbox)
4. กดปุ่ม "🚀 Start Refresh"
5. รอให้กระบวนการเสร็จสิ้น

### เทคนิคการใช้งาน:
- **Select All**: เลือกทุกไฟล์ในคลิกเดียว
- **Progress Tracking**: ดูความคืบหน้าแบบ real-time
- **Auto Backup**: ระบบสำรองไฟล์อัตโนมัติ
- **Error Handling**: ระบบจัดการข้อผิดพลาด

## 5. การตั้งค่าขั้นสูง

### ตั้งค่าโฟลเดอร์:
```json
{
  "data_folder": "data",
  "backup_folder": "data/backups",
  "log_folder": "data/logs"
}
```

### ตั้งค่าการสำรอง:
- เปิด auto backup เพื่อความปลอดภัย
- ไฟล์สำรองจะมี timestamp
- ลบไฟล์สำรองเก่าเป็นระยะ

### ตั้งค่าประสิทธิภาพ:
- เพิ่ม timeout สำหรับไฟล์ขนาดใหญ่
- รีเฟชไฟล์ทีละน้อย
- ปิดแอปพลิเคชัน Excel อื่น

## 6. Troubleshooting

### ปัญหาที่พบบ่อย:

#### แอปไม่เปิด:
```bash
# ตรวจสอบการติดตั้ง
pip list | grep customtkinter

# ติดตั้งใหม่
pip install --upgrade customtkinter
```

#### ไฟล์ไม่แสดง:
- ตรวจสอบ data folder ในการตั้งค่า
- ตรวจสอบไฟล์ .xlsx ใน folder
- กดปุ่ม "🔄 Refresh List"

#### รีเฟชล้มเหลว:
- ตรวจสอบ Excel ว่าปิดแล้ว
- ตรวจสอบ network connection
- เพิ่ม timeout ในการตั้งค่า

### Log Files:
ตรวจสอบ log ใน `data/logs/` folder:
```
data/logs/refresh_log_YYYYMMDD.log
```

## 7. เปรียบเทียบกับ GUI เก่า

| Feature | GUI เก่า | GUI ใหม่ |
|---------|----------|----------|
| ความสวยงาม | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| ความใช้งานง่าย | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| ฟีเจอร์ | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| ประสิทธิภาพ | ⭐⭐⭐⭐ | ⭐⭐⭐⭐ |
| ความเสถียร | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ |

## 8. การพัฒนาต่อ

### สำหรับนักพัฒนา:
- Code อยู่ใน `src/gui/modern_main_gui.py`
- ใช้ CustomTkinter framework
- Modern design patterns
- Responsive layout

### การปรับแต่ง:
- เปลี่ยนสีธีม
- เพิ่มฟีเจอร์ใหม่
- ปรับแต่ง layout
- เพิ่มภาษาอื่น

### การอัปเดต:
```bash
# อัปเดต dependencies
pip install --upgrade -r requirements.txt

# Pull latest changes
git pull origin main
```

---

🎉 **ขอให้สนุกกับการใช้งาน PowerQuery Refresh Tool แบบใหม่!**
