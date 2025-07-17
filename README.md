# PowerQuery Refresh Application

แอปพลิเคชันสำหรับรีเฟช Power Query ใน Excel โดยอัตโนมัติ พร้อม GUI ที่ใช้งานง่าย

## 🎨 **New Modern GUI with CustomTkinter**

### **เปรียบเทียบ GUI เก่าและใหม่**:

#### **GUI เก่า (tkinter)**:
- **Dark Theme**: Professional dark interface matching VS Code
- **Compact Size**: 750x500 pixels, fixed size for focused workflow
- **File Management**: Check file status with visual indicators (✓ Ready / ✗ Missing)
- **Selective Refresh**: Choose which files to refresh with checkboxes
- **Progress Tracking**: Real-time progress bar and status updates
- **English Interface**: Clean, professional English text

#### **GUI ใหม่ (CustomTkinter) - แนะนำ!**:
- **Modern Design**: สวยงาม ทันสมัย เหมือน macOS/Windows 11
- **Smooth Animation**: การเคลื่อนไหวที่นุ่มนวล
- **Card-based Layout**: แสดงไฟล์ในรูปแบบ Card สวยงาม
- **Rounded Corners**: มุมโค้งมนทั่วทั้งแอป
- **Better Colors**: ใช้สีที่สวยงามและอ่านง่าย
- **Responsive Layout**: ปรับขนาดได้ตามหน้าจอ
- **Modern Icons**: ใช้ Emoji Icons ที่ดูทันสมัย
- **Tabbed Settings**: การตั้งค่าแบ่งเป็นแท็บง่ายต่อการใช้งาน

### **วิธีเรียกใช้ GUI**:
1. **GUI เก่า**: `python run_gui.py` หรือ `start_gui.bat`
2. **GUI ใหม่**: `python run_modern_gui.py` หรือ `start_modern_gui.bat`

### **Features ของ GUI ใหม่**:
- **Beautiful File Cards**: แสดงไฟล์ในรูปแบบ Card พร้อมไอคอน ชื่อไฟล์ ขนาดไฟล์
- **Smart Progress Bar**: แสดงความคืบหน้าที่สวยงาม
- **Modern Settings**: หน้าต่างการตั้งค่าแบ่งเป็น 3 แท็บ
  - 📁 **Folders**: จัดการโฟลเดอร์
  - 💾 **Backup**: ตั้งค่าการสำรอง
  - 🔄 **Refresh**: ตั้งค่าการรีเฟช
- **Better Button Design**: ปุ่มกดสวยงาม มีสีสันเหมาะสม
- **Improved Layout**: การจัดวางที่ดีกว่า ใช้พื้นที่ได้มีประสิทธิภาพ

## โครงสร้างโปรเจกต์

```
PowerQueryRefreshApp/
├── src/                          # ซอร์สโค้ดหลัก
│   ├── __init__.py              # Package initializer
│   ├── main.py                  # แอปพลิเคชันหลัก
│   ├── core/                    # โมดูลหลัก
│   │   ├── __init__.py
│   │   ├── config_manager.py    # จัดการการตั้งค่า
│   │   ├── logger_manager.py    # จัดการ logging
│   │   └── file_manager.py      # จัดการไฟล์และสำรอง
│   ├── refreshers/              # โมดูลรีเฟช
│   │   ├── __init__.py
│   │   └── excel_refresher.py   # รีเฟช Excel Power Query
│   └── gui/                     # โมดูล GUI
│       ├── __init__.py
│       ├── main_gui.py          # หน้าต่างหลัก
│       └── settings_window.py   # หน้าต่างการตั้งค่า
├── config/                      # การตั้งค่า
│   └── config.json             # ไฟล์การตั้งค่าหลัก
├── data/                        # ข้อมูลและไฟล์ทำงาน
│   ├── backups/                # ไฟล์สำรอง
│   ├── logs/                   # ไฟล์ log
│   └── test/                   # ไฟล์ทดสอบ
├── run.py                      # จุดเริ่มต้น (Command Line)
├── run_gui.py                  # จุดเริ่มต้น (GUI)
├── start_gui.bat               # เปิด GUI ด้วย batch file
├── requirements.txt            # Python dependencies
└── README.md                   # เอกสารนี้
```

## 🚀 การติดตั้งและใช้งาน

### ติดตั้ง Dependencies:
```bash
pip install -r requirements.txt
```

### วิธีใช้งาน:

#### 1. **GUI แบบใหม่ (แนะนำ)**:
```bash
python run_modern_gui.py
```
หรือ
```bash
start_modern_gui.bat
```

#### 2. **GUI แบบเก่า**:
```bash
python run_gui.py
```
หรือ
```bash
start_gui.bat
```

#### 3. **Command Line**:
```bash
python run.py
```

## การตั้งค่า

แก้ไขไฟล์ `config/config.json`:
```json
{
  "excel_files": [
    {
      "path": "data/test/ex_data1_query.xlsx"
    },
    {
      "path": "data/test/ex_data2_query.xlsx"
    }
  ],
  "settings": {
    "auto_save": true,
    "backup_before_refresh": true,
    "log_refresh_activity": true,
    "refresh_timeout_minutes": 30
  }
}
```

## คุณสมบัติ

- รีเฟช Power Query ใน Excel อัตโนมัติ
- สำรองไฟล์ก่อนการรีเฟช
- จัดการ log การทำงาน
- ลบไฟล์สำรองเก่าอัตโนมัติ
- โครงสร้างโปรเจกต์แบบ modular

## โมดูลหลัก

### Core Modules
- **ConfigManager**: จัดการการโหลดและบันทึกการตั้งค่า
- **LoggerManager**: จัดการระบบ logging
- **FileManager**: จัดการไฟล์และการสำรอง

### Refreshers
- **ExcelRefresher**: รีเฟช Power Query ใน Excel

## ข้อกำหนด

- Python 3.7+
- Microsoft Excel (สำหรับ xlwings)
- xlwings package
