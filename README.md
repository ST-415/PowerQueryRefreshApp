# PowerQuery Refresh Application

แอปพลิเคชันสำหรับรีเฟช Power Query ใน Excel โดยอัตโนมัติ พร้อม GUI ที่ใช้งานง่าย

## คุณสมบัติ

### GUI Application (หน้าต่างกราฟิก)
- **หน้าต่างหลัก**: 
  - ตรวจสอบรายการไฟล์ที่ตั้งค่าไว้
  - เลือกไฟล์ที่ต้องการรีเฟช (ติ๊กเลือกได้)
  - ปุ่มเริ่มการรีเฟชไฟล์ที่เลือก
  - แสดงสถานะการดำเนินการแบบ real-time
  
- **หน้าต่างการตั้งค่า**:
  - เพิ่ม/ลบ/แก้ไขไฟล์ Excel ที่ต้องการรีเฟช
  - ตั้งค่า Auto Save, Backup, Logging
  - กำหนดเวลา timeout สำหรับการรีเฟช

### Command Line Interface
- รีเฟชแบบอัตโนมัติทั้งหมด
- สำรองไฟล์อัตโนมัติก่อนรีเฟช
- ลบไฟล์สำรองเก่าอัตโนมัติ

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

## การติดตั้ง

1. ติดตั้ง Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. ปรับแต่งการตั้งค่าใน `config/config.json`

## การใช้งาน

### GUI Application (แนะนำ)
เรียกใช้ GUI:
```bash
python run_gui.py
```

หรือดับเบิลคลิกที่ `start_gui.bat`

### Command Line Interface
เรียกใช้แบบอัตโนมัติ:
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
