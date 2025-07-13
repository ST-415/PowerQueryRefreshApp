# PowerQuery Refresh Application

แอปพลิเคชันสำหรับรีเฟช Power Query ใน Excel โดยอัตโนมัติ พร้อม GUI ที่ใช้งานง่าย

## 🎨 **New Dark Theme GUI**

### **Main Window Features**:
- **VS Code Dark Theme**: Professional dark interface matching VS Code
- **Compact Size**: 750x500 pixels, fixed size for focused workflow
- **File Management**: Check file status with visual indicators (✓ Ready / ✗ Missing)
- **Selective Refresh**: Choose which files to refresh with checkboxes
- **Progress Tracking**: Real-time progress bar and status updates
- **English Interface**: Clean, professional English text

### **Settings Window Features**:
- **File Management Tab**: Add, edit, remove Excel files easily
- **General Settings Tab**: Configure auto-save, backup, logging, and timeout
- **Detailed Descriptions**: Built-in help text for all options
- **Compact Design**: 700x550 pixels, modal window design

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
