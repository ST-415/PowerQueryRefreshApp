# PowerQuery Refresh Application

แอปพลิเคชันสำหรับรีเฟช Power Query ใน Excel โดยอัตโนมัติ

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
│   └── refreshers/              # โมดูลรีเฟช
│       ├── __init__.py
│       └── excel_refresher.py   # รีเฟช Excel Power Query
├── config/                      # การตั้งค่า
│   └── config.json             # ไฟล์การตั้งค่าหลัก
├── data/                        # ข้อมูลและไฟล์ทำงาน
│   ├── backups/                # ไฟล์สำรอง
│   ├── logs/                   # ไฟล์ log
│   └── test/                   # ไฟล์ทดสอบ
├── run.py                      # จุดเริ่มต้นการทำงาน
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

เรียกใช้แอปพลิเคชัน:
```bash
python run.py
```

หรือเรียกใช้โดยตรงจาก src:
```bash
python src/main.py
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
