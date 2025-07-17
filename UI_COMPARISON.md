"""
UI Comparison and Migration Guide
คู่มือเปรียบเทียบ UI เก่าและใหม่
"""

# 🎨 UI Comparison: tkinter vs CustomTkinter

## Overview
แอปพลิเคชันนี้มี GUI 2 แบบให้เลือกใช้:

### 1. GUI เก่า (tkinter)
- **ไฟล์**: `src/gui/main_gui.py`
- **เรียกใช้**: `python run_gui.py`
- **ขนาด**: 750x500 pixels (ขนาดคงที่)
- **ธีม**: Dark theme คล้าย VS Code

### 2. GUI ใหม่ (CustomTkinter) - 🌟 แนะนำ!
- **ไฟล์**: `src/gui/modern_main_gui.py`
- **เรียกใช้**: `python run_modern_gui.py`
- **ขนาด**: 900x700 pixels (ปรับขนาดได้)
- **ธีม**: Modern light theme

## Feature Comparison

| Feature | GUI เก่า | GUI ใหม่ |
|---------|----------|----------|
| **Design** | Dark theme | Modern light theme |
| **Layout** | Fixed size | Responsive |
| **File Display** | List view | Card view |
| **Buttons** | Standard | Rounded with icons |
| **Colors** | Dark colors | Modern colors |
| **Animation** | None | Smooth transitions |
| **Settings** | Single window | Tabbed interface |
| **Progress** | Basic bar | Modern progress bar |
| **Icons** | Text-based | Emoji icons |
| **Corners** | Square | Rounded |

## Visual Differences

### Main Window
- **เก่า**: รายการไฟล์แบบธรรมดา
- **ใหม่**: การ์ดไฟล์ที่สวยงามพร้อมรายละเอียด

### Settings Window
- **เก่า**: หน้าต่างเดียว
- **ใหม่**: แบ่งเป็น 3 แท็บ (Folders, Backup, Refresh)

### Color Scheme
- **เก่า**: Dark theme (เทา-ดำ)
- **ใหม่**: Light theme (ขาว-สี)

## Why Choose Modern GUI?

### ข้อดีของ GUI ใหม่:
1. **สวยงามกว่า**: ดีไซน์ทันสมัย เหมือน macOS/Windows 11
2. **ใช้งานง่ายกว่า**: การจัดวางที่ดีกว่า ข้อมูลชัดเจน
3. **ทันสมัยกว่า**: ใช้เทคโนโลยี CustomTkinter ที่ทันสมัย
4. **ปรับขนาดได้**: สามารถปรับขนาดหน้าต่างได้
5. **รองรับอนาคต**: พร้อมสำหรับการพัฒนาต่อ

### ข้อดีของ GUI เก่า:
1. **เสถียร**: ใช้ tkinter ที่เสถียร
2. **ขนาดเล็ก**: ขนาดไฟล์เล็กกว่า
3. **เรียบง่าย**: ไม่ซับซ้อน
4. **Dark Theme**: เหมาะสำหรับคนชอบธีมมืด

## Migration Tips

### สำหรับผู้ใช้ GUI เก่า:
1. ลองใช้ GUI ใหม่ด้วย `python run_modern_gui.py`
2. การตั้งค่าจะถูกแชร์ระหว่าง GUI ทั้งสอง
3. ฟังก์ชันการทำงานเหมือนกัน เพียงแต่ UI สวยกว่า

### สำหรับนักพัฒนา:
1. GUI ใหม่ใช้ CustomTkinter แทน tkinter
2. Code structure คล้ายกัน แต่ใช้ `ctk.` แทน `ttk.`
3. มีการจัดการสีและธีมที่ดีกว่า

## Usage Examples

### เปิด GUI ใหม่:
```bash
# Command line
python run_modern_gui.py

# Batch file
start_modern_gui.bat
```

### เปิด GUI เก่า:
```bash
# Command line
python run_gui.py

# Batch file
start_gui.bat
```

## Future Plans

GUI ใหม่จะเป็นหลักในการพัฒนาต่อ:
- เพิ่มฟีเจอร์ใหม่
- ปรับปรุงประสิทธิภาพ
- เพิ่มการปรับแต่ง theme
- รองรับหลายภาษา

GUI เก่าจะยังคงไว้เพื่อความเข้ากันได้ (backward compatibility)
