"""
Modern GUI Entry Point
จุดเริ่มต้นสำหรับเรียกใช้แอปพลิเคชัน GUI แบบใหม่ด้วย CustomTkinter
"""

import sys
import os

# เพิ่ม src ไปยัง Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.gui.modern_main_gui import main

if __name__ == "__main__":
    main()
