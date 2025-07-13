"""
Entry Point Script
จุดเริ่มต้นสำหรับเรียกใช้แอปพลิเคชัน PowerQuery Refresh
"""

import sys
import os

# เพิ่ม src ไปยัง Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.main import PowerQueryRefreshApp

if __name__ == "__main__":
    app = PowerQueryRefreshApp()
    app.run()
