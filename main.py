"""
Main Application
โปรแกรมหลักสำหรับรีเฟช Power Query
"""

from typing import Dict, List, Optional, Any
from config_manager import ConfigManager
from logger_manager import LoggerManager
from file_manager import FileManager
from excel_refresher import ExcelRefresher
import time
import os

# เปลี่ยน working directory ไปที่โฟลเดอร์ที่เก็บไฟล์ script นี้
os.chdir(os.path.dirname(os.path.abspath(__file__)))


class PowerQueryRefreshApp:
    """แอปพลิเคชันหลักสำหรับรีเฟช Power Query"""
    
    def __init__(self, config_path: str = "config.json"):
        """
        เริ่มต้นแอปพลิเคชัน
        
        Args:
            config_path (str): เส้นทางไฟล์การตั้งค่า
        """
        # สร้าง managers
        self.config_manager = ConfigManager(config_path)
        self.logger_manager = LoggerManager()
        self.file_manager = FileManager()
        
        # สร้าง refreshers
        self.excel_refresher = ExcelRefresher(self.logger_manager, self.file_manager)
        
        self.logger = self.logger_manager.get_logger()
    
    def show_menu(self) -> None:
        """แสดงเมนูหลัก (สำหรับ reference เท่านั้น ไม่ใช้งานในโหมดอัตโนมัติ)"""
        print("\nเลือกการทำงาน:")
        print("1. แสดงรายการไฟล์")
        print("2. รีเฟช Excel ทั้งหมด")
        print("3. แสดงไฟล์สำรอง")
        print("4. ลบไฟล์สำรองเก่า")
        print("5. รันอัตโนมัติ")
        print("6. ออกจากโปรแกรม")
    
    def list_files(self) -> None:
        """แสดงรายการไฟล์"""
        print("\n=== รายการไฟล์ที่จะรีเฟช ===")
        
        # แสดง Excel
        excel_files = self.config_manager.excel_files
        if excel_files:
            print("\nไฟล์ Excel:")
            for i, file_info in enumerate(excel_files, 1):
                print(f"  {i}. {file_info['name']}")
                print(f"     Path: {file_info['path']}")
        else:
            print("\nไม่มีไฟล์ Excel ที่ตั้งค่าไว้")
    
    def list_backups(self) -> None:
        """แสดงรายการไฟล์สำรอง"""
        backup_files = self.file_manager.list_backup_files()
        
        if not backup_files:
            print("\nไม่มีไฟล์สำรอง")
            return
        
        print("\n=== รายการไฟล์สำรอง ===")
        for i, backup in enumerate(backup_files, 1):
            print(f"\n{i}. {backup['name']}")
            print(f"   Size: {backup['size'] / 1024:.1f} KB")
            print(f"   Modified: {backup['modified']}")
            if backup['folder'] != 'root':
                print(f"   Folder: {backup['folder']}")
            print(f"   Path: {backup['path']}")
    
    def cleanup_backups(self) -> None:
        """ลบไฟล์สำรองเก่า (ใช้ค่าเริ่มต้น 30 วัน)"""
        days_to_keep = 30
        deleted_count = self.file_manager.cleanup_old_backups(days_to_keep)
        print(f"\nลบไฟล์สำรองแล้ว {deleted_count} ไฟล์")
        self.logger.info(f"ลบไฟล์สำรองแล้ว {deleted_count} ไฟล์")
    
    def refresh_all_excel(self) -> None:
        """รีเฟชไฟล์ Excel ทั้งหมด"""
        result = self.excel_refresher.refresh_multiple_files(
            self.config_manager.excel_files,
            self.config_manager.settings
        )
        
        print(f"\nรีเฟช Excel เสร็จสิ้น:")
        print(f"สำเร็จ: {result['success']}")
        print(f"ล้มเหลว: {result['failed']}")
        print(f"รวม: {result['total']}")
    
    def refresh_all(self) -> None:
        """รีเฟชไฟล์ Excel ทั้งหมด"""
        print("\n=== เริ่มรีเฟช Excel ===")
        
        # รีเฟช Excel
        excel_result = self.excel_refresher.refresh_multiple_files(
            self.config_manager.excel_files,
            self.config_manager.settings
        )
        
        # แสดงผลลัพธ์
        print("\nสรุปผลการรีเฟช:")
        print(f"Excel - สำเร็จ: {excel_result['success']}, ล้มเหลว: {excel_result['failed']}")
        print(f"รวมทั้งหมด: {excel_result['total']} ไฟล์")
    
    def verify_files(self) -> Dict[str, Any]:
        """
        ตรวจสอบไฟล์ทั้งหมดที่ตั้งค่าไว้
        
        Returns:
            Dict[str, Any]: ผลการตรวจสอบไฟล์
        """
        self.logger.info("=== เริ่มตรวจสอบไฟล์ ===")
        
        result = {
            "excel_files": {"valid": [], "invalid": []},
            "total_valid": 0,
            "total_invalid": 0
        }
        
        # ตรวจสอบไฟล์ Excel
        for file_info in self.config_manager.excel_files:
            if self.file_manager.file_exists(file_info["path"]) and self.file_manager.is_excel_file(file_info["path"]):
                result["excel_files"]["valid"].append(file_info)
                self.logger.info(f"✓ Excel: {file_info['name']} - {file_info['path']}")
            else:
                result["excel_files"]["invalid"].append(file_info)
                self.logger.error(f"✗ Excel: {file_info['name']} - {file_info['path']} (ไม่พบหรือไฟล์ไม่ถูกต้อง)")
        
        # คำนวณสรุป
        result["total_valid"] = len(result["excel_files"]["valid"])
        result["total_invalid"] = len(result["excel_files"]["invalid"])
        
        self.logger.info(f"ผลการตรวจสอบ: ไฟล์ถูกต้อง {result['total_valid']}, ไฟล์ไม่ถูกต้อง {result['total_invalid']}")
        
        return result
    
    def create_backups(self) -> Dict[str, Any]:
        """
        สร้างไฟล์สำรองสำหรับไฟล์ที่จะรีเฟช
        
        Returns:
            Dict[str, Any]: ผลการสำรองไฟล์
        """
        self.logger.info("=== เริ่มสร้างไฟล์สำรอง ===")
        
        backup_result = {
            "success": 0,
            "failed": 0,
            "backup_paths": []
        }
        
        # สำรองไฟล์ Excel
        for file_info in self.config_manager.excel_files:
            if self.file_manager.file_exists(file_info["path"]):
                backup_path = self.file_manager.backup_file(file_info["path"], True)
                if backup_path:
                    backup_result["success"] += 1
                    backup_result["backup_paths"].append(backup_path)
                    self.logger.info(f"✓ สำรอง Excel: {file_info['name']} → {backup_path}")
                else:
                    backup_result["failed"] += 1
                    self.logger.error(f"✗ ไม่สามารถสำรอง Excel: {file_info['name']}")
        
        self.logger.info(f"ผลการสำรอง: สำเร็จ {backup_result['success']}, ล้มเหลว {backup_result['failed']}")
        
        return backup_result
    
    def auto_cleanup_backups(self, days_to_keep: int = 30) -> int:
        """
        ลบไฟล์สำรองเก่าแบบอัตโนมัติ
        
        Args:
            days_to_keep (int): จำนวนวันที่จะเก็บไฟล์สำรอง
            
        Returns:
            int: จำนวนไฟล์ที่ถูกลบ
        """
        self.logger.info(f"=== เริ่มลบไฟล์สำรองเก่า (เก็บไว้ {days_to_keep} วัน) ===")
        
        deleted_count = self.file_manager.cleanup_old_backups(days_to_keep)
        
        self.logger.info(f"ลบไฟล์สำรองแล้ว {deleted_count} ไฟล์")
        
        return deleted_count

    def run(self) -> None:
        """เริ่มการทำงานแอปพลิเคชันแบบอัตโนมัติ"""
        self.logger.info("=== โปรแกรมรีเฟช Excel เริ่มทำงาน ===")
        print("=== โปรแกรมรีเฟช Excel ===")
        print("โปรแกรมจะทำงานแบบอัตโนมัติ: ตรวจสอบไฟล์ → แบ็กอัพ → ลบไฟล์เก่า → รีเฟช Excel")
        
        # รันกระบวนการอัตโนมัติทันที
        self.run_auto_refresh()
        
        print("\nโปรแกรมทำงานเสร็จสิ้น")
        self.logger.info("=== โปรแกรมจบการทำงาน ===")

    def run_auto_refresh(self) -> None:
        """รันกระบวนการรีเฟชอัตโนมัติ"""
        self.logger.info("=== เริ่มการรีเฟชอัตโนมัติ ===")
        
        try:
            # ขั้นตอนที่ 1: ตรวจสอบไฟล์
            print("\n1. ตรวจสอบไฟล์ที่ตั้งค่าไว้...")
            verification_result = self.verify_files()
            
            if verification_result["total_invalid"] > 0:
                self.logger.warning(f"พบไฟล์ไม่ถูกต้อง {verification_result['total_invalid']} ไฟล์ แต่จะดำเนินการต่อ")
            
            # ขั้นตอนที่ 2: สร้างไฟล์สำรอง
            print("\n2. สร้างไฟล์สำรอง...")
            backup_result = self.create_backups()
            
            if backup_result["failed"] > 0:
                self.logger.warning(f"มีไฟล์ที่ไม่สามารถสำรองได้ {backup_result['failed']} ไฟล์")
            
            # ขั้นตอนที่ 3: ลบไฟล์สำรองเก่า
            print("\n3. ลบไฟล์สำรองเก่า...")
            deleted_count = self.auto_cleanup_backups(30)
            
            # ขั้นตอนที่ 4: รีเฟชไฟล์ Excel
            print("\n4. เริ่มรีเฟชไฟล์ Excel...")
            
            # รีเฟช Excel
            self.logger.info("=== เริ่มรีเฟช Excel ===")
            excel_result = self.excel_refresher.refresh_multiple_files(
                self.config_manager.excel_files,
                self.config_manager.settings
            )
            
            # สรุปผลลัพธ์
            print("\n=== สรุปผลการรีเฟชอัตโนมัติ ===")
            print(f"ไฟล์ที่ตรวจสอบ: ถูกต้อง {verification_result['total_valid']}, ไม่ถูกต้อง {verification_result['total_invalid']}")
            print(f"ไฟล์สำรอง: สำเร็จ {backup_result['success']}, ล้มเหลว {backup_result['failed']}")
            print(f"ไฟล์สำรองเก่าที่ลบ: {deleted_count} ไฟล์")
            print(f"Excel - สำเร็จ: {excel_result['success']}, ล้มเหลว: {excel_result['failed']}")
            print(f"รวมทั้งหมด: {excel_result['total']} ไฟล์")
            
            self.logger.info("=== การรีเฟชอัตโนมัติเสร็จสิ้น ===")
            
        except Exception as e:
            self.logger.error(f"เกิดข้อผิดพลาดในการรีเฟชอัตโนมัติ: {e}")
            print(f"เกิดข้อผิดพลาด: {e}")


if __name__ == "__main__":
    app = PowerQueryRefreshApp()
    app.run()
