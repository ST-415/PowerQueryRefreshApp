"""
Excel Refresher
คลาสสำหรับรีเฟช Power Query ใน Excel
"""

import time
from typing import Dict, List, Optional, Any
from ..core.logger_manager import LoggerManager
from ..core.file_manager import FileManager

try:
    import xlwings as xw
except ImportError:
    print("กรุณาติดตั้ง xlwings: pip install xlwings")
    xw = None


class ExcelRefresher:
    """คลาสสำหรับรีเฟช Excel Power Query"""
    
    def __init__(self, logger: LoggerManager, file_manager: FileManager):
        """
        เริ่มต้น ExcelRefresher
        
        Args:
            logger (LoggerManager): ตัวจัดการ logging
            file_manager (FileManager): ตัวจัดการไฟล์
        """
        self.logger = logger
        self.file_manager = file_manager
        self.app = None
        self.workbook = None
        
        if xw is None:
            self.logger.error("xlwings ไม่พร้อมใช้งาน กรุณาติดตั้ง: pip install xlwings")
    
    def _check_dependencies(self) -> bool:
        """
        ตรวจสอบ dependencies
        
        Returns:
            bool: True หากพร้อมใช้งาน
        """
        return xw is not None
    
    def _open_excel_app(self, visible: bool = False) -> bool:
        """
        เปิดแอปพลิเคชัน Excel
        
        Args:
            visible (bool): แสดง Excel หรือไม่
            
        Returns:
            bool: True หากเปิดสำเร็จ
        """
        try:
            self.app = xw.App(visible=visible)
            self.logger.info(f"เปิด Excel Application (visible={visible})")
            return True
        except Exception as e:
            self.logger.error(f"ไม่สามารถเปิด Excel Application: {e}")
            return False
    
    def _close_excel_app(self) -> None:
        """ปิดแอปพลิเคชัน Excel"""
        try:
            if self.workbook:
                self.workbook.close()
                self.workbook = None
            if self.app:
                self.app.quit()
                self.app = None
            self.logger.info("ปิด Excel Application")
        except Exception as e:
            self.logger.error(f"เกิดข้อผิดพลาดในการปิด Excel: {e}")
    
    def _open_workbook(self, file_path: str) -> bool:
        """
        เปิด workbook
        
        Args:
            file_path (str): เส้นทางไฟล์
            
        Returns:
            bool: True หากเปิดสำเร็จ
        """
        try:
            self.workbook = self.app.books.open(file_path)
            self.logger.info(f"เปิดไฟล์: {file_path}")
            return True
        except Exception as e:
            self.logger.error(f"ไม่สามารถเปิดไฟล์ {file_path}: {e}")
            return False
    
    def _refresh_connections(self, timeout_seconds: int = 1800) -> bool:
        """
        รีเฟชการเชื่อมต่อทั้งหมดใน workbook
        
        Args:
            timeout_seconds (int): timeout ในหน่วยวินาที
            
        Returns:
            bool: True หากรีเฟชสำเร็จ
        """
        try:
            connections = self.workbook.api.Connections
            
            if connections.Count == 0:
                self.logger.info("ไม่มีการเชื่อมต่อที่ต้องรีเฟช")
                return True
            
            self.logger.info(f"พบการเชื่อมต่อ {connections.Count} รายการ")
            
            # รีเฟชทุกการเชื่อมต่อ
            for i, connection in enumerate(connections, 1):
                self.logger.info(f"รีเฟชการเชื่อมต่อ {i}/{connections.Count}: {connection.Name}")
                connection.Refresh()
            
            # รอให้การรีเฟชเสร็จสิ้น
            return self._wait_for_refresh_completion(timeout_seconds)
            
        except Exception as e:
            self.logger.error(f"เกิดข้อผิดพลาดในการรีเฟช: {e}")
            return False
    
    def _wait_for_refresh_completion(self, timeout_seconds: int) -> bool:
        """
        รอให้การรีเฟชเสร็จสิ้น
        
        Args:
            timeout_seconds (int): timeout ในหน่วยวินาที
            
        Returns:
            bool: True หากรีเฟชเสร็จสิ้น
        """
        start_time = time.time()
        
        while time.time() - start_time < timeout_seconds:
            refreshing = False
            
            try:
                for connection in self.workbook.api.Connections:
                    # ตรวจสอบสถานะการรีเฟช
                    if hasattr(connection, 'OLEDBConnection') and connection.OLEDBConnection:
                        if connection.OLEDBConnection.Refreshing:
                            refreshing = True
                            break
                    elif hasattr(connection, 'WorkbookConnection') and connection.WorkbookConnection:
                        # สำหรับการเชื่อมต่อประเภทอื่น ๆ
                        pass
                
                if not refreshing:
                    self.logger.info("การรีเฟชเสร็จสิ้น")
                    return True
                    
            except Exception as e:
                self.logger.error(f"เกิดข้อผิดพลาดในการตรวจสอบสถานะรีเฟช: {e}")
                break
            
            time.sleep(1)
        
        self.logger.error(f"การรีเฟชใช้เวลานานเกิน {timeout_seconds} วินาที")
        return False
    
    def _save_workbook(self, auto_save: bool = True) -> bool:
        """
        บันทึก workbook
        
        Args:
            auto_save (bool): บันทึกอัตโนมัติหรือไม่
            
        Returns:
            bool: True หากบันทึกสำเร็จ
        """
        if not auto_save:
            return True
            
        try:
            self.workbook.save()
            self.logger.info("บันทึกไฟล์สำเร็จ")
            return True
        except Exception as e:
            self.logger.error(f"ไม่สามารถบันทึกไฟล์: {e}")
            return False
    
    def refresh_file(self, file_info: Dict[str, Any], settings: Dict[str, Any]) -> bool:
        """
        รีเฟชไฟล์ Excel
        
        Args:
            file_info (Dict[str, Any]): ข้อมูลไฟล์
            settings (Dict[str, Any]): การตั้งค่า
            
        Returns:
            bool: True หากรีเฟชสำเร็จ
        """
        if not self._check_dependencies():
            return False
        
        file_path = file_info["path"]
        
        # ตรวจสอบไฟล์
        if not self.file_manager.file_exists(file_path):
            self.logger.error(f"ไม่พบไฟล์ Excel: {file_path}")
            return False
        
        if not self.file_manager.is_excel_file(file_path):
            self.logger.error(f"ไฟล์ไม่ใช่ Excel: {file_path}")
            return False
        
        self.logger.info(f"เริ่มรีเฟช Excel: {file_info['name']} - {file_path}")
        
        # สำรองไฟล์
        backup_path = self.file_manager.backup_file(
            file_path, 
            settings.get("backup_before_refresh", False)
        )
        if backup_path:
            self.logger.info(f"สำรองไฟล์: {backup_path}")
        
        success = False
        try:
            # เปิด Excel
            if not self._open_excel_app(visible=False):
                return False
            
            # เปิดไฟล์
            if not self._open_workbook(file_path):
                return False
            
            # รีเฟชการเชื่อมต่อ
            timeout_minutes = settings.get("refresh_timeout_minutes", 30)
            timeout_seconds = timeout_minutes * 60
            
            if not self._refresh_connections(timeout_seconds):
                return False
            
            # บันทึกไฟล์
            if not self._save_workbook(settings.get("auto_save", True)):
                return False
            
            success = True
            self.logger.info(f"รีเฟช Excel เสร็จสิ้น: {file_info['name']}")
            
        except Exception as e:
            self.logger.error(f"เกิดข้อผิดพลาดในการรีเฟช Excel {file_path}: {e}")
            success = False
        
        finally:
            # ปิด Excel
            self._close_excel_app()
        
        return success
    
    def refresh_multiple_files(self, files: List[Dict[str, Any]], settings: Dict[str, Any]) -> Dict[str, int]:
        """
        รีเฟชไฟล์ Excel หลายไฟล์
        
        Args:
            files (List[Dict[str, Any]]): รายการไฟล์
            settings (Dict[str, Any]): การตั้งค่า
            
        Returns:
            Dict[str, int]: ผลลัพธ์การรีเฟช
        """
        if not files:
            self.logger.info("ไม่มีไฟล์ Excel ที่จะรีเฟช")
            return {"success": 0, "failed": 0, "total": 0}
        
        self.logger.info(f"เริ่มรีเฟช Excel {len(files)} ไฟล์")
        
        success_count = 0
        failed_count = 0
        
        for file_info in files:
            if self.refresh_file(file_info, settings):
                success_count += 1
            else:
                failed_count += 1
        
        result = {
            "success": success_count,
            "failed": failed_count,
            "total": len(files)
        }
        
        self.logger.info(f"รีเฟช Excel เสร็จสิ้น: {success_count}/{len(files)} ไฟล์")
        
        return result
