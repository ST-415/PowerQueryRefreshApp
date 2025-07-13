"""
Configuration Manager
จัดการการโหลดและจัดเก็บการตั้งค่าจากไฟล์ JSON
"""

import json
import os
from typing import Dict, List, Any, Optional


class ConfigManager:
    """คลาสสำหรับจัดการการตั้งค่า"""
    
    def __init__(self, config_path: str = "config/config.json"):
        """
        เริ่มต้น ConfigManager
        
        Args:
            config_path (str): เส้นทางไฟล์การตั้งค่า
        """
        self.config_path = config_path
        self._config = self.load_config()
    
    def load_config(self) -> Dict[str, Any]:
        """
        โหลดการตั้งค่าจากไฟล์ JSON
        
        Returns:
            Dict[str, Any]: ข้อมูลการตั้งค่า
        """
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                self._validate_config(config)
                return config
        except FileNotFoundError:
            print(f"ไม่พบไฟล์การตั้งค่า: {self.config_path}")
            return self._get_default_config()
        except json.JSONDecodeError as e:
            print(f"ไฟล์การตั้งค่า JSON มีรูปแบบผิด: {self.config_path}")
            print(f"Error: {e}")
            return self._get_default_config()
    
    def _validate_config(self, config: Dict[str, Any]) -> None:
        """
        ตรวจสอบความถูกต้องของการตั้งค่า
        
        Args:
            config (Dict[str, Any]): ข้อมูลการตั้งค่า
        """
        required_sections = ["excel_files", "settings"]
        for section in required_sections:
            if section not in config:
                print(f"Warning: ไม่พบส่วน '{section}' ในไฟล์การตั้งค่า")
    
    def _get_default_config(self) -> Dict[str, Any]:
        """
        สร้างการตั้งค่าเริ่มต้น
        
        Returns:
            Dict[str, Any]: การตั้งค่าเริ่มต้น
        """
        return {
            "excel_files": [],
            "settings": {
                "auto_save": True,
                "backup_before_refresh": True,
                "log_refresh_activity": True,
                "refresh_timeout_minutes": 30
            }
        }
    
    def save_config(self) -> bool:
        """
        บันทึกการตั้งค่าลงไฟล์
        
        Returns:
            bool: True หากบันทึกสำเร็จ
        """
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"ไม่สามารถบันทึกไฟล์การตั้งค่า: {e}")
            return False
    
    @property
    def excel_files(self) -> List[Dict[str, str]]:
        """รายการไฟล์ Excel พร้อมชื่อไฟล์อัตโนมัติ"""
        files = self._config.get("excel_files", [])
        # เพิ่มชื่อไฟล์จาก path อัตโนมัติ
        for file in files:
            if 'name' not in file and 'path' in file:
                file['name'] = os.path.splitext(os.path.basename(file['path']))[0]
        return files
    
    @property
    def settings(self) -> Dict[str, Any]:
        """การตั้งค่าทั่วไป"""
        return self._config.get("settings", {})
    
    def get_setting(self, key: str, default: Any = None) -> Any:
        """
        ดึงค่าการตั้งค่าเฉพาะ
        
        Args:
            key (str): คีย์การตั้งค่า
            default (Any): ค่าเริ่มต้นหากไม่พบ
            
        Returns:
            Any: ค่าการตั้งค่า
        """
        return self.settings.get(key, default)
    
    def add_excel_file(self, path: str) -> None:
        """
        เพิ่มไฟล์ Excel
        
        Args:
            path (str): เส้นทางไฟล์
        """
        file_info = {
            "path": path
        }
        self._config["excel_files"].append(file_info)
