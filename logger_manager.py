"""
Logger Manager
จัดการระบบ logging สำหรับโปรแกรม
"""

import logging
import os
from datetime import datetime
from typing import Optional


class LoggerManager:
    """คลาสสำหรับจัดการ logging"""
    
    def __init__(self, log_dir: str = "logs", log_level: int = logging.INFO):
        """
        เริ่มต้น LoggerManager
        
        Args:
            log_dir (str): โฟลเดอร์สำหรับเก็บ log files
            log_level (int): ระดับการ logging
        """
        self.log_dir = log_dir
        self.log_level = log_level
        self.logger = self._setup_logger()
    
    def _setup_logger(self) -> logging.Logger:
        """
        ตั้งค่า logger
        
        Returns:
            logging.Logger: logger object
        """
        # สร้างโฟลเดอร์ log หากยังไม่มี
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
        
        # สร้างชื่อไฟล์ log
        log_filename = f"refresh_log_{datetime.now().strftime('%Y%m%d')}.log"
        log_filepath = os.path.join(self.log_dir, log_filename)
        
        # ตั้งค่า logging format
        log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        date_format = '%Y-%m-%d %H:%M:%S'
        
        # สร้าง logger
        logger = logging.getLogger('PowerQueryRefresher')
        logger.setLevel(self.log_level)
        
        # ล้าง handlers เก่า (ถ้ามี)
        if logger.handlers:
            logger.handlers.clear()
        
        # สร้าง file handler
        file_handler = logging.FileHandler(log_filepath, encoding='utf-8')
        file_handler.setLevel(self.log_level)
        file_formatter = logging.Formatter(log_format, date_format)
        file_handler.setFormatter(file_formatter)
        
        # สร้าง console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(self.log_level)
        console_formatter = logging.Formatter(log_format, date_format)
        console_handler.setFormatter(console_formatter)
        
        # เพิ่ม handlers ใน logger
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        return logger
    
    def get_logger(self) -> logging.Logger:
        """
        ดึง logger object
        
        Returns:
            logging.Logger: logger object
        """
        return self.logger
    
    def info(self, message: str) -> None:
        """
        Log ข้อความระดับ INFO
        
        Args:
            message (str): ข้อความ
        """
        self.logger.info(message)
    
    def warning(self, message: str) -> None:
        """
        Log ข้อความระดับ WARNING
        
        Args:
            message (str): ข้อความ
        """
        self.logger.warning(message)
    
    def error(self, message: str) -> None:
        """
        Log ข้อความระดับ ERROR
        
        Args:
            message (str): ข้อความ
        """
        self.logger.error(message)
    
    def debug(self, message: str) -> None:
        """
        Log ข้อความระดับ DEBUG
        
        Args:
            message (str): ข้อความ
        """
        self.logger.debug(message)
    
    def critical(self, message: str) -> None:
        """
        Log ข้อความระดับ CRITICAL
        
        Args:
            message (str): ข้อความ
        """
        self.logger.critical(message)
