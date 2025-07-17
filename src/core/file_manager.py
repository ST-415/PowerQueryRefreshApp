"""
File Manager
จัดการการสำรองไฟล์และการตรวจสอบไฟล์
"""

import os
import shutil
from datetime import datetime, timedelta
from typing import Optional
from pathlib import Path


class FileManager:
    """คลาสสำหรับจัดการไฟล์"""
    
    def __init__(self, backup_dir: str = "data/backups"):
        """
        เริ่มต้น FileManager
        
        Args:
            backup_dir (str): โฟลเดอร์สำหรับสำรองไฟล์
        """
        self.backup_dir = backup_dir
        self._ensure_backup_dir()
    
    def _ensure_backup_dir(self) -> None:
        """สร้างโฟลเดอร์สำรองหากยังไม่มี"""
        if not os.path.exists(self.backup_dir):
            os.makedirs(self.backup_dir)
    
    def file_exists(self, file_path: str) -> bool:
        """
        ตรวจสอบว่าไฟล์มีอยู่หรือไม่
        
        Args:
            file_path (str): เส้นทางไฟล์
            
        Returns:
            bool: True หากไฟล์มีอยู่
        """
        return os.path.exists(file_path) and os.path.isfile(file_path)
    
    def get_file_size(self, file_path: str) -> int:
        """
        ดึงขนาดไฟล์
        
        Args:
            file_path (str): เส้นทางไฟล์
            
        Returns:
            int: ขนาดไฟล์เป็น bytes
        """
        if self.file_exists(file_path):
            return os.path.getsize(file_path)
        return 0
    
    def get_file_modified_time(self, file_path: str) -> Optional[datetime]:
        """
        ดึงวันที่แก้ไขไฟล์ล่าสุด
        
        Args:
            file_path (str): เส้นทางไฟล์
            
        Returns:
            Optional[datetime]: วันที่แก้ไข หรือ None หากไฟล์ไม่มี
        """
        if self.file_exists(file_path):
            timestamp = os.path.getmtime(file_path)
            return datetime.fromtimestamp(timestamp)
        return None
    
    def backup_file(self, file_path: str, enable_backup: bool = True) -> Optional[str]:
        """
        สำรองไฟล์ในโฟลเดอร์ตามชื่อไฟล์
        
        Args:
            file_path (str): เส้นทางไฟล์ที่จะสำรอง
            enable_backup (bool): เปิดใช้งานการสำรองหรือไม่
            
        Returns:
            Optional[str]: เส้นทางไฟล์สำรอง หรือ None หากไม่สำรอง
        """
        if not enable_backup:
            return None
            
        if not self.file_exists(file_path):
            return None
        
        # ดึงชื่อไฟล์โดยไม่มีนามสกุล
        file_name = os.path.basename(file_path)
        file_name_without_ext = os.path.splitext(file_name)[0]
        
        # สร้างโฟลเดอร์ย่อยตามชื่อไฟล์
        file_backup_dir = os.path.join(self.backup_dir, file_name_without_ext)
        if not os.path.exists(file_backup_dir):
            os.makedirs(file_backup_dir)
        
        # สร้างชื่อไฟล์สำรอง
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"{timestamp}_{file_name}"
        backup_path = os.path.join(file_backup_dir, backup_filename)
        
        try:
            shutil.copy2(file_path, backup_path)
            return backup_path
        except Exception as e:
            print(f"ไม่สามารถสำรองไฟล์ {file_path}: {e}")
            return None
    
    def get_absolute_path(self, file_path: str) -> str:
        """
        แปลงเส้นทางเป็น absolute path
        
        Args:
            file_path (str): เส้นทางไฟล์
            
        Returns:
            str: เส้นทางแบบ absolute
        """
        return os.path.abspath(file_path)
    
    def get_file_extension(self, file_path: str) -> str:
        """
        ดึงนามสกุลไฟล์
        
        Args:
            file_path (str): เส้นทางไฟล์
            
        Returns:
            str: นามสกุลไฟล์
        """
        return Path(file_path).suffix.lower()
    
    def is_excel_file(self, file_path: str) -> bool:
        """
        ตรวจสอบว่าเป็นไฟล์ Excel หรือไม่
        
        Args:
            file_path (str): เส้นทางไฟล์
            
        Returns:
            bool: True หากเป็นไฟล์ Excel
        """
        excel_extensions = ['.xlsx', '.xlsm', '.xls']
        return self.get_file_extension(file_path) in excel_extensions
    
    def is_powerbi_file(self, file_path: str) -> bool:
        """
        ตรวจสอบว่าเป็นไฟล์ Power BI หรือไม่
        
        Args:
            file_path (str): เส้นทางไฟล์
            
        Returns:
            bool: True หากเป็นไฟล์ Power BI
        """
        powerbi_extensions = ['.pbix', '.pbit']
        return self.get_file_extension(file_path) in powerbi_extensions
    
    def list_backup_files(self) -> list:
        """
        แสดงรายการไฟล์สำรองทั้งหมด (รวมในโฟลเดอร์ย่อย)
        
        Returns:
            list: รายการไฟล์สำรอง
        """
        backup_files = []
        if os.path.exists(self.backup_dir):
            # เรียกดูไฟล์ในโฟลเดอร์หลัก
            for item in os.listdir(self.backup_dir):
                item_path = os.path.join(self.backup_dir, item)
                
                if os.path.isfile(item_path):
                    # ไฟล์ในโฟลเดอร์หลัก
                    backup_files.append({
                        'name': item,
                        'path': item_path,
                        'size': self.get_file_size(item_path),
                        'modified': self.get_file_modified_time(item_path),
                        'folder': 'root'
                    })
                elif os.path.isdir(item_path):
                    # โฟลเดอร์ย่อย - เรียกดูไฟล์ข้างใน
                    for file in os.listdir(item_path):
                        file_path = os.path.join(item_path, file)
                        if os.path.isfile(file_path):
                            backup_files.append({
                                'name': file,
                                'path': file_path,
                                'size': self.get_file_size(file_path),
                                'modified': self.get_file_modified_time(file_path),
                                'folder': item
                            })
        
        return sorted(backup_files, key=lambda x: x['modified'], reverse=True)
    
    def cleanup_old_backups(self, days_to_keep: int = 30) -> int:
        """
        ลบไฟล์สำรองเก่า (รวมในโฟลเดอร์ย่อย)
        
        Args:
            days_to_keep (int): จำนวนวันที่จะเก็บไฟล์สำรอง
            
        Returns:
            int: จำนวนไฟล์ที่ถูกลบ
        """
        deleted_count = 0
        cutoff_date = datetime.now() - timedelta(days=days_to_keep)
        
        backup_files = self.list_backup_files()
        for backup_file in backup_files:
            if backup_file['modified'] and backup_file['modified'] < cutoff_date:
                try:
                    os.remove(backup_file['path'])
                    deleted_count += 1
                    
                    # ตรวจสอบว่าโฟลเดอร์ย่อยว่างหรือไม่ หากว่างให้ลบ
                    if backup_file['folder'] != 'root':
                        folder_path = os.path.dirname(backup_file['path'])
                        if os.path.exists(folder_path) and not os.listdir(folder_path):
                            try:
                                os.rmdir(folder_path)
                            except Exception:
                                pass  # ไม่สำคัญหากลบโฟลเดอร์ไม่ได้
                                
                except Exception as e:
                    print(f"ไม่สามารถลบไฟล์สำรอง {backup_file['name']}: {e}")
        
        return deleted_count
    
    def show_backup_structure(self) -> None:
        """แสดงโครงสร้างโฟลเดอร์สำรองแบบเป็นระเบียบ"""
        if not os.path.exists(self.backup_dir):
            print("ไม่มีโฟลเดอร์สำรอง")
            return
        
        print(f"\n=== โครงสร้างโฟลเดอร์สำรอง: {self.backup_dir} ===")
        
        # เก็บข้อมูลสำหรับการนับ
        folder_stats = {}
        total_files = 0
        total_size = 0
        
        for item in os.listdir(self.backup_dir):
            item_path = os.path.join(self.backup_dir, item)
            
            if os.path.isfile(item_path):
                # ไฟล์ในโฟลเดอร์หลัก (ไฟล์เก่า)
                size = self.get_file_size(item_path)
                modified = self.get_file_modified_time(item_path)
                print(f"📄 {item} ({size/1024:.1f} KB) - {modified.strftime('%Y-%m-%d %H:%M:%S')}")
                total_files += 1
                total_size += size
                
            elif os.path.isdir(item_path):
                # โฟลเดอร์ย่อย
                files_in_folder = []
                folder_size = 0
                
                for file in os.listdir(item_path):
                    file_path = os.path.join(item_path, file)
                    if os.path.isfile(file_path):
                        size = self.get_file_size(file_path)
                        modified = self.get_file_modified_time(file_path)
                        files_in_folder.append({
                            'name': file,
                            'size': size,
                            'modified': modified
                        })
                        folder_size += size
                        total_files += 1
                        total_size += size
                
                # เรียงตามวันที่แก้ไข
                files_in_folder.sort(key=lambda x: x['modified'], reverse=True)
                
                print(f"\n📁 {item}/ ({len(files_in_folder)} ไฟล์, {folder_size/1024:.1f} KB)")
                for file_info in files_in_folder:
                    print(f"   📄 {file_info['name']} ({file_info['size']/1024:.1f} KB) - {file_info['modified'].strftime('%Y-%m-%d %H:%M:%S')}")
                
                folder_stats[item] = {
                    'files': len(files_in_folder),
                    'size': folder_size
                }
        
        # สรุป
        print(f"\n=== สรุป ===")
        print(f"โฟลเดอร์ย่อย: {len(folder_stats)} โฟลเดอร์")
        print(f"ไฟล์ทั้งหมด: {total_files} ไฟล์")
        print(f"ขนาดรวม: {total_size/1024:.1f} KB ({total_size/(1024*1024):.2f} MB)")
    
    def get_excel_files(self, data_folder: str = "data") -> list:
        """
        ดึงรายการไฟล์ Excel จากโฟลเดอร์ข้อมูล
        
        Args:
            data_folder (str): โฟลเดอร์ที่จะค้นหาไฟล์ Excel
            
        Returns:
            list: รายการเส้นทางไฟล์ Excel
        """
        excel_files = []
        
        # ตรวจสอบว่าโฟลเดอร์มีอยู่
        if not os.path.exists(data_folder):
            return excel_files
        
        # ค้นหาไฟล์ Excel ในโฟลเดอร์
        for item in os.listdir(data_folder):
            item_path = os.path.join(data_folder, item)
            
            # ตรวจสอบว่าเป็นไฟล์และเป็นไฟล์ Excel
            if os.path.isfile(item_path) and self.is_excel_file(item_path):
                excel_files.append(self.get_absolute_path(item_path))
        
        # เรียงตามชื่อไฟล์
        excel_files.sort()
        return excel_files
