"""
Main GUI Application
หน้าต่างหลักสำหรับ PowerQuery Refresh Application
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from typing import Dict, List, Any, Optional
import os
import sys

# เพิ่ม path สำหรับ import
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.config_manager import ConfigManager
from core.logger_manager import LoggerManager
from core.file_manager import FileManager
from refreshers.excel_refresher import ExcelRefresher
from gui.settings_window import SettingsWindow


class MainGUI:
    """หน้าต่างหลักของแอปพลิเคชัน"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PowerQuery Refresh Application")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # สร้าง components ตรงๆ แทนการใช้ PowerQueryRefreshApp
        self.config_manager = ConfigManager()
        self.logger_manager = LoggerManager()
        self.file_manager = FileManager()
        self.excel_refresher = ExcelRefresher(self.logger_manager, self.file_manager)
        
        # ตัวแปรสำหรับเก็บสถานะไฟล์
        self.file_vars = []  # เก็บ BooleanVar สำหรับแต่ละไฟล์
        self.file_info = []  # เก็บข้อมูลไฟล์
        
        # สร้าง GUI
        self.create_widgets()
        self.refresh_file_list()
        
        # ตั้งค่า protocol สำหรับปิดหน้าต่าง
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def create_widgets(self):
        """สร้าง widgets ทั้งหมด"""
        # Frame หลัก
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # กำหนด weight สำหรับ responsive
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # หัวข้อ
        title_label = ttk.Label(main_frame, text="PowerQuery Refresh Application", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # ปุ่มการตั้งค่า
        settings_button = ttk.Button(main_frame, text="⚙️ การตั้งค่า", 
                                   command=self.open_settings)
        settings_button.grid(row=1, column=0, sticky=tk.W, pady=(0, 10))
        
        # ปุ่มรีเฟรชรายการไฟล์
        refresh_button = ttk.Button(main_frame, text="🔄 รีเฟรชรายการ", 
                                  command=self.refresh_file_list)
        refresh_button.grid(row=1, column=1, pady=(0, 10))
        
        # สถานะการตรวจสอบ
        self.status_label = ttk.Label(main_frame, text="กำลังโหลดรายการไฟล์...", 
                                     foreground="blue")
        self.status_label.grid(row=1, column=2, sticky=tk.E, pady=(0, 10))
        
        # Frame สำหรับรายการไฟล์
        files_frame = ttk.LabelFrame(main_frame, text="รายการไฟล์ที่ตั้งค่าไว้", padding="10")
        files_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(0, weight=1)
        
        # Treeview สำหรับแสดงรายการไฟล์
        self.create_file_treeview(files_frame)
        
        # Frame สำหรับปุ่มควบคุม
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        control_frame.columnconfigure(0, weight=1)
        control_frame.columnconfigure(1, weight=1)
        control_frame.columnconfigure(2, weight=1)
        
        # ปุ่มเลือกทั้งหมด/ไม่เลือกทั้งหมด
        select_all_button = ttk.Button(control_frame, text="✓ เลือกทั้งหมด", 
                                     command=self.select_all_files)
        select_all_button.grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        deselect_all_button = ttk.Button(control_frame, text="✗ ไม่เลือกทั้งหมด", 
                                       command=self.deselect_all_files)
        deselect_all_button.grid(row=0, column=1, padx=5)
        
        # ปุ่มเริ่มรีเฟช
        self.refresh_button = ttk.Button(control_frame, text="🚀 เริ่มรีเฟชไฟล์ที่เลือก", 
                                       command=self.start_refresh, style='Accent.TButton')
        self.refresh_button.grid(row=0, column=2, sticky=tk.E, padx=(5, 0))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, 
                                          mode='determinate')
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # ข้อความสถานะ
        self.progress_label = ttk.Label(main_frame, text="พร้อมดำเนินการ")
        self.progress_label.grid(row=5, column=0, columnspan=3, pady=(5, 0))
    
    def create_file_treeview(self, parent):
        """สร้าง Treeview สำหรับแสดงรายการไฟล์"""
        # สร้าง Treeview
        columns = ("select", "name", "path", "status")
        self.tree = ttk.Treeview(parent, columns=columns, show="headings", height=10)
        
        # กำหนดหัวคอลัมน์
        self.tree.heading("select", text="เลือก")
        self.tree.heading("name", text="ชื่อไฟล์")
        self.tree.heading("path", text="เส้นทาง")
        self.tree.heading("status", text="สถานะ")
        
        # กำหนดความกว้างคอลัมน์
        self.tree.column("select", width=80, anchor="center")
        self.tree.column("name", width=200)
        self.tree.column("path", width=300)
        self.tree.column("status", width=100, anchor="center")
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # วางใน grid
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Bind double-click event
        self.tree.bind("<Double-1>", self.toggle_file_selection)
    
    def verify_files(self) -> Dict[str, Any]:
        """
        ตรวจสอบไฟล์ทั้งหมดที่ตั้งค่าไว้
        
        Returns:
            Dict[str, Any]: ผลการตรวจสอบไฟล์
        """
        result = {
            "excel_files": {"valid": [], "invalid": []},
            "total_valid": 0,
            "total_invalid": 0
        }
        
        # ตรวจสอบไฟล์ Excel
        for file_info in self.config_manager.excel_files:
            if self.file_manager.file_exists(file_info["path"]) and self.file_manager.is_excel_file(file_info["path"]):
                result["excel_files"]["valid"].append(file_info)
            else:
                result["excel_files"]["invalid"].append(file_info)
        
        # คำนวณสรุป
        result["total_valid"] = len(result["excel_files"]["valid"])
        result["total_invalid"] = len(result["excel_files"]["invalid"])
        
        return result
    
    def create_backups(self) -> Dict[str, Any]:
        """
        สร้างไฟล์สำรองสำหรับไฟล์ที่เลือก
        
        Returns:
            Dict[str, Any]: ผลการสำรองไฟล์
        """
        backup_result = {
            "success": 0,
            "failed": 0,
            "backup_paths": []
        }
        
        # สำรองไฟล์ Excel ที่เลือก
        selected_files = self.get_selected_files()
        for file_info in selected_files:
            if self.file_manager.file_exists(file_info["path"]):
                backup_path = self.file_manager.backup_file(file_info["path"], True)
                if backup_path:
                    backup_result["success"] += 1
                    backup_result["backup_paths"].append(backup_path)
                else:
                    backup_result["failed"] += 1
        
        return backup_result
    
    def auto_cleanup_backups(self, days_to_keep: int = 30) -> int:
        """
        ลบไฟล์สำรองเก่าแบบอัตโนมัติ
        
        Args:
            days_to_keep (int): จำนวนวันที่จะเก็บไฟล์สำรอง
            
        Returns:
            int: จำนวนไฟล์ที่ถูกลบ
        """
        return self.file_manager.cleanup_old_backups(days_to_keep)
    
    def refresh_file_list(self):
        """รีเฟรชรายการไฟล์และตรวจสอบสถานะ"""
        self.status_label.config(text="กำลังตรวจสอบไฟล์...", foreground="blue")
        self.root.update_idletasks()
        
        try:
            # ตรวจสอบไฟล์
            verification_result = self.verify_files()
            
            # ล้างรายการเก่า
            self.tree.delete(*self.tree.get_children())
            self.file_vars.clear()
            self.file_info.clear()
            
            # เพิ่มไฟล์ valid
            for file_info in verification_result["excel_files"]["valid"]:
                var = tk.BooleanVar(value=True)  # เลือกไว้ตั้งแต่แรก
                self.file_vars.append(var)
                self.file_info.append(file_info)
                
                name = os.path.basename(file_info["path"])
                self.tree.insert("", "end", values=(
                    "✓", name, file_info["path"], "✓ พร้อม"
                ))
            
            # เพิ่มไฟล์ invalid
            for file_info in verification_result["excel_files"]["invalid"]:
                var = tk.BooleanVar(value=False)  # ไม่เลือก
                self.file_vars.append(var)
                self.file_info.append(file_info)
                
                name = os.path.basename(file_info["path"])
                self.tree.insert("", "end", values=(
                    "✗", name, file_info["path"], "✗ ไม่พบ"
                ))
            
            # อัพเดทสถานะ
            total_files = len(self.file_info)
            valid_files = verification_result["total_valid"]
            invalid_files = verification_result["total_invalid"]
            
            if invalid_files > 0:
                self.status_label.config(
                    text=f"พบไฟล์ {total_files} ไฟล์ (ใช้ได้ {valid_files}, ไม่พบ {invalid_files})",
                    foreground="orange"
                )
            else:
                self.status_label.config(
                    text=f"พบไฟล์ {valid_files} ไฟล์ (ทั้งหมดพร้อมใช้งาน)",
                    foreground="green"
                )
                
        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถโหลดรายการไฟล์ได้: {e}")
            self.status_label.config(text="เกิดข้อผิดพลาด", foreground="red")
    
    def toggle_file_selection(self, event):
        """สลับการเลือกไฟล์เมื่อ double-click"""
        item = self.tree.selection()[0]
        index = self.tree.index(item)
        
        if index < len(self.file_vars):
            current_value = self.file_vars[index].get()
            new_value = not current_value
            self.file_vars[index].set(new_value)
            
            # อัพเดท display
            values = list(self.tree.item(item, "values"))
            values[0] = "✓" if new_value else "✗"
            self.tree.item(item, values=values)
    
    def select_all_files(self):
        """เลือกไฟล์ทั้งหมดที่ใช้ได้"""
        for i, var in enumerate(self.file_vars):
            item = self.tree.get_children()[i]
            values = list(self.tree.item(item, "values"))
            
            # เลือกเฉพาะไฟล์ที่สถานะ "✓ พร้อม"
            if values[3] == "✓ พร้อม":
                var.set(True)
                values[0] = "✓"
                self.tree.item(item, values=values)
    
    def deselect_all_files(self):
        """ไม่เลือกไฟล์ทั้งหมด"""
        for i, var in enumerate(self.file_vars):
            var.set(False)
            item = self.tree.get_children()[i]
            values = list(self.tree.item(item, "values"))
            values[0] = "✗"
            self.tree.item(item, values=values)
    
    def get_selected_files(self):
        """รับรายการไฟล์ที่เลือก"""
        selected_files = []
        for i, var in enumerate(self.file_vars):
            if var.get() and i < len(self.file_info):
                selected_files.append(self.file_info[i])
        return selected_files
    
    def start_refresh(self):
        """เริ่มกระบวนการรีเฟช"""
        selected_files = self.get_selected_files()
        
        if not selected_files:
            messagebox.showwarning("ไม่มีไฟล์ที่เลือก", 
                                 "กรุณาเลือกไฟล์อย่างน้อย 1 ไฟล์ที่ต้องการรีเฟช")
            return
        
        # ยืนยันการดำเนินการ
        result = messagebox.askyesno("ยืนยันการรีเฟช", 
                                   f"คุณต้องการรีเฟช {len(selected_files)} ไฟล์ใช่หรือไม่?\n\n"
                                   "หมายเหตุ: ไฟล์จะถูกสำรองอัตโนมัติก่อนรีเฟช")
        
        if result:
            # เริ่มการรีเฟชใน thread แยก
            self.refresh_button.config(state="disabled", text="กำลังรีเฟช...")
            self.progress_var.set(0)
            self.progress_label.config(text="เริ่มต้นการรีเฟช...")
            
            thread = threading.Thread(target=self.refresh_worker, args=(selected_files,))
            thread.daemon = True
            thread.start()
    
    def refresh_worker(self, selected_files):
        """Worker function สำหรับรีเฟชไฟล์ (รันใน thread แยก)"""
        try:
            total_steps = 4  # verify, backup, cleanup, refresh
            current_step = 0
            
            # ขั้นตอนที่ 1: ตรวจสอบไฟล์
            self.update_progress(current_step / total_steps * 100, "ตรวจสอบไฟล์...")
            verification_result = self.verify_files()
            current_step += 1
            
            # ขั้นตอนที่ 2: สำรองไฟล์
            self.update_progress(current_step / total_steps * 100, "สร้างไฟล์สำรอง...")
            backup_result = self.create_backups()
            current_step += 1
            
            # ขั้นตอนที่ 3: ลบไฟล์สำรองเก่า
            self.update_progress(current_step / total_steps * 100, "ลบไฟล์สำรองเก่า...")
            deleted_count = self.auto_cleanup_backups(30)
            current_step += 1
            
            # ขั้นตอนที่ 4: รีเฟชไฟล์
            self.update_progress(current_step / total_steps * 100, "รีเฟชไฟล์ Excel...")
            
            # ใช้เฉพาะไฟล์ที่เลือก
            refresh_result = self.excel_refresher.refresh_multiple_files(
                selected_files,
                self.config_manager.settings
            )
            
            # เสร็จสิ้น
            self.update_progress(100, "เสร็จสิ้น!")
            
            # แสดงผลลัพธ์
            self.show_refresh_result(verification_result, backup_result, deleted_count, refresh_result)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("ข้อผิดพลาด", f"เกิดข้อผิดพลาดในการรีเฟช: {e}"))
        finally:
            # เปิดใช้งานปุ่มอีกครั้ง
            self.root.after(0, self.reset_refresh_button)
    
    def update_progress(self, value, text):
        """อัพเดท progress bar และข้อความ"""
        self.root.after(0, lambda: self.progress_var.set(value))
        self.root.after(0, lambda: self.progress_label.config(text=text))
    
    def show_refresh_result(self, verification_result, backup_result, deleted_count, refresh_result):
        """แสดงผลการรีเฟช"""
        message = "=== สรุปผลการรีเฟช ===\n\n"
        message += f"ไฟล์ที่ตรวจสอบ: ถูกต้อง {verification_result['total_valid']}, ไม่ถูกต้อง {verification_result['total_invalid']}\n"
        message += f"ไฟล์สำรอง: สำเร็จ {backup_result['success']}, ล้มเหลว {backup_result['failed']}\n"
        message += f"ไฟล์สำรองเก่าที่ลบ: {deleted_count} ไฟล์\n"
        message += f"Excel รีเฟช: สำเร็จ {refresh_result['success']}, ล้มเหลว {refresh_result['failed']}\n"
        message += f"รวมทั้งหมด: {refresh_result['total']} ไฟล์"
        
        if refresh_result['failed'] > 0:
            self.root.after(0, lambda: messagebox.showwarning("รีเฟชเสร็จสิ้น (มีข้อผิดพลาดบางส่วน)", message))
        else:
            self.root.after(0, lambda: messagebox.showinfo("รีเฟชเสร็จสิ้น", message))
        
        # รีเฟรชรายการไฟล์
        self.root.after(0, self.refresh_file_list)
    
    def reset_refresh_button(self):
        """รีเซ็ตปุ่มรีเฟช"""
        self.refresh_button.config(state="normal", text="🚀 เริ่มรีเฟชไฟล์ที่เลือก")
        self.progress_var.set(0)
        self.progress_label.config(text="พร้อมดำเนินการ")
    
    def open_settings(self):
        """เปิดหน้าต่างการตั้งค่า"""
        settings_window = SettingsWindow(self.root, self, self.refresh_file_list)
    
    def on_closing(self):
        """เมื่อปิดหน้าต่าง"""
        self.root.quit()
        self.root.destroy()
    
    def run(self):
        """เริ่มการทำงาน GUI"""
        self.root.mainloop()


def main():
    """ฟังก์ชันหลักสำหรับเริ่มการทำงาน GUI"""
    app = MainGUI()
    app.run()


if __name__ == "__main__":
    main()
