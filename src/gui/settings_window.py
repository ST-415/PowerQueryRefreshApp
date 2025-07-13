"""
Settings Window
หน้าต่างการตั้งค่าสำหรับ PowerQuery Refresh Application
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import json
from typing import Dict, List, Any, Optional, Callable


class SettingsWindow:
    """หน้าต่างการตั้งค่า"""
    
    def __init__(self, parent: tk.Tk, main_gui, refresh_callback: Callable):
        self.parent = parent
        self.main_gui = main_gui
        self.refresh_callback = refresh_callback
        
        # สร้างหน้าต่างใหม่
        self.window = tk.Toplevel(parent)
        self.window.title("การตั้งค่า - PowerQuery Refresh Application")
        self.window.geometry("900x700")
        self.window.resizable(True, True)
        
        # ทำให้หน้าต่างนี้เป็น modal
        self.window.transient(parent)
        self.window.grab_set()
        
        # ตัวแปรสำหรับการตั้งค่า
        self.auto_save_var = tk.BooleanVar()
        self.backup_before_refresh_var = tk.BooleanVar()
        self.log_refresh_activity_var = tk.BooleanVar()
        self.refresh_timeout_var = tk.IntVar()
        
        # สร้าง GUI
        self.create_widgets()
        self.load_settings()
        
        # ตั้งค่า protocol สำหรับปิดหน้าต่าง
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # ให้โฟกัสที่หน้าต่างนี้
        self.window.focus_set()
    
    def create_widgets(self):
        """สร้าง widgets ทั้งหมด"""
        # Frame หลัก
        main_frame = ttk.Frame(self.window, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # กำหนด weight สำหรับ responsive
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # หัวข้อ
        title_label = ttk.Label(main_frame, text="การตั้งค่า", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # สร้าง notebook สำหรับแท็บ
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        
        # แท็บไฟล์
        self.create_files_tab(notebook)
        
        # แท็บการตั้งค่าทั่วไป
        self.create_general_settings_tab(notebook)
        
        # Frame สำหรับปุ่มควบคุม
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        control_frame.columnconfigure(0, weight=1)
        
        # ปุ่มควบคุม
        button_frame = ttk.Frame(control_frame)
        button_frame.grid(row=0, column=0)
        
        save_button = ttk.Button(button_frame, text="💾 บันทึกการตั้งค่า", 
                               command=self.save_settings, style='Accent.TButton')
        save_button.grid(row=0, column=0, padx=(0, 10))
        
        cancel_button = ttk.Button(button_frame, text="❌ ยกเลิก", 
                                 command=self.on_closing)
        cancel_button.grid(row=0, column=1)
    
    def create_files_tab(self, notebook):
        """สร้างแท็บจัดการไฟล์"""
        files_frame = ttk.Frame(notebook, padding="15")
        notebook.add(files_frame, text="📁 จัดการไฟล์")
        
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(1, weight=1)
        
        # หัวข้อและคำอธิบาย
        header_label = ttk.Label(files_frame, text="ไฟล์ Excel ที่ต้องการรีเฟช", 
                                font=('Arial', 12, 'bold'))
        header_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Frame สำหรับ Treeview และปุ่ม
        tree_frame = ttk.Frame(files_frame)
        tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # สร้าง Treeview สำหรับรายการไฟล์
        self.create_files_treeview(tree_frame)
        
        # Frame สำหรับปุ่มจัดการไฟล์
        files_button_frame = ttk.Frame(files_frame)
        files_button_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        add_file_button = ttk.Button(files_button_frame, text="➕ เพิ่มไฟล์", 
                                   command=self.add_file)
        add_file_button.grid(row=0, column=0, padx=(0, 10))
        
        remove_file_button = ttk.Button(files_button_frame, text="➖ ลบไฟล์ที่เลือก", 
                                      command=self.remove_selected_file)
        remove_file_button.grid(row=0, column=1, padx=(0, 10))
        
        edit_file_button = ttk.Button(files_button_frame, text="✏️ แก้ไขไฟล์", 
                                    command=self.edit_selected_file)
        edit_file_button.grid(row=0, column=2)
    
    def create_files_treeview(self, parent):
        """สร้าง Treeview สำหรับรายการไฟล์"""
        # สร้าง Treeview
        columns = ("name", "path", "status")
        self.files_tree = ttk.Treeview(parent, columns=columns, show="headings", height=12)
        
        # กำหนดหัวคอลัมน์
        self.files_tree.heading("name", text="ชื่อไฟล์")
        self.files_tree.heading("path", text="เส้นทางไฟล์")
        self.files_tree.heading("status", text="สถานะ")
        
        # กำหนดความกว้างคอลัมน์
        self.files_tree.column("name", width=200)
        self.files_tree.column("path", width=400)
        self.files_tree.column("status", width=100, anchor="center")
        
        # Scrollbar
        files_scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, 
                                       command=self.files_tree.yview)
        self.files_tree.configure(yscrollcommand=files_scrollbar.set)
        
        # วางใน grid
        self.files_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    
    def create_general_settings_tab(self, notebook):
        """สร้างแท็บการตั้งค่าทั่วไป"""
        settings_frame = ttk.Frame(notebook, padding="15")
        notebook.add(settings_frame, text="⚙️ การตั้งค่าทั่วไป")
        
        # Frame สำหรับการตั้งค่าต่างๆ
        options_frame = ttk.LabelFrame(settings_frame, text="ตัวเลือกการทำงาน", padding="15")
        options_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        options_frame.columnconfigure(1, weight=1)
        
        # Auto Save
        ttk.Checkbutton(options_frame, text="บันทึกอัตโนมัติหลังรีเฟช (Auto Save)", 
                       variable=self.auto_save_var).grid(row=0, column=0, columnspan=2, 
                                                        sticky=tk.W, pady=(0, 10))
        
        # Backup Before Refresh
        ttk.Checkbutton(options_frame, text="สร้างไฟล์สำรองก่อนรีเฟช (Backup Before Refresh)", 
                       variable=self.backup_before_refresh_var).grid(row=1, column=0, columnspan=2, 
                                                                   sticky=tk.W, pady=(0, 10))
        
        # Log Refresh Activity
        ttk.Checkbutton(options_frame, text="บันทึกกิจกรรมการรีเฟช (Log Refresh Activity)", 
                       variable=self.log_refresh_activity_var).grid(row=2, column=0, columnspan=2, 
                                                                  sticky=tk.W, pady=(0, 15))
        
        # Refresh Timeout
        ttk.Label(options_frame, text="เวลารอสูงสุดในการรีเฟช (นาที):").grid(row=3, column=0, 
                                                                           sticky=tk.W, pady=(0, 5))
        timeout_spinbox = ttk.Spinbox(options_frame, from_=5, to=180, width=10, 
                                    textvariable=self.refresh_timeout_var)
        timeout_spinbox.grid(row=3, column=1, sticky=tk.W, padx=(10, 0), pady=(0, 5))
        
        # คำอธิบายการตั้งค่า
        info_frame = ttk.LabelFrame(settings_frame, text="คำอธิบาย", padding="15")
        info_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        info_frame.columnconfigure(0, weight=1)
        settings_frame.rowconfigure(1, weight=1)
        
        info_text = tk.Text(info_frame, height=8, wrap=tk.WORD, 
                           font=('Arial', 9), state=tk.DISABLED)
        info_scrollbar = ttk.Scrollbar(info_frame, orient=tk.VERTICAL, command=info_text.yview)
        info_text.configure(yscrollcommand=info_scrollbar.set)
        
        info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        info_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # เพิ่มข้อความคำอธิบาย
        info_content = """การตั้งค่าทั่วไป:

• Auto Save: บันทึกไฟล์อัตโนมัติหลังจากรีเฟชเสร็จสิ้น

• Backup Before Refresh: สร้างไฟล์สำรองก่อนดำเนินการรีเฟช เพื่อป้องกันการสูญหายของข้อมูล

• Log Refresh Activity: บันทึกกิจกรรมการรีเฟชลงในไฟล์ log เพื่อติดตามการทำงาน

• Refresh Timeout: กำหนดเวลารอสูงสุดสำหรับการรีเฟชแต่ละไฟล์ หากเกินเวลาที่กำหนดจะหยุดการทำงาน

หมายเหตุ: การตั้งค่าจะมีผลทันทีหลังจากบันทึก"""
        
        info_text.config(state=tk.NORMAL)
        info_text.insert(tk.END, info_content)
        info_text.config(state=tk.DISABLED)
    
    def load_settings(self):
        """โหลดการตั้งค่าปัจจุบัน"""
        try:
            # โหลดรายการไฟล์
            self.refresh_files_list()
            
            # โหลดการตั้งค่าทั่วไป
            settings = self.main_gui.config_manager.settings
            self.auto_save_var.set(settings.get("auto_save", True))
            self.backup_before_refresh_var.set(settings.get("backup_before_refresh", True))
            self.log_refresh_activity_var.set(settings.get("log_refresh_activity", True))
            self.refresh_timeout_var.set(settings.get("refresh_timeout_minutes", 60))
            
        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถโหลดการตั้งค่าได้: {e}")
    
    def refresh_files_list(self):
        """รีเฟรชรายการไฟล์"""
        # ล้างรายการเก่า
        self.files_tree.delete(*self.files_tree.get_children())
        
        # เพิ่มไฟล์ใหม่
        for file_info in self.main_gui.config_manager.excel_files:
            name = os.path.basename(file_info["path"])
            path = file_info["path"]
            
            # ตรวจสอบสถานะไฟล์
            if os.path.exists(path) and path.endswith(('.xlsx', '.xls')):
                status = "✓ พร้อม"
            else:
                status = "✗ ไม่พบ"
            
            self.files_tree.insert("", "end", values=(name, path, status))
    
    def add_file(self):
        """เพิ่มไฟล์ใหม่"""
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=os.getcwd()
        )
        
        if file_path:
            # แปลงเป็น relative path ถ้าเป็นไปได้
            try:
                rel_path = os.path.relpath(file_path, os.getcwd())
                if not rel_path.startswith('..'):
                    file_path = rel_path
            except ValueError:
                # ใช้ absolute path ถ้าไม่สามารถสร้าง relative path ได้
                pass
            
            # ตรวจสอบว่าไฟล์ซ้ำหรือไม่
            for file_info in self.main_gui.config_manager.excel_files:
                if file_info["path"] == file_path:
                    messagebox.showwarning("ไฟล์ซ้ำ", "ไฟล์นี้มีอยู่ในรายการแล้ว")
                    return
            
            # เพิ่มไฟล์ใหม่
            new_file = {"path": file_path}
            self.main_gui.config_manager.excel_files.append(new_file)
            
            # รีเฟรชรายการ
            self.refresh_files_list()
    
    def remove_selected_file(self):
        """ลบไฟล์ที่เลือก"""
        selected_items = self.files_tree.selection()
        
        if not selected_items:
            messagebox.showwarning("ไม่มีไฟล์ที่เลือก", "กรุณาเลือกไฟล์ที่ต้องการลบ")
            return
        
        # ยืนยันการลบ
        result = messagebox.askyesno("ยืนยันการลบ", 
                                   f"คุณต้องการลบไฟล์ {len(selected_items)} ไฟล์ออกจากรายการใช่หรือไม่?")
        
        if result:
            # รับ path ของไฟล์ที่เลือก
            selected_paths = []
            for item in selected_items:
                values = self.files_tree.item(item, "values")
                selected_paths.append(values[1])  # path อยู่ในคอลัมน์ที่ 1
            
            # ลบออกจาก config
            self.main_gui.config_manager.excel_files = [
                file_info for file_info in self.main_gui.config_manager.excel_files
                if file_info["path"] not in selected_paths
            ]
            
            # รีเฟรชรายการ
            self.refresh_files_list()
    
    def edit_selected_file(self):
        """แก้ไขไฟล์ที่เลือก"""
        selected_items = self.files_tree.selection()
        
        if not selected_items:
            messagebox.showwarning("ไม่มีไฟล์ที่เลือก", "กรุณาเลือกไฟล์ที่ต้องการแก้ไข")
            return
        
        if len(selected_items) > 1:
            messagebox.showwarning("เลือกไฟล์มากเกินไป", "กรุณาเลือกไฟล์เพียง 1 ไฟล์เท่านั้น")
            return
        
        # รับข้อมูลไฟล์ปัจจุบัน
        item = selected_items[0]
        values = self.files_tree.item(item, "values")
        current_path = values[1]
        
        # เปิด dialog สำหรับเลือกไฟล์ใหม่
        new_path = filedialog.askopenfilename(
            title="เลือกไฟล์ Excel ใหม่",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialfile=os.path.basename(current_path),
            initialdir=os.path.dirname(current_path) if os.path.isabs(current_path) else os.getcwd()
        )
        
        if new_path:
            # แปลงเป็น relative path ถ้าเป็นไปได้
            try:
                rel_path = os.path.relpath(new_path, os.getcwd())
                if not rel_path.startswith('..'):
                    new_path = rel_path
            except ValueError:
                pass
            
            # อัพเดทใน config
            for file_info in self.main_gui.config_manager.excel_files:
                if file_info["path"] == current_path:
                    file_info["path"] = new_path
                    break
            
            # รีเฟรชรายการ
            self.refresh_files_list()
    
    def save_settings(self):
        """บันทึกการตั้งค่า"""
        try:
            # อัพเดทการตั้งค่าทั่วไป
            self.main_gui.config_manager.settings["auto_save"] = self.auto_save_var.get()
            self.main_gui.config_manager.settings["backup_before_refresh"] = self.backup_before_refresh_var.get()
            self.main_gui.config_manager.settings["log_refresh_activity"] = self.log_refresh_activity_var.get()
            self.main_gui.config_manager.settings["refresh_timeout_minutes"] = self.refresh_timeout_var.get()
            
            # บันทึกลงไฟล์
            self.main_gui.config_manager.save_config()
            
            messagebox.showinfo("บันทึกสำเร็จ", "การตั้งค่าถูกบันทึกเรียบร้อยแล้ว")
            
            # รีเฟรชหน้าหลัก
            if self.refresh_callback:
                self.refresh_callback()
            
            # ปิดหน้าต่าง
            self.on_closing()
            
        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถบันทึกการตั้งค่าได้: {e}")
    
    def on_closing(self):
        """เมื่อปิดหน้าต่าง"""
        self.window.grab_release()
        self.window.destroy()
