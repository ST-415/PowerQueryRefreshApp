"""
Main GUI Application
PowerQuery Refresh Application Main Window
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
    """Main application window"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PowerQuery Refresh Tool")
        self.root.geometry("750x500")
        self.root.resizable(False, False)  # ห้ามปรับขนาด
        
        # Modern Minimalist Theme Colors
        self.colors = {
            'bg': '#ffffff',           # White background
            'fg': '#2c2c2c',           # Dark text
            'select_bg': '#e67e22',    # Orange selection
            'button_bg': '#e67e22',    # Orange button
            'button_hover': '#f39c12', # Orange hover
            'accent': '#8b4513',       # Brown accent
            'success': '#27ae60',      # Green success
            'warning': '#f39c12',      # Orange warning
            'error': '#e74c3c',        # Red error
            'border': '#dee2e6',       # Light border
            'input_bg': '#f8f9fa',     # Light gray input
        }
        
        # Apply modern theme
        self.setup_modern_theme()
        
        # Configure messagebox colors
        self.configure_messagebox_colors()
        
        # Apply additional dark styling
        self.apply_additional_dark_styling()
        
        # Set window icon (optional)
        try:
            # You can add a .ico file here if available
            # self.root.iconbitmap('icon.ico')
            pass
        except:
            pass
        
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
    
    def setup_modern_theme(self):
        """Setup modern minimalist theme with white, brown, orange, and black colors"""
        # Configure root window
        self.root.configure(bg=self.colors['bg'])
        
        # Create custom style
        self.style = ttk.Style()
        
        # Configure ttk styles for modern theme
        self.style.theme_create('modern_minimal', parent='clam', settings={
            'TLabel': {
                'configure': {
                    'background': self.colors['bg'],
                    'foreground': self.colors['fg'],
                    'font': ('Segoe UI', 9)
                }
            },
            'TButton': {
                'configure': {
                    'background': self.colors['button_bg'],
                    'foreground': 'white',
                    'borderwidth': 0,
                    'focuscolor': 'none',
                    'font': ('Segoe UI', 9, 'bold'),
                    'relief': 'flat',
                    'padding': [20, 8]
                },
                'map': {
                    'background': [('active', self.colors['button_hover']),
                                 ('pressed', self.colors['accent'])]
                }
            },
            'TFrame': {
                'configure': {
                    'background': self.colors['bg'],
                    'borderwidth': 0,
                    'relief': 'flat'
                }
            },
            'TLabelFrame': {
                'configure': {
                    'background': self.colors['bg'],
                    'foreground': self.colors['fg'],
                    'borderwidth': 1,
                    'relief': 'solid',
                    'font': ('Segoe UI', 9, 'bold'),
                    'bordercolor': self.colors['border']
                }
            },
            'TLabelFrame.Label': {
                'configure': {
                    'background': self.colors['bg'],
                    'foreground': self.colors['accent']
                }
            },
            'Treeview': {
                'configure': {
                    'background': self.colors['input_bg'],
                    'foreground': self.colors['fg'],
                    'fieldbackground': self.colors['input_bg'],
                    'borderwidth': 1,
                    'relief': 'solid',
                    'font': ('Segoe UI', 9),
                    'insertcolor': self.colors['fg']
                },
                'map': {
                    'background': [('selected', self.colors['select_bg'])],
                    'foreground': [('selected', 'white')]
                }
            },
            'Treeview.Heading': {
                'configure': {
                    'background': self.colors['border'],
                    'foreground': self.colors['fg'],
                    'relief': 'flat',
                    'font': ('Segoe UI', 9, 'bold'),
                    'borderwidth': 1
                },
                'map': {
                    'background': [('active', self.colors['input_bg'])]
                }
            },
            'Horizontal.TProgressbar': {
                'configure': {
                    'background': self.colors['accent'],
                    'troughcolor': self.colors['border'],
                    'borderwidth': 0,
                    'lightcolor': self.colors['accent'],
                    'darkcolor': self.colors['accent']
                }
            },
            'TScrollbar': {
                'configure': {
                    'background': self.colors['border'],
                    'troughcolor': self.colors['bg'],
                    'borderwidth': 1,
                    'arrowcolor': self.colors['fg'],
                    'darkcolor': self.colors['border'],
                    'lightcolor': self.colors['border']
                },
                'map': {
                    'background': [('active', self.colors['input_bg'])]
                }
            },
            'TNotebook': {
                'configure': {
                    'background': self.colors['bg'],
                    'borderwidth': 1,
                    'tabmargins': [2, 5, 2, 0]
                }
            },
            'TNotebook.Tab': {
                'configure': {
                    'background': self.colors['input_bg'],
                    'foreground': self.colors['fg'],
                    'padding': [20, 10],
                    'font': ('Segoe UI', 9),
                    'borderwidth': 1,
                    'relief': 'solid'
                },
                'map': {
                    'background': [('selected', self.colors['bg']),
                                 ('active', self.colors['button_bg'])],
                    'foreground': [('selected', self.colors['accent']),
                                 ('active', 'white')]
                }
            }
        })
        
        # Apply the custom theme
        self.style.theme_use('modern_minimal')
    
    def create_widgets(self):
        """Create all widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="8")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure responsive layout
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="PowerQuery Refresh Tool", 
                               font=('Segoe UI', 12, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 12))
        
        # Button control frame - moved to top
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 12))
        button_frame.columnconfigure(5, weight=1)
        
        # Settings button
        settings_button = ttk.Button(button_frame, text="⚙ Settings", 
                                   command=self.open_settings, width=10)
        settings_button.grid(row=0, column=0, padx=(0, 4))
        
        # Refresh list button
        refresh_list_button = ttk.Button(button_frame, text="🔄 Refresh", 
                                       command=self.refresh_file_list, width=10)
        refresh_list_button.grid(row=0, column=1, padx=4)
        
        # Select/Deselect buttons
        select_all_button = ttk.Button(button_frame, text="✓ All", 
                                     command=self.select_all_files, width=8)
        select_all_button.grid(row=0, column=2, padx=4)
        
        deselect_all_button = ttk.Button(button_frame, text="✗ None", 
                                       command=self.deselect_all_files, width=8)
        deselect_all_button.grid(row=0, column=3, padx=4)
        
        # Start refresh button
        self.refresh_button = ttk.Button(button_frame, text="🚀 Start Refresh", 
                                       command=self.start_refresh, width=12)
        self.refresh_button.grid(row=0, column=4, padx=(4, 8))
        
        # Status label - moved to right side of button frame
        self.status_label = ttk.Label(button_frame, text="Loading files...", 
                                     foreground=self.colors['warning'])
        self.status_label.grid(row=0, column=5, sticky=tk.E)
        
        # Files frame
        files_frame = ttk.LabelFrame(main_frame, text="Excel Files", padding="8")
        files_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 8))
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(0, weight=1)
        
        # Treeview for file list
        self.create_file_treeview(files_frame)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, 
                                          mode='determinate')
        self.progress_bar.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(8, 0))
        
        # Progress label
        self.progress_label = ttk.Label(main_frame, text="Ready")
        self.progress_label.grid(row=4, column=0, pady=(4, 0))
    
    def create_file_treeview(self, parent):
        """Create treeview for file list"""
        # Create Treeview
        columns = ("select", "name", "path", "status")
        self.tree = ttk.Treeview(parent, columns=columns, show="headings", height=8)
        
        # Configure column headings
        self.tree.heading("select", text="Select")
        self.tree.heading("name", text="File Name")
        self.tree.heading("path", text="Path")
        self.tree.heading("status", text="Status")
        
        # Configure column widths
        self.tree.column("select", width=60, anchor="center")
        self.tree.column("name", width=150)
        self.tree.column("path", width=250)
        self.tree.column("status", width=80, anchor="center")
        
        # Apply dark theme using style map instead of direct configure
        style_name = "Dark.Treeview"
        self.style.configure(style_name,
                           background=self.colors['input_bg'],
                           foreground=self.colors['fg'],
                           fieldbackground=self.colors['input_bg'])
        self.style.map(style_name,
                      background=[('selected', self.colors['select_bg'])],
                      foreground=[('selected', 'white')])
        self.tree.configure(style=style_name)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Additional dark theme configuration for treeview
        self.tree.tag_configure('oddrow', background=self.colors['input_bg'])
        self.tree.tag_configure('evenrow', background=self.colors['bg'])
        
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
        """Refresh file list and check status"""
        self.status_label.config(text="Checking files...", foreground=self.colors['warning'])
        self.root.update_idletasks()
        
        try:
            # Verify files
            verification_result = self.verify_files()
            
            # Clear old list
            self.tree.delete(*self.tree.get_children())
            self.file_vars.clear()
            self.file_info.clear()
            
            # Add valid files
            for file_info in verification_result["excel_files"]["valid"]:
                var = tk.BooleanVar(value=True)  # Selected by default
                self.file_vars.append(var)
                self.file_info.append(file_info)
                
                name = os.path.basename(file_info["path"])
                self.tree.insert("", "end", values=(
                    "✓", name, file_info["path"], "✓ Ready"
                ))
            
            # Add invalid files
            for file_info in verification_result["excel_files"]["invalid"]:
                var = tk.BooleanVar(value=False)  # Not selected
                self.file_vars.append(var)
                self.file_info.append(file_info)
                
                name = os.path.basename(file_info["path"])
                self.tree.insert("", "end", values=(
                    "✗", name, file_info["path"], "✗ Missing"
                ))
            
            # Update status
            total_files = len(self.file_info)
            valid_files = verification_result["total_valid"]
            invalid_files = verification_result["total_invalid"]
            
            if invalid_files > 0:
                self.status_label.config(
                    text=f"{total_files} files ({valid_files} ready, {invalid_files} missing)",
                    foreground=self.colors['warning']
                )
            else:
                self.status_label.config(
                    text=f"{valid_files} files (all ready)",
                    foreground=self.colors['success']
                )
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file list: {e}")
            self.status_label.config(text="Error occurred", foreground=self.colors['error'])
    
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
        """Select all available files"""
        for i, var in enumerate(self.file_vars):
            item = self.tree.get_children()[i]
            values = list(self.tree.item(item, "values"))
            
            # Select only files with "✓ Ready" status
            if values[3] == "✓ Ready":
                var.set(True)
                values[0] = "✓"
                self.tree.item(item, values=values)
    
    def deselect_all_files(self):
        """Deselect all files"""
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
        """Start refresh process"""
        selected_files = self.get_selected_files()
        
        if not selected_files:
            messagebox.showwarning("No Files Selected", 
                                 "Please select at least 1 file to refresh")
            return
        
        # Confirm action
        result = messagebox.askyesno("Confirm Refresh", 
                                   f"Do you want to refresh {len(selected_files)} file(s)?\n\n"
                                   "Note: Files will be automatically backed up before refresh")
        
        if result:
            # Start refresh in separate thread
            self.refresh_button.config(state="disabled", text="Refreshing...")
            self.progress_var.set(0)
            self.progress_label.config(text="Starting refresh...")
            
            thread = threading.Thread(target=self.refresh_worker, args=(selected_files,))
            thread.daemon = True
            thread.start()
    
    def refresh_worker(self, selected_files):
        """Worker function for refreshing files (runs in separate thread)"""
        try:
            total_steps = 4  # verify, backup, cleanup, refresh
            current_step = 0
            
            # Step 1: Verify files
            self.update_progress(current_step / total_steps * 100, "Verifying files...")
            verification_result = self.verify_files()
            current_step += 1
            
            # Step 2: Create backups
            self.update_progress(current_step / total_steps * 100, "Creating backups...")
            backup_result = self.create_backups()
            current_step += 1
            
            # Step 3: Cleanup old backups
            self.update_progress(current_step / total_steps * 100, "Cleaning up old backups...")
            deleted_count = self.auto_cleanup_backups(30)
            current_step += 1
            
            # Step 4: Refresh files
            self.update_progress(current_step / total_steps * 100, "Refreshing Excel files...")
            
            # Use only selected files
            refresh_result = self.excel_refresher.refresh_multiple_files(
                selected_files,
                self.config_manager.settings
            )
            
            # Complete
            self.update_progress(100, "Complete!")
            
            # Show results
            self.show_refresh_result(verification_result, backup_result, deleted_count, refresh_result)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Refresh failed: {e}"))
        finally:
            # Re-enable button
            self.root.after(0, self.reset_refresh_button)
    
    def update_progress(self, value, text):
        """อัพเดท progress bar และข้อความ"""
        self.root.after(0, lambda: self.progress_var.set(value))
        self.root.after(0, lambda: self.progress_label.config(text=text))
    
    def show_refresh_result(self, verification_result, backup_result, deleted_count, refresh_result):
        """Show refresh results"""
        message = "=== Refresh Summary ===\n\n"
        message += f"Files verified: {verification_result['total_valid']} valid, {verification_result['total_invalid']} invalid\n"
        message += f"Backups: {backup_result['success']} success, {backup_result['failed']} failed\n"
        message += f"Old backups deleted: {deleted_count} files\n"
        message += f"Excel refresh: {refresh_result['success']} success, {refresh_result['failed']} failed\n"
        message += f"Total: {refresh_result['total']} files"
        
        if refresh_result['failed'] > 0:
            self.root.after(0, lambda: messagebox.showwarning("Refresh Complete (with errors)", message))
        else:
            self.root.after(0, lambda: messagebox.showinfo("Refresh Complete", message))
        
        # Refresh file list
        self.root.after(0, self.refresh_file_list)
    
    def reset_refresh_button(self):
        """Reset refresh button"""
        self.refresh_button.config(state="normal", text="🚀 Start Refresh")
        self.progress_var.set(0)
        self.progress_label.config(text="Ready")
    
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
    
    def configure_messagebox_colors(self):
        """Configure messagebox colors for dark theme"""
        try:
            # Configure tkinter default colors for messageboxes
            self.root.option_add('*Dialog.msg.background', self.colors['bg'])
            self.root.option_add('*Dialog.msg.foreground', self.colors['fg'])
            self.root.option_add('*Dialog.background', self.colors['bg'])
            self.root.option_add('*Dialog.foreground', self.colors['fg'])
            self.root.option_add('*Button.background', self.colors['button_bg'])
            self.root.option_add('*Button.foreground', 'white')
            self.root.option_add('*Button.activeBackground', self.colors['button_hover'])
            self.root.option_add('*Button.activeForeground', 'white')
        except:
            pass
    
    def apply_additional_dark_styling(self):
        """Apply additional dark styling to ensure no white backgrounds"""
        try:
            # Force dark colors on any remaining white elements
            self.root.tk_setPalette(
                background=self.colors['bg'],
                foreground=self.colors['fg'],
                activeBackground=self.colors['button_bg'],
                activeForeground='white',
                selectBackground=self.colors['select_bg'],
                selectForeground='white'
            )
            
            # Configure default button and entry colors
            self.root.option_add('*TCombobox*Listbox.background', self.colors['input_bg'])
            self.root.option_add('*TCombobox*Listbox.foreground', self.colors['fg'])
            self.root.option_add('*TCombobox*Listbox.selectBackground', self.colors['select_bg'])
            
        except Exception as e:
            pass  # Ignore any errors in styling


def main():
    """ฟังก์ชันหลักสำหรับเริ่มการทำงาน GUI"""
    app = MainGUI()
    app.run()


if __name__ == "__main__":
    main()
