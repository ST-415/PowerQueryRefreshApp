"""
Settings Window
Settings window for PowerQuery Refresh Application
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import json
from typing import Dict, List, Any, Optional, Callable


class SettingsWindow:
    """Settings window"""
    
    def __init__(self, parent: tk.Tk, main_gui, refresh_callback: Callable):
        self.parent = parent
        self.main_gui = main_gui
        self.refresh_callback = refresh_callback
        
        # Create new window
        self.window = tk.Toplevel(parent)
        self.window.title("Settings - PowerQuery Refresh Tool")
        self.window.geometry("700x550")
        self.window.resizable(False, False)  # ห้ามปรับขนาด
        
        # Apply same colors as main window
        self.colors = main_gui.colors
        self.window.configure(bg=self.colors['bg'])
        
        # Configure messagebox colors for this window
        self.window.option_add('*Dialog.msg.background', self.colors['bg'])
        self.window.option_add('*Dialog.msg.foreground', self.colors['fg'])
        self.window.option_add('*Dialog.background', self.colors['bg'])
        self.window.option_add('*Dialog.foreground', self.colors['fg'])
        self.window.option_add('*Button.background', self.colors['button_bg'])
        self.window.option_add('*Button.foreground', 'white')
        
        # Make this window modal
        self.window.transient(parent)
        self.window.grab_set()
        
        # Variables for settings
        self.auto_save_var = tk.BooleanVar()
        self.backup_before_refresh_var = tk.BooleanVar()
        self.log_refresh_activity_var = tk.BooleanVar()
        self.refresh_timeout_var = tk.IntVar()
        
        # Apply modern theme
        self.setup_modern_theme()
        
        # Create GUI
        self.create_widgets()
        self.load_settings()
        
        # Set protocol for window closing
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Focus on this window
        self.window.focus_set()
    
    def setup_modern_theme(self):
        """Setup modern theme for settings window"""
        # Use the existing style from main window if available
        try:
            # Try to use the existing style
            self.style = self.main_gui.style
        except:
            # Create new style if needed
            self.style = ttk.Style()
            
            # Create comprehensive modern theme
            self.style.theme_create('settings_modern', parent='clam', settings={
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
                'TCheckbutton': {
                    'configure': {
                        'background': self.colors['bg'],
                        'foreground': self.colors['fg'],
                        'focuscolor': 'none',
                        'font': ('Segoe UI', 9),
                        'borderwidth': 0
                    },
                    'map': {
                        'background': [('active', self.colors['bg'])],
                        'foreground': [('active', self.colors['accent'])]
                    }
                },
                'TSpinbox': {
                    'configure': {
                        'fieldbackground': self.colors['input_bg'],
                        'background': self.colors['input_bg'],
                        'foreground': self.colors['fg'],
                        'borderwidth': 1,
                        'font': ('Segoe UI', 9),
                        'insertcolor': self.colors['fg'],
                        'selectbackground': self.colors['select_bg'],
                        'selectforeground': 'white'
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
                        'background': self.colors['border'],
                        'foreground': self.colors['fg'],
                        'padding': [12, 8],
                        'font': ('Segoe UI', 9)
                    },
                    'map': {
                        'background': [('selected', self.colors['input_bg']),
                                     ('active', self.colors['button_bg'])],
                        'foreground': [('selected', 'white'),
                                     ('active', 'white')]
                    }
                }
            })
            
            # Apply the theme
            self.style.theme_use('settings_modern')

    def create_widgets(self):
        """Create all widgets"""
        # Main frame
        main_frame = ttk.Frame(self.window, padding="12")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure responsive layout
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Settings", 
                               font=('Segoe UI', 12, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 12))
        
        # Top button frame - contains all buttons
        top_button_frame = ttk.Frame(main_frame)
        top_button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 12))
        top_button_frame.columnconfigure(4, weight=1)
        
        # File management buttons
        self.add_file_button = ttk.Button(top_button_frame, text="➕ Add File", 
                                         command=self.add_file, width=10)
        self.add_file_button.grid(row=0, column=0, padx=(0, 4))
        
        self.remove_file_button = ttk.Button(top_button_frame, text="➖ Remove", 
                                           command=self.remove_selected_file, width=10)
        self.remove_file_button.grid(row=0, column=1, padx=4)
        
        self.edit_file_button = ttk.Button(top_button_frame, text="✏️ Edit", 
                                         command=self.edit_selected_file, width=10)
        self.edit_file_button.grid(row=0, column=2, padx=(4, 20))
        
        # Settings control buttons
        save_button = ttk.Button(top_button_frame, text="💾 Save Settings", 
                               command=self.save_settings, width=14)
        save_button.grid(row=0, column=3, padx=(0, 4))
        
        cancel_button = ttk.Button(top_button_frame, text="❌ Cancel", 
                                 command=self.on_closing, width=10)
        cancel_button.grid(row=0, column=4, sticky=tk.E)
        
        # Create notebook for tabs
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Files tab
        self.create_files_tab(notebook)
        
        # General settings tab
        self.create_general_settings_tab(notebook)
    
    def create_files_tab(self, notebook):
        """Create file management tab"""
        files_frame = ttk.Frame(notebook, padding="12")
        notebook.add(files_frame, text="📁 File Management")
        
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(1, weight=1)
        
        # Header and description
        header_label = ttk.Label(files_frame, text="Excel Files to Refresh", 
                                font=('Segoe UI', 10, 'bold'))
        header_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 8))
        
        # Frame for Treeview - buttons are now at top level
        tree_frame = ttk.Frame(files_frame)
        tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # Create Treeview for file list
        self.create_files_treeview(tree_frame)
    
    def create_files_treeview(self, parent):
        """Create treeview for file list"""
        # Create Treeview
        columns = ("name", "path", "status")
        self.files_tree = ttk.Treeview(parent, columns=columns, show="headings", height=10)
        
        # Configure column headings
        self.files_tree.heading("name", text="File Name")
        self.files_tree.heading("path", text="File Path")
        self.files_tree.heading("status", text="Status")
        
        # Configure column widths
        self.files_tree.column("name", width=150)
        self.files_tree.column("path", width=300)
        self.files_tree.column("status", width=80, anchor="center")
        
        # Apply dark theme using style map instead of direct configure
        try:
            style_name = "Settings.Dark.Treeview"
            self.style.configure(style_name,
                               background=self.colors['input_bg'],
                               foreground=self.colors['fg'],
                               fieldbackground=self.colors['input_bg'])
            self.style.map(style_name,
                          background=[('selected', self.colors['select_bg'])],
                          foreground=[('selected', 'white')])
            self.files_tree.configure(style=style_name)
        except:
            pass
        
        # Scrollbar
        files_scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, 
                                       command=self.files_tree.yview)
        self.files_tree.configure(yscrollcommand=files_scrollbar.set)
        
        # Grid layout
        self.files_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Additional dark theme configuration for treeview
        self.files_tree.tag_configure('oddrow', background=self.colors['input_bg'])
        self.files_tree.tag_configure('evenrow', background=self.colors['bg'])
    
    def create_general_settings_tab(self, notebook):
        """Create general settings tab"""
        settings_frame = ttk.Frame(notebook, padding="12")
        notebook.add(settings_frame, text="⚙️ General Settings")
        
        # Frame for various settings
        options_frame = ttk.LabelFrame(settings_frame, text="Refresh Options", padding="12")
        options_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        options_frame.columnconfigure(1, weight=1)
        
        # Auto Save
        ttk.Checkbutton(options_frame, text="Auto save after refresh", 
                       variable=self.auto_save_var).grid(row=0, column=0, columnspan=2, 
                                                        sticky=tk.W, pady=(0, 8))
        
        # Backup Before Refresh
        ttk.Checkbutton(options_frame, text="Create backup before refresh", 
                       variable=self.backup_before_refresh_var).grid(row=1, column=0, columnspan=2, 
                                                                   sticky=tk.W, pady=(0, 8))
        
        # Log Refresh Activity
        ttk.Checkbutton(options_frame, text="Log refresh activity", 
                       variable=self.log_refresh_activity_var).grid(row=2, column=0, columnspan=2, 
                                                                  sticky=tk.W, pady=(0, 12))
        
        # Refresh Timeout
        ttk.Label(options_frame, text="Refresh timeout (minutes):").grid(row=3, column=0, 
                                                                           sticky=tk.W, pady=(0, 4))
        timeout_spinbox = ttk.Spinbox(options_frame, from_=5, to=180, width=8, 
                                    textvariable=self.refresh_timeout_var)
        timeout_spinbox.grid(row=3, column=1, sticky=tk.W, padx=(8, 0), pady=(0, 4))
        
        # Description frame
        info_frame = ttk.LabelFrame(settings_frame, text="Description", padding="12")
        info_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        info_frame.columnconfigure(0, weight=1)
        settings_frame.rowconfigure(1, weight=1)
        
        info_text = tk.Text(info_frame, height=6, wrap=tk.WORD, 
                           font=('Segoe UI', 9), state=tk.DISABLED,
                           bg=self.colors['input_bg'], fg=self.colors['fg'],
                           borderwidth=1, relief='solid',
                           selectbackground=self.colors['select_bg'],
                           selectforeground='white',
                           insertbackground=self.colors['fg'])
        info_scrollbar = ttk.Scrollbar(info_frame, orient=tk.VERTICAL, command=info_text.yview)
        info_text.configure(yscrollcommand=info_scrollbar.set)
        
        info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        info_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Add description text
        info_content = """General Settings:

• Auto Save: Automatically save files after refresh is complete

• Backup Before Refresh: Create backup copies before performing refresh to prevent data loss

• Log Refresh Activity: Record refresh activities in log files for tracking and debugging

• Refresh Timeout: Maximum time to wait for each file refresh operation before timeout

Note: Settings take effect immediately after saving"""
        
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
        """Refresh file list"""
        # Clear old list
        self.files_tree.delete(*self.files_tree.get_children())
        
        # Add new files
        for file_info in self.main_gui.config_manager.excel_files:
            name = os.path.basename(file_info["path"])
            path = file_info["path"]
            
            # Check file status
            if os.path.exists(path) and path.endswith(('.xlsx', '.xls')):
                status = "✓ Ready"
            else:
                status = "✗ Missing"
            
            self.files_tree.insert("", "end", values=(name, path, status))
    
    def add_file(self):
        """Add new file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=os.getcwd()
        )
        
        if file_path:
            # Convert to relative path if possible
            try:
                rel_path = os.path.relpath(file_path, os.getcwd())
                if not rel_path.startswith('..'):
                    file_path = rel_path
            except ValueError:
                # Use absolute path if relative path cannot be created
                pass
            
            # Check for duplicate files
            for file_info in self.main_gui.config_manager.excel_files:
                if file_info["path"] == file_path:
                    messagebox.showwarning("Duplicate File", "This file already exists in the list")
                    return
            
            # Add new file
            new_file = {"path": file_path}
            self.main_gui.config_manager.excel_files.append(new_file)
            
            # Refresh list
            self.refresh_files_list()
    
    def remove_selected_file(self):
        """Remove selected file"""
        selected_items = self.files_tree.selection()
        
        if not selected_items:
            messagebox.showwarning("No File Selected", "Please select a file to remove")
            return
        
        # Confirm removal
        result = messagebox.askyesno("Confirm Removal", 
                                   f"Do you want to remove {len(selected_items)} file(s) from the list?")
        
        if result:
            # Get paths of selected files
            selected_paths = []
            for item in selected_items:
                values = self.files_tree.item(item, "values")
                selected_paths.append(values[1])  # path is in column 1
            
            # Remove from config using the new setter
            filtered_files = [
                file_info for file_info in self.main_gui.config_manager.excel_files
                if file_info["path"] not in selected_paths
            ]
            self.main_gui.config_manager.excel_files = filtered_files
            
            # Refresh list
            self.refresh_files_list()
    
    def edit_selected_file(self):
        """Edit selected file"""
        selected_items = self.files_tree.selection()
        
        if not selected_items:
            messagebox.showwarning("No File Selected", "Please select a file to edit")
            return
        
        if len(selected_items) > 1:
            messagebox.showwarning("Too Many Files Selected", "Please select only 1 file")
            return
        
        # Get current file data
        item = selected_items[0]
        values = self.files_tree.item(item, "values")
        current_path = values[1]
        
        # Open dialog for new file selection
        new_path = filedialog.askopenfilename(
            title="Select New Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialfile=os.path.basename(current_path),
            initialdir=os.path.dirname(current_path) if os.path.isabs(current_path) else os.getcwd()
        )
        
        if new_path:
            # Convert to relative path if possible
            try:
                rel_path = os.path.relpath(new_path, os.getcwd())
                if not rel_path.startswith('..'):
                    new_path = rel_path
            except ValueError:
                pass
            
            # Update in config
            for file_info in self.main_gui.config_manager.excel_files:
                if file_info["path"] == current_path:
                    file_info["path"] = new_path
                    break
            
            # Refresh list
            self.refresh_files_list()
    
    def save_settings(self):
        """Save settings"""
        try:
            # Update general settings
            self.main_gui.config_manager.settings["auto_save"] = self.auto_save_var.get()
            self.main_gui.config_manager.settings["backup_before_refresh"] = self.backup_before_refresh_var.get()
            self.main_gui.config_manager.settings["log_refresh_activity"] = self.log_refresh_activity_var.get()
            self.main_gui.config_manager.settings["refresh_timeout_minutes"] = self.refresh_timeout_var.get()
            
            # Save to file
            self.main_gui.config_manager.save_config()
            
            messagebox.showinfo("Save Successful", "Settings saved successfully")
            
            # Refresh main window
            if self.refresh_callback:
                self.refresh_callback()
            
            # Close window
            self.on_closing()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {e}")
    
    def on_closing(self):
        """เมื่อปิดหน้าต่าง"""
        self.window.grab_release()
        self.window.destroy()
