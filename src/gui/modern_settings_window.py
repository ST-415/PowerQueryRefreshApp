"""
Modern Settings Window with CustomTkinter
Beautiful and Modern Settings Interface
"""

import customtkinter as ctk
from tkinter import messagebox, filedialog
import os
from typing import Any, Dict

class ModernSettingsWindow:
    """Modern settings window using CustomTkinter"""
    
    def __init__(self, parent, config_manager):
        self.parent = parent
        self.config_manager = config_manager
        
        # Create settings window
        self.window = ctk.CTkToplevel(parent)
        self.window.title("Settings")
        self.window.geometry("600x500")
        self.window.resizable(False, False)
        
        # Make window modal
        self.window.transient(parent)
        self.window.grab_set()
        
        # Center window
        self.center_window()
        
        # Load current settings
        self.load_settings()
        
        # Create widgets
        self.create_widgets()
        
        # Handle window closing
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def center_window(self):
        """Center the window on screen"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')
    
    def load_settings(self):
        """Load current settings from config manager"""
        try:
            self.settings = self.config_manager.get_config()
        except:
            self.settings = {
                'data_folder': 'data',
                'backup_folder': 'data/backups',
                'log_folder': 'data/logs',
                'auto_backup': True,
                'refresh_timeout': 300
            }
    
    def create_widgets(self):
        """Create all widgets with modern design"""
        # Main frame
        main_frame = ctk.CTkFrame(self.window, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title_label = ctk.CTkLabel(
            main_frame,
            text="⚙️ Settings",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=("#1f538d", "#4a90e2")
        )
        title_label.pack(pady=(0, 20))
        
        # Tabview for different setting categories
        self.tabview = ctk.CTkTabview(main_frame, width=550, height=350)
        self.tabview.pack(fill="both", expand=True, pady=(0, 20))
        
        # Create tabs
        self.create_folders_tab()
        self.create_backup_tab()
        self.create_refresh_tab()
        
        # Button frame
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x")
        
        # Cancel button
        cancel_button = ctk.CTkButton(
            button_frame,
            text="❌ Cancel",
            command=self.on_closing,
            width=100,
            height=40,
            corner_radius=20,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#95a5a6",
            hover_color="#7f8c8d"
        )
        cancel_button.pack(side="right", padx=(10, 0))
        
        # Save button
        save_button = ctk.CTkButton(
            button_frame,
            text="💾 Save",
            command=self.save_settings,
            width=100,
            height=40,
            corner_radius=20,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#27ae60",
            hover_color="#219a52"
        )
        save_button.pack(side="right", padx=(0, 10))
        
        # Reset button
        reset_button = ctk.CTkButton(
            button_frame,
            text="🔄 Reset",
            command=self.reset_settings,
            width=100,
            height=40,
            corner_radius=20,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#e74c3c",
            hover_color="#c0392b"
        )
        reset_button.pack(side="right", padx=(0, 10))
    
    def create_folders_tab(self):
        """Create folders settings tab"""
        folder_tab = self.tabview.add("📁 Folders")
        
        # Data folder setting
        data_frame = ctk.CTkFrame(folder_tab, corner_radius=10)
        data_frame.pack(fill="x", pady=10, padx=20)
        
        data_label = ctk.CTkLabel(
            data_frame,
            text="📂 Data Folder",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        data_label.pack(pady=(15, 5), anchor="w", padx=20)
        
        data_desc = ctk.CTkLabel(
            data_frame,
            text="Folder containing Excel files to refresh",
            font=ctk.CTkFont(size=12),
            text_color=("#7f8c8d", "#bdc3c7")
        )
        data_desc.pack(anchor="w", padx=20)
        
        data_entry_frame = ctk.CTkFrame(data_frame, fg_color="transparent")
        data_entry_frame.pack(fill="x", pady=(5, 15), padx=20)
        
        self.data_folder_entry = ctk.CTkEntry(
            data_entry_frame,
            placeholder_text="Enter data folder path...",
            height=35,
            corner_radius=10
        )
        self.data_folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.data_folder_entry.insert(0, self.settings.get('data_folder', 'data'))
        
        data_browse_button = ctk.CTkButton(
            data_entry_frame,
            text="📁 Browse",
            command=lambda: self.browse_folder(self.data_folder_entry),
            width=80,
            height=35,
            corner_radius=10
        )
        data_browse_button.pack(side="right")
        
        # Backup folder setting
        backup_frame = ctk.CTkFrame(folder_tab, corner_radius=10)
        backup_frame.pack(fill="x", pady=10, padx=20)
        
        backup_label = ctk.CTkLabel(
            backup_frame,
            text="💾 Backup Folder",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        backup_label.pack(pady=(15, 5), anchor="w", padx=20)
        
        backup_desc = ctk.CTkLabel(
            backup_frame,
            text="Folder to store backup files",
            font=ctk.CTkFont(size=12),
            text_color=("#7f8c8d", "#bdc3c7")
        )
        backup_desc.pack(anchor="w", padx=20)
        
        backup_entry_frame = ctk.CTkFrame(backup_frame, fg_color="transparent")
        backup_entry_frame.pack(fill="x", pady=(5, 15), padx=20)
        
        self.backup_folder_entry = ctk.CTkEntry(
            backup_entry_frame,
            placeholder_text="Enter backup folder path...",
            height=35,
            corner_radius=10
        )
        self.backup_folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.backup_folder_entry.insert(0, self.settings.get('backup_folder', 'data/backups'))
        
        backup_browse_button = ctk.CTkButton(
            backup_entry_frame,
            text="📁 Browse",
            command=lambda: self.browse_folder(self.backup_folder_entry),
            width=80,
            height=35,
            corner_radius=10
        )
        backup_browse_button.pack(side="right")
        
        # Log folder setting
        log_frame = ctk.CTkFrame(folder_tab, corner_radius=10)
        log_frame.pack(fill="x", pady=10, padx=20)
        
        log_label = ctk.CTkLabel(
            log_frame,
            text="📜 Log Folder",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        log_label.pack(pady=(15, 5), anchor="w", padx=20)
        
        log_desc = ctk.CTkLabel(
            log_frame,
            text="Folder to store log files",
            font=ctk.CTkFont(size=12),
            text_color=("#7f8c8d", "#bdc3c7")
        )
        log_desc.pack(anchor="w", padx=20)
        
        log_entry_frame = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_entry_frame.pack(fill="x", pady=(5, 15), padx=20)
        
        self.log_folder_entry = ctk.CTkEntry(
            log_entry_frame,
            placeholder_text="Enter log folder path...",
            height=35,
            corner_radius=10
        )
        self.log_folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.log_folder_entry.insert(0, self.settings.get('log_folder', 'data/logs'))
        
        log_browse_button = ctk.CTkButton(
            log_entry_frame,
            text="📁 Browse",
            command=lambda: self.browse_folder(self.log_folder_entry),
            width=80,
            height=35,
            corner_radius=10
        )
        log_browse_button.pack(side="right")
    
    def create_backup_tab(self):
        """Create backup settings tab"""
        backup_tab = self.tabview.add("💾 Backup")
        
        # Auto backup setting
        auto_backup_frame = ctk.CTkFrame(backup_tab, corner_radius=10)
        auto_backup_frame.pack(fill="x", pady=20, padx=20)
        
        auto_backup_label = ctk.CTkLabel(
            auto_backup_frame,
            text="🔄 Auto Backup",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        auto_backup_label.pack(pady=(15, 5), anchor="w", padx=20)
        
        auto_backup_desc = ctk.CTkLabel(
            auto_backup_frame,
            text="Automatically create backups before refreshing files",
            font=ctk.CTkFont(size=12),
            text_color=("#7f8c8d", "#bdc3c7")
        )
        auto_backup_desc.pack(anchor="w", padx=20)
        
        self.auto_backup_switch = ctk.CTkSwitch(
            auto_backup_frame,
            text="Enable automatic backup",
            font=ctk.CTkFont(size=14),
            onvalue=True,
            offvalue=False
        )
        self.auto_backup_switch.pack(pady=(10, 15), padx=20, anchor="w")
        
        # Set initial value
        if self.settings.get('auto_backup', True):
            self.auto_backup_switch.select()
        
        # Backup info frame
        info_frame = ctk.CTkFrame(backup_tab, corner_radius=10)
        info_frame.pack(fill="x", pady=10, padx=20)
        
        info_label = ctk.CTkLabel(
            info_frame,
            text="ℹ️ Backup Information",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        info_label.pack(pady=(15, 10), anchor="w", padx=20)
        
        info_text = ctk.CTkTextbox(
            info_frame,
            height=100,
            corner_radius=10,
            font=ctk.CTkFont(size=12)
        )
        info_text.pack(fill="x", pady=(0, 15), padx=20)
        
        info_content = """• Backups are created with timestamps
• Original files are preserved
• Backups are stored in the backup folder
• You can restore from backups manually
• Clean up old backups regularly to save space"""
        
        info_text.insert("0.0", info_content)
        info_text.configure(state="disabled")
    
    def create_refresh_tab(self):
        """Create refresh settings tab"""
        refresh_tab = self.tabview.add("🔄 Refresh")
        
        # Timeout setting
        timeout_frame = ctk.CTkFrame(refresh_tab, corner_radius=10)
        timeout_frame.pack(fill="x", pady=20, padx=20)
        
        timeout_label = ctk.CTkLabel(
            timeout_frame,
            text="⏱️ Refresh Timeout",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        timeout_label.pack(pady=(15, 5), anchor="w", padx=20)
        
        timeout_desc = ctk.CTkLabel(
            timeout_frame,
            text="Maximum time to wait for refresh (seconds)",
            font=ctk.CTkFont(size=12),
            text_color=("#7f8c8d", "#bdc3c7")
        )
        timeout_desc.pack(anchor="w", padx=20)
        
        self.timeout_entry = ctk.CTkEntry(
            timeout_frame,
            placeholder_text="Enter timeout in seconds...",
            height=35,
            corner_radius=10,
            width=200
        )
        self.timeout_entry.pack(pady=(10, 15), padx=20, anchor="w")
        self.timeout_entry.insert(0, str(self.settings.get('refresh_timeout', 300)))
        
        # Performance tips
        tips_frame = ctk.CTkFrame(refresh_tab, corner_radius=10)
        tips_frame.pack(fill="x", pady=10, padx=20)
        
        tips_label = ctk.CTkLabel(
            tips_frame,
            text="💡 Performance Tips",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        tips_label.pack(pady=(15, 10), anchor="w", padx=20)
        
        tips_text = ctk.CTkTextbox(
            tips_frame,
            height=120,
            corner_radius=10,
            font=ctk.CTkFont(size=12)
        )
        tips_text.pack(fill="x", pady=(0, 15), padx=20)
        
        tips_content = """• Close other Excel applications before refreshing
• Use faster storage (SSD) for better performance
• Refresh files in smaller batches
• Increase timeout for large files
• Check network connectivity for external data sources
• Keep Excel and Power Query updated"""
        
        tips_text.insert("0.0", tips_content)
        tips_text.configure(state="disabled")
    
    def browse_folder(self, entry_widget):
        """Browse for folder"""
        folder = filedialog.askdirectory(
            title="Select Folder",
            initialdir=entry_widget.get()
        )
        if folder:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, folder)
    
    def save_settings(self):
        """Save settings"""
        try:
            # Validate timeout
            timeout_str = self.timeout_entry.get()
            try:
                timeout = int(timeout_str)
                if timeout <= 0:
                    raise ValueError("Timeout must be positive")
            except ValueError:
                messagebox.showerror("Error", "Invalid timeout value. Please enter a positive number.")
                return
            
            # Prepare settings
            new_settings = {
                'data_folder': self.data_folder_entry.get(),
                'backup_folder': self.backup_folder_entry.get(),
                'log_folder': self.log_folder_entry.get(),
                'auto_backup': self.auto_backup_switch.get(),
                'refresh_timeout': timeout
            }
            
            # Save to config manager
            self.config_manager.update_config(new_settings)
            
            # Show success message
            messagebox.showinfo("Success", "Settings saved successfully!")
            
            # Close window
            self.window.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
    
    def reset_settings(self):
        """Reset settings to default"""
        if messagebox.askyesno("Confirm Reset", "Are you sure you want to reset all settings to default values?"):
            # Reset to default values
            self.data_folder_entry.delete(0, "end")
            self.data_folder_entry.insert(0, "data")
            
            self.backup_folder_entry.delete(0, "end")
            self.backup_folder_entry.insert(0, "data/backups")
            
            self.log_folder_entry.delete(0, "end")
            self.log_folder_entry.insert(0, "data/logs")
            
            self.auto_backup_switch.select()
            
            self.timeout_entry.delete(0, "end")
            self.timeout_entry.insert(0, "300")
    
    def on_closing(self):
        """Handle window closing"""
        self.window.destroy()

# Legacy wrapper for compatibility
SettingsWindow = ModernSettingsWindow
