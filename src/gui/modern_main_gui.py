"""
Modern Main GUI Application with CustomTkinter
PowerQuery Refresh Application - Beautiful Modern UI
"""

import customtkinter as ctk
from tkinter import messagebox, filedialog
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
from gui.modern_settings_window import ModernSettingsWindow

# Set appearance mode and color theme
ctk.set_appearance_mode("light")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class ModernMainGUI:
    """Modern main application window using CustomTkinter"""
    
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("PowerQuery Refresh Tool")
        self.root.geometry("900x800")
        self.root.resizable(True, True)
        
        # Initialize managers
        self.config_manager = ConfigManager()
        self.logger_manager = LoggerManager()
        self.file_manager = FileManager()
        self.excel_refresher = ExcelRefresher(self.logger_manager, self.file_manager)
        
        # Variables for file management
        self.file_vars = []  # เก็บ BooleanVar สำหรับแต่ละไฟล์
        self.file_info = []  # เก็บข้อมูลไฟล์
        
        # Threading variables
        self.refresh_thread = None
        self.stop_refresh_flag = False
        
        # Create widgets
        self.create_widgets()
        self.refresh_file_list()
        
        # Setup window closing protocol
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def create_widgets(self):
        """Create all widgets with modern design"""
        # Main frame with padding
        main_frame = ctk.CTkFrame(self.root, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title section
        title_frame = ctk.CTkFrame(main_frame, corner_radius=15, height=80)
        title_frame.pack(fill="x", pady=(0, 20))
        title_frame.pack_propagate(False)
        
        title_label = ctk.CTkLabel(
            title_frame, 
            text="🔄 PowerQuery Refresh Tool",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=("#1f538d", "#4a90e2")
        )
        title_label.pack(pady=20)
        
        # Control buttons section
        button_frame = ctk.CTkFrame(main_frame, corner_radius=15, height=80)
        button_frame.pack(fill="x", pady=(0, 20))
        button_frame.pack_propagate(False)
        
        # Inner frame for buttons
        inner_button_frame = ctk.CTkFrame(button_frame, fg_color="transparent")
        inner_button_frame.pack(pady=15)
        
        # Settings button
        self.settings_button = ctk.CTkButton(
            inner_button_frame,
            text="⚙️ Settings",
            command=self.open_settings,
            width=120,
            height=40,
            corner_radius=20,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.settings_button.pack(side="left", padx=10)
        
        # Refresh list button
        self.refresh_button = ctk.CTkButton(
            inner_button_frame,
            text="🔄 Refresh List",
            command=self.refresh_file_list,
            width=120,
            height=40,
            corner_radius=20,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.refresh_button.pack(side="left", padx=10)
        
        # Select All button
        self.select_all_button = ctk.CTkButton(
            inner_button_frame,
            text="✅ Select All",
            command=self.select_all_files,
            width=120,
            height=40,
            corner_radius=20,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#27ae60",
            hover_color="#219a52"
        )
        self.select_all_button.pack(side="left", padx=10)
        
        # Deselect All button
        self.deselect_all_button = ctk.CTkButton(
            inner_button_frame,
            text="❌ Deselect All",
            command=self.deselect_all_files,
            width=120,
            height=40,
            corner_radius=20,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#e74c3c",
            hover_color="#c0392b"
        )
        self.deselect_all_button.pack(side="left", padx=10)
        
        # File list section
        file_frame = ctk.CTkFrame(main_frame, corner_radius=15)
        file_frame.pack(fill="both", expand=True, pady=(0, 20))
        
        # File list title
        file_title_label = ctk.CTkLabel(
            file_frame,
            text="📁 Excel Files",
            font=ctk.CTkFont(size=18, weight="bold"),
            anchor="w"
        )
        file_title_label.pack(pady=(15, 10), padx=20, anchor="w")
        
        # Scrollable frame for file list
        self.file_scroll_frame = ctk.CTkScrollableFrame(
            file_frame,
            corner_radius=10,
            height=300
        )
        self.file_scroll_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # Action buttons section
        action_frame = ctk.CTkFrame(main_frame, corner_radius=15, height=80)
        action_frame.pack(fill="x")
        action_frame.pack_propagate(False)
        
        # Inner frame for action buttons
        inner_action_frame = ctk.CTkFrame(action_frame, fg_color="transparent")
        inner_action_frame.pack(pady=15)
        
        # Start refresh button
        self.start_button = ctk.CTkButton(
            inner_action_frame,
            text="🚀 Start Refresh",
            command=self.start_refresh,
            width=150,
            height=50,
            corner_radius=25,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#e67e22",
            hover_color="#d35400"
        )
        self.start_button.pack(side="left", padx=15)
        
        # Stop refresh button
        self.stop_button = ctk.CTkButton(
            inner_action_frame,
            text="⏹️ Stop Refresh",
            command=self.stop_refresh,
            width=150,
            height=50,
            corner_radius=25,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#95a5a6",
            hover_color="#7f8c8d",
            state="disabled"
        )
        self.stop_button.pack(side="left", padx=15)
        
        # Progress bar
        self.progress = ctk.CTkProgressBar(
            inner_action_frame,
            width=200,
            height=20,
            corner_radius=10
        )
        self.progress.pack(side="left", padx=15)
        self.progress.set(0)
        
        # Status label
        self.status_label = ctk.CTkLabel(
            inner_action_frame,
            text="Ready",
            font=ctk.CTkFont(size=14),
            text_color=("#2c3e50", "#ecf0f1")
        )
        self.status_label.pack(side="left", padx=15)
        
        # Log viewer button
        self.log_viewer_button = ctk.CTkButton(
            main_frame,
            text="📜 View Logs",
            command=self.show_logs,
            width=120,
            height=40,
            corner_radius=20,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2980b9",
            hover_color="#2471a3"
        )
        self.log_viewer_button.pack(side="left", padx=10, pady=(10, 0))
        
        # About button
        self.about_button = ctk.CTkButton(
            main_frame,
            text="ℹ️ About",
            command=self.show_about,
            width=120,
            height=40,
            corner_radius=20,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#8e44ad",
            hover_color="#732d91"
        )
        self.about_button.pack(side="left", padx=10, pady=(10, 0))
    
    def refresh_file_list(self):
        """Refresh the file list with modern cards"""
        # Clear existing widgets
        for widget in self.file_scroll_frame.winfo_children():
            widget.destroy()
        
        # Clear variables
        self.file_vars.clear()
        self.file_info.clear()
        
        # Get file list from file manager
        try:
            files = self.file_manager.get_excel_files()
            
            if not files:
                # Show empty state
                empty_frame = ctk.CTkFrame(self.file_scroll_frame, corner_radius=10)
                empty_frame.pack(fill="x", pady=10)
                
                empty_label = ctk.CTkLabel(
                    empty_frame,
                    text="📂 No Excel files found\nPlace your Excel files in the data folder",
                    font=ctk.CTkFont(size=14),
                    text_color=("#7f8c8d", "#bdc3c7")
                )
                empty_label.pack(pady=30)
                return
            
            # Create file cards
            for i, file_path in enumerate(files):
                self.create_file_card(i, file_path)
                
        except Exception as e:
            self.show_error(f"Error loading files: {str(e)}")
    
    def create_file_card(self, index: int, file_path: str):
        """Create a modern file card"""
        # Create boolean variable for checkbox
        var = ctk.BooleanVar()
        self.file_vars.append(var)
        
        # Store file info
        file_name = os.path.basename(file_path)
        file_size = self.get_file_size(file_path)
        self.file_info.append({
            'path': file_path,
            'name': file_name,
            'size': file_size
        })
        
        # Create card frame
        card_frame = ctk.CTkFrame(self.file_scroll_frame, corner_radius=10)
        card_frame.pack(fill="x", pady=5, padx=5)
        
        # Inner frame for content
        content_frame = ctk.CTkFrame(card_frame, fg_color="transparent")
        content_frame.pack(fill="x", padx=15, pady=10)
        
        # Checkbox
        checkbox = ctk.CTkCheckBox(
            content_frame,
            text="",
            variable=var,
            width=20,
            height=20,
            corner_radius=5
        )
        checkbox.pack(side="left", padx=(0, 15))
        
        # File info frame
        info_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        info_frame.pack(side="left", fill="x", expand=True)
        
        # File name
        name_label = ctk.CTkLabel(
            info_frame,
            text=f"📄 {file_name}",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w"
        )
        name_label.pack(anchor="w")
        
        # File details
        details_label = ctk.CTkLabel(
            info_frame,
            text=f"Size: {file_size} | Path: {file_path}",
            font=ctk.CTkFont(size=12),
            text_color=("#7f8c8d", "#bdc3c7"),
            anchor="w"
        )
        details_label.pack(anchor="w")
    
    def get_file_size(self, file_path: str) -> str:
        """Get human-readable file size"""
        try:
            size = os.path.getsize(file_path)
            for unit in ['B', 'KB', 'MB', 'GB']:
                if size < 1024.0:
                    return f"{size:.1f} {unit}"
                size /= 1024.0
            return f"{size:.1f} TB"
        except:
            return "Unknown"
    
    def select_all_files(self):
        """Select all files"""
        for var in self.file_vars:
            var.set(True)
    
    def deselect_all_files(self):
        """Deselect all files"""
        for var in self.file_vars:
            var.set(False)
    
    def start_refresh(self):
        """Start the refresh process"""
        # Get selected files
        selected_files = []
        for i, var in enumerate(self.file_vars):
            if var.get():
                selected_files.append(self.file_info[i]['path'])
        
        if not selected_files:
            self.show_warning("Please select at least one file to refresh.")
            return
        
        # Update UI state
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.progress.set(0)
        self.status_label.configure(text="Starting refresh...")
        
        # Start refresh in separate thread
        self.refresh_thread = threading.Thread(
            target=self.refresh_worker,
            args=(selected_files,),
            daemon=True
        )
        self.refresh_thread.start()
    
    def refresh_worker(self, selected_files: List[str]):
        """Worker thread for refresh process"""
        try:
            total_files = len(selected_files)
            self.stop_refresh_flag = False
            
            for i, file_path in enumerate(selected_files):
                # Check if stop was requested
                if self.stop_refresh_flag:
                    break
                    
                # Update progress
                progress = (i + 1) / total_files
                self.root.after(0, lambda p=progress: self.progress.set(p))
                
                # Update status
                file_name = os.path.basename(file_path)
                self.root.after(0, lambda n=file_name: self.status_label.configure(
                    text=f"Refreshing: {n}"))
                
                # Refresh the file
                self.excel_refresher.refresh_file(file_path)
                
            # Check if stopped or completed
            if self.stop_refresh_flag:
                self.root.after(0, lambda: self.status_label.configure(text="Refresh stopped"))
            else:
                # Refresh completed
                self.root.after(0, self.refresh_completed)
            
        except Exception as e:
            self.root.after(0, lambda: self.show_error(f"Refresh error: {str(e)}"))
            self.root.after(0, self.refresh_completed)
    
    def refresh_completed(self):
        """Handle refresh completion"""
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.progress.set(1.0)
        self.status_label.configure(text="Refresh completed!")
        self.show_success("Refresh completed successfully!")
    
    def stop_refresh(self):
        """Stop the refresh process"""
        self.stop_refresh_flag = True
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.progress.set(0)
        self.status_label.configure(text="Refresh stopped")
        self.show_info("Refresh process stopped.")
    
    def open_settings(self):
        """Open settings window"""
        try:
            settings_window = ModernSettingsWindow(self.root, self.config_manager)
        except Exception as e:
            self.show_error(f"Error opening settings: {str(e)}")
    
    def show_logs(self):
        """Show log viewer window"""
        try:
            log_window = LogViewerWindow(self.root, self.logger_manager)
        except Exception as e:
            self.show_error(f"Error opening log viewer: {str(e)}")
    
    def show_about(self):
        """Show about dialog"""
        about_text = """
PowerQuery Refresh Tool v2.0
Modern GUI Application

Features:
• Beautiful modern interface using CustomTkinter
• Batch refresh multiple Excel files
• Automatic backup before refresh
• Real-time progress tracking
• Comprehensive logging
• Flexible configuration

Created with ❤️ by PowerQuery Team
        """
        messagebox.showinfo("About", about_text)
    
    def add_files(self):
        """Add files to the list"""
        files = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if files:
            # Copy files to data folder
            data_folder = self.config_manager.get_config().get('data_folder', 'data')
            
            for file_path in files:
                try:
                    import shutil
                    file_name = os.path.basename(file_path)
                    dest_path = os.path.join(data_folder, file_name)
                    shutil.copy2(file_path, dest_path)
                except Exception as e:
                    self.show_error(f"Error copying file {file_name}: {str(e)}")
            
            # Refresh the file list
            self.refresh_file_list()
            self.show_success("Files added successfully!")
    
    def remove_selected_files(self):
        """Remove selected files from the list"""
        selected_files = []
        for i, var in enumerate(self.file_vars):
            if var.get():
                selected_files.append(self.file_info[i])
        
        if not selected_files:
            self.show_warning("Please select files to remove.")
            return
        
        # Confirm deletion
        file_names = [info['name'] for info in selected_files]
        message = f"Are you sure you want to remove these files?\n\n" + "\n".join(file_names)
        
        if messagebox.askyesno("Confirm Removal", message):
            try:
                for file_info in selected_files:
                    if os.path.exists(file_info['path']):
                        os.remove(file_info['path'])
                
                # Refresh the file list
                self.refresh_file_list()
                self.show_success("Files removed successfully!")
                
            except Exception as e:
                self.show_error(f"Error removing files: {str(e)}")

    def show_success(self, message: str):
        """Show success message"""
        messagebox.showinfo("Success", message)
    
    def show_warning(self, message: str):
        """Show warning message"""
        messagebox.showwarning("Warning", message)
    
    def show_error(self, message: str):
        """Show error message"""
        messagebox.showerror("Error", message)
    
    def show_info(self, message: str):
        """Show info message"""
        messagebox.showinfo("Info", message)
    
    def on_closing(self):
        """Handle window closing"""
        self.root.destroy()
    
    def run(self):
        """Start the application"""
        self.root.mainloop()

def main():
    """Main function"""
    app = ModernMainGUI()
    app.run()

if __name__ == "__main__":
    main()
