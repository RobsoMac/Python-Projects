import pyautogui
import time
from PIL import Image, ImageDraw, ImageFont
import sys
import keyboard
import logging
from datetime import datetime
from pystray import Icon, MenuItem, Menu
from PIL import Image as PILImage
import threading
import winreg
import os
import threading
import ctypes
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import shutil
import json

# Global states
HOTKEY_ENABLED = True
BASE_URL = "https://your_url/"

# Define log file path (avoid dynamic timestamps in the file name for executables)
log_path = os.path.join(os.getenv('APPDATA'), 'RiskAssessmentTool', 'ra_autoclicker.log')
if not os.path.exists(os.path.dirname(log_path)):
    os.makedirs(os.path.dirname(log_path))

# Configure logging with rotation to avoid excessive files
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_path),
        logging.StreamHandler()  # Optionally also output to console
    ]
)

# Function to handle when the tray icon is clicked
def on_quit(icon, item):
    logging.info("Exiting the application...")
    icon.stop()

# Create a system tray icon with the hidden menu
def create_tray_icon():
    icon = Icon("test", create_image(), menu=Menu(MenuItem('Exit', on_quit)))
    icon.run()

# Run the tray icon in a separate thread
def run_tray_icon():
    threading.Thread(target=create_tray_icon, daemon=True).start()

# Main application logic
def main():
    logging.info("Risk Assessment Tool is starting in the system tray...")
    run_tray_icon()

    # Your existing application logic goes here...

    # For testing, we can simulate some actions
    logging.info("Application logic is running in the background.")

if __name__ == "__main__":

# Job Titles List
JOB_TITLES = [
    "Technician", "Audit Manager", "Cabling Technician", "Construction Manager", "Construction Safety Specialist","Security Manager",
    "Engineering Operations","Facility Manager", "H&S Program Manager", "Industrial Hygienist"
]

# Setup GUI
class SetupGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Risk Assessment Tool Setup")
        self.root.geometry("600x400")
        self.root.resizable(False, False)

        self.load_last_job_title()
        self.style = ttk.Style()
        self.style.configure('Custom.TButton', padding=10, font=('Segoe UI', 10))

        self.setup_ui()

    def load_last_job_title(self):
        """Load the last job title from settings.json, or set default"""
        try:
            with open('settings.json', 'r') as f:
                settings = json.load(f)
                self.last_job_title = settings.get('job_title', 'Technician')
        except FileNotFoundError:
            self.last_job_title = 'Technician'

    def save_job_title(self):
        """Save the selected job title to settings.json"""
        with open('settings.json', 'w') as f:
            json.dump({'job_title': self.role_var.get()}, f)

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header
        header_label = ttk.Label(
            main_frame,
            text="Risk Assessment Tool Setup",
            font=('Segoe UI', 16, 'bold')
        )
        header_label.pack(pady=20)

        # Author info
        author_label = ttk.Label(
            main_frame,
            text="Created by: macrobso@",  # Removed from the GUI, keeping in the file metadata
            font=('Segoe UI', 10, 'italic')
        )
        author_label.pack()

        # Job title dropdown
        role_label = ttk.Label(main_frame, text="Select Job Title:")
        role_label.pack(pady=5)

        self.role_var = tk.StringVar(value=self.last_job_title)
        role_combobox = ttk.Combobox(
            main_frame, textvariable=self.role_var, values=JOB_TITLES, state="readonly"
        )
        role_combobox.pack(pady=5)

        # Save button to remember job title
        save_button = ttk.Button(main_frame, text="Save", command=self.save_job_title)
        save_button.pack(pady=10)

        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)

        # Install button
        self.install_button = ttk.Button(
            button_frame,
            text="Install Tool",
            style='Custom.TButton',
            command=self.install_tool
        )
        self.install_button.pack(side=tk.LEFT, padx=10)

        # Run button
        self.run_button = ttk.Button(
            button_frame,
            text="Run Tool",
            style='Custom.TButton',
            command=self.run_tool
        )
        self.run_button.pack(side=tk.LEFT, padx=10)

        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=20)

        # Startup option
        self.startup_var = tk.BooleanVar(value=True)
        startup_check = ttk.Checkbutton(
            options_frame,
            text="Start with Windows",
            variable=self.startup_var
        )
        startup_check.pack()

    def install_tool(self):
        try:
            dest_path = os.path.join(os.getenv('APPDATA'), 'RiskAssessmentTool')
            os.makedirs(dest_path, exist_ok=True)

            # Copy executable
            if getattr(sys, 'frozen', False):
                current_exe = sys.executable
                shutil.copy2(current_exe, os.path.join(dest_path, 'risk_assessment_tool.exe'))

            # Create images folder
            os.makedirs(os.path.join(dest_path, 'images'), exist_ok=True)

            # Add to startup if selected
            if self.startup_var.get():
                self.add_to_startup(dest_path)

            messagebox.showinfo("Success", 
                              "Tool installed successfully!\n\n"
                              "You can now use F8 to trigger the automation.")
            
            self.run_tool()

        except Exception as e:
            messagebox.showerror("Error", f"Installation failed: {str(e)}")

    def add_to_startup(self, install_path):
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            exe_path = os.path.join(install_path, 'risk_assessment_tool.exe')

            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, "RiskAssessmentTool", 0, winreg.REG_SZ, exe_path)
            winreg.CloseKey(key)
        except Exception as e:
            messagebox.showwarning("Warning", f"Could not add to startup: {str(e)}")

    def run_tool(self):
        self.root.destroy()
        start_main_tool()

# Main Tool Logic
def start_main_tool():
    """Original tool functionality"""
    hide_console()
    add_to_startup()
    keyboard.add_hotkey('F8', handle_hotkey)
    icon = create_system_tray()
    icon.run()

def hide_console():
    """Hide console window on Windows"""
    if sys.platform == 'win32':
        ctypes.windll.kernel32.FreeConsole()

def create_system_tray():
    """Create system tray icon using pystray"""
    icon_path = os.path.join(os.path.dirname(sys.executable), "images", "arrowhslogo.ico")
    icon_image = PILImage.open(icon_path)
    icon = pystray.Icon("Risk Assessment Tool", icon_image, "Risk Assessment Tool")
    return icon

def main():
    if getattr(sys, 'frozen', False):
        app_path = os.path.join(os.getenv('APPDATA'), 'RiskAssessmentTool')
        if not os.path.exists(app_path):
            setup = SetupGUI()
            setup.root.mainloop()
        else:
            start_main_tool()
    else:
        setup = SetupGUI()
        setup.root.mainloop()

if __name__ == "__main__":
    if not os.path.exists('images'):
        os.makedirs('images')
        logging.info('Created images directory')

    main()
