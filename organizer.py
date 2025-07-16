import os
import shutil
import json
import logging
import csv
from datetime import datetime
import requests
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import time
import pystray
from PIL import Image
import winshell
import pythoncom
import win32com.client
import sys

# Configure logging
logging.basicConfig(
    filename='organizer.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Application version
APP_VERSION = "1.0.0"

def get_resource_path(relative_path):
    """Get the absolute path to a resource, handling PyInstaller bundled environment."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # Running as script, use current directory
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def load_config(config_file='categories.json'):
    """Load file extension categories, monitored folders, and settings from a JSON config file."""
    default_config = {
        "categories": {
            '.pdf': 'Documents',
            '.doc': 'Documents',
            '.docx': 'Documents',
            '.txt': 'Documents',
            '.jpg': 'Images',
            '.jpeg': 'Images',
            '.png': 'Images',
            '.gif': 'Images',
            '.mp4': 'Videos',
            '.mkv': 'Videos',
            '.avi': 'Videos',
            '.zip': 'Archives',
            '.rar': 'Archives',
            '.py': 'Code',
            '.java': 'Code',
            '.cpp': 'Code',
        },
        "monitored_folders": [
            "C:/Users/Dell/Downloads/file-organizer/test",
            "C:/Users/Dell/OneDrive/Desktop/test"
        ],
        "startup_enabled": False,
        "appearance_mode": "system",
        "organize_by_date": False,
        "folder_settings": {
            "C:/Users/Dell/Downloads/file-organizer/test": {"recursive": True, "exclusions": []},
            "C:/Users/Dell/OneDrive/Desktop/test": {"recursive": True, "exclusions": []}
        }
    }
    
    config_path = get_resource_path(config_file)
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                loaded_config = json.load(f)
            loaded_config['monitored_folders'] = [os.path.normpath(f) for f in loaded_config.get('monitored_folders', []) if os.path.isabs(f) and '\x0c' not in f]
            folder_settings = loaded_config.get('folder_settings', {})
            for folder in folder_settings:
                folder_settings[folder] = {
                    "recursive": folder_settings[folder].get("recursive", False),
                    "exclusions": [os.path.normpath(e) for e in folder_settings[folder].get("exclusions", []) if os.path.isabs(e)]
                }
            loaded_config['folder_settings'] = folder_settings
            default_config.update(loaded_config)
            logging.info("Loaded config from %s with %d monitored folders", config_path, len(loaded_config.get('monitored_folders', [])))
        else:
            logging.info("No config file found at %s, using default config", config_path)
            save_config(default_config, config_path)
        return default_config
    except Exception as e:
        logging.error("Error loading config file %s: %s", config_path, e)
        return default_config

def save_config(config, config_file='categories.json'):
    """Save the configuration to a JSON file."""
    config_path = get_resource_path(config_file)
    try:
        with open(config_file, 'w') as f:
            json.dump(config, f, indent=4)
        logging.info("Saved config to %s", config_path)
    except Exception as e:
        logging.error("Error saving config file %s: %s", config_path, e)

def get_category(file_name, categories):
    """Return the category folder for a given file based on its extension."""
    _, ext = os.path.splitext(file_name)
    ext = ext.lower()
    
    if ext in ('.tmp', '.download', '.crdownload', '.onetoc2', '.onecache'):
        logging.info("Skipped temporary/OneDrive file: %s", file_name)
        return None
    
    return categories.get(ext, 'Others')

def is_already_organized(file_path, base_folder, categories):
    """Check if a file is already in its correct category folder."""
    file_name = os.path.basename(file_path)
    expected_category = get_category(file_name, categories)
    if not expected_category:
        return True
    
    current_folder = os.path.normpath(os.path.normcase(os.path.dirname(file_path)))
    expected_folder = os.path.normpath(os.path.normcase(os.path.join(base_folder, expected_category)))
    base_folder_normalized = os.path.normpath(os.path.normcase(base_folder))
    
    is_correct = current_folder.startswith(expected_folder) and current_folder != base_folder_normalized
    if is_correct:
        logging.info("Skipped %s: already in correct folder %s", file_name, expected_category)
    else:
        logging.info("Processing %s: current=%s, expected=%s, root=%s", file_name, current_folder, expected_folder, base_folder_normalized)
    return is_correct

def organize_file(file_path, base_folder, categories, organize_by_date=False):
    """Move a file to its category folder, optionally by date, creating folders if needed."""
    try:
        file_name = os.path.basename(file_path)
        
        if is_already_organized(file_path, base_folder, categories):
            return False, f"Skipped {file_name}: already in correct folder {get_category(file_name, categories)}"
        
        category = get_category(file_name, categories)
        if not category:
            return False, f"Skipped {file_name}: temporary or unsupported file type"
        
        target_folder = os.path.join(base_folder, category)
        if organize_by_date:
            try:
                timestamp = os.path.getctime(file_path)
                date_str = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d')
                target_folder = os.path.join(target_folder, date_str)
            except Exception as e:
                logging.error("Error getting timestamp for %s: %s", file_name, e)
                target_folder = os.path.join(base_folder, category)
        
        if not os.access(os.path.dirname(file_path), os.W_OK) or not os.access(base_folder, os.W_OK):
            raise PermissionError("No write access to source or destination folder")
        
        os.makedirs(target_folder, exist_ok=True)
        
        target_path = os.path.join(target_folder, file_name)
        base_name, ext = os.path.splitext(file_name)
        counter = 1
        while os.path.exists(target_path):
            new_file_name = f"{base_name}_{counter}{ext}"
            target_path = os.path.join(target_folder, new_file_name)
            counter += 1
        
        shutil.move(file_path, target_path)
        logging.info("Moved %s to %s", file_name, target_path)
        return True, f"Moved {file_name} to {category}{'/' + date_str if organize_by_date else ''}"
    except PermissionError as e:
        logging.error("Permission error moving %s: %s", file_name, e)
        return False, f"Permission error moving {file_name}: {str(e)}. Try running as administrator."
    except FileNotFoundError as e:
        logging.error("File not found for %s: %s", file_name, e)
        return False, f"File not found for {file_name}: {str(e)}"
    except Exception as e:
        logging.error("Error moving %s: %s", file_name, e)
        return False, f"Error moving {file_name}: {str(e)}"

class FileOrganizerHandler(FileSystemEventHandler):
    """Handle file system events to organize new or renamed files."""
    def __init__(self, base_folder, categories, log_callback, recursive=False, exclusions=None, organize_by_date=False):
        self.base_folder = base_folder
        self.categories = categories
        self.log_callback = log_callback
        self.is_running = True
        self.is_paused = False
        self.recursive = recursive
        self.exclusions = exclusions or []
        self.organize_by_date = organize_by_date
        self.recent_deletions = {}
        logging.debug("Initialized handler for %s (recursive=%s, exclusions=%s, organize_by_date=%s)", base_folder, recursive, exclusions, organize_by_date)

    def on_any_event(self, event):
        """Log all file system events for debugging."""
        logging.debug("Received event: type=%s, src_path=%s, is_directory=%s", event.event_type, event.src_path, event.is_directory)

    def on_deleted(self, event):
        """Track deleted files to detect potential renames."""
        if not event.is_directory:
            file_path = event.src_path
            file_name = os.path.basename(file_path)
            self.recent_deletions[file_path] = time.time()
            logging.debug("Tracked deletion of %s for rename detection", file_name)

    def on_created(self, event):
        """Handle file creation events and potential renames with retry mechanism."""
        if not self.is_running or self.is_paused or event.is_directory:
            self.log_callback(f"Skipped create event for {event.src_path}: {'stopped' if not self.is_running else 'paused' if self.is_paused else 'directory'}")
            return
        
        file_path = event.src_path
        if any(os.path.normpath(file_path).startswith(os.path.normpath(excl)) for excl in self.exclusions):
            self.log_callback(f"Skipped {os.path.basename(file_path)}: in excluded folder")
            return
        
        file_name = os.path.basename(file_path)
        logging.info("Detected create event for %s in %s", file_name, os.path.dirname(file_path))
        
        is_rename = False
        original_path = None
        for deleted_path, deletion_time in list(self.recent_deletions.items()):
            if time.time() - deletion_time < 5:
                if os.path.dirname(deleted_path) == os.path.dirname(file_path):
                    is_rename = True
                    original_path = deleted_path
                    del self.recent_deletions[deleted_path]
                    break
        
        for attempt in range(5):
            try:
                if os.path.isfile(file_path):
                    success, message = organize_file(file_path, self.base_folder, self.categories, self.organize_by_date)
                    if is_rename:
                        message = f"Renamed {os.path.basename(original_path)} to {file_name}: {message}"
                    self.log_callback(message)
                    logging.debug("Successfully processed %s: %s", file_name, message)
                    break
                else:
                    logging.info("Attempt %d: Skipped %s: file not ready", attempt + 1, file_name)
                    self.log_callback(f"Attempt {attempt + 1}: Skipped {file_name}: file not ready")
                    time.sleep(3)
            except PermissionError as e:
                logging.error("Permission error for %s: %s", file_name, e)
                self.log_callback(f"Permission error for {file_name}: {str(e)}. Try running as administrator.")
                time.sleep(3)
            except FileNotFoundError as e:
                logging.error("File not found for %s: %s", file_name, e)
                self.log_callback(f"File not found for {file_name}: {str(e)}")
                break
            except Exception as e:
                logging.error("Error processing %s: %s", file_name, e)
                self.log_callback(f"Error processing {file_name}: {str(e)}")
                time.sleep(3)
        else:
            message = f"Failed to process created file {file_name} after 5 attempts"
            logging.warning(message)
            self.log_callback(message)

    def on_moved(self, event):
        """Handle file rename events with retry mechanism."""
        logging.debug("Processing moved event: src_path=%s, dest_path=%s, is_directory=%s", event.src_path, event.dest_path, event.is_directory)
        
        if not self.is_running or self.is_paused or event.is_directory:
            self.log_callback(f"Skipped rename event for {event.src_path}: {'stopped' if not self.is_running else 'paused' if self.is_paused else 'directory'}")
            return
        
        file_path = event.dest_path
        file_name = os.path.basename(file_path)
        
        if any(os.path.normpath(file_path).startswith(os.path.normpath(excl)) for excl in self.exclusions):
            self.log_callback(f"Skipped {file_name}: in excluded folder")
            return
        
        logging.info("Detected rename event for %s (from %s) in %s", file_name, os.path.basename(event.src_path), os.path.dirname(file_path))
        
        for attempt in range(5):
            try:
                logging.debug("Attempt %d: Checking file %s", attempt + 1, file_path)
                if not os.path.exists(file_path):
                    logging.error("Rename target %s does not exist", file_name)
                    self.log_callback(f"Skipped {file_name}: rename target does not exist")
                    break
                if not os.path.isfile(file_path):
                    logging.info("Attempt %d: Skipped %s: not a file", attempt + 1, file_name)
                    self.log_callback(f"Attempt {attempt + 1}: Skipped {file_name}: not a file")
                    time.sleep(3)
                    continue
                if not os.access(file_path, os.R_OK | os.W_OK):
                    logging.error("Permission error for %s: no read/write access", file_name)
                    self.log_callback(f"Permission error for {file_name}: no read/write access. Try running as administrator.")
                    time.sleep(3)
                    continue
                success, message = organize_file(file_path, self.base_folder, self.categories, self.organize_by_date)
                message = f"Renamed {os.path.basename(event.src_path)} to {file_name}: {message}"
                self.log_callback(message)
                logging.debug("Successfully processed %s: %s", file_name, message)
                break
            except PermissionError as e:
                logging.error("Permission error for %s: %s", file_name, e)
                self.log_callback(f"Permission error for {file_name}: {str(e)}. Try running as administrator.")
                time.sleep(3)
            except FileNotFoundError as e:
                logging.error("File not found for %s: %s", file_name, e)
                self.log_callback(f"File not found for {file_name}: {str(e)}")
                break
            except Exception as e:
                logging.error("Error processing renamed %s: %s", file_name, e)
                self.log_callback(f"Error processing renamed {file_name}: {str(e)}")
                time.sleep(3)
        else:
            message = f"Failed to process renamed file {file_name} after 5 attempts"
            logging.warning(message)
            self.log_callback(message)

    def pause(self):
        """Pause the handler."""
        self.is_paused = True
        logging.info("Paused file organizer for %s", self.base_folder)
        self.log_callback(f"Paused watching {self.base_folder}")

    def resume(self):
        """Resume the handler."""
        self.is_paused = False
        logging.info("Resumed file organizer for %s", self.base_folder)
        self.log_callback(f"Resumed watching {self.base_folder}")

    def stop(self):
        """Stop the handler."""
        self.is_running = False
        logging.info("File organizer handler stopped for %s", self.base_folder)

class FileOrganizerApp:
    """GUI application for the file organizer."""
    def __init__(self, root):
        self.root = root
        self.root.title("File Organizer")
        self.root.geometry("600x500")
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Initialize log buffer and state
        self._log_buffer = []  # Buffer for early log messages
        self.log_text = None
        self.config = None
        self.categories = None
        self.monitored_folders = None
        self.startup_enabled = False
        self.appearance_mode = "system"
        self.organize_by_date = False
        self.folder_settings = {}
        self.observers = []
        self.handlers = []
        self.is_watching = False
        self.tray = None

        # Load configuration
        try:
            self.config = load_config()
            self.categories = self.config.get('categories', {})
            self.monitored_folders = self.config.get('monitored_folders', [])
            self.startup_enabled = self.config.get('startup_enabled', False)
            self.appearance_mode = self.config.get('appearance_mode', 'system')
            self.organize_by_date = self.config.get('organize_by_date', False)
            self.folder_settings = self.config.get('folder_settings', {})
            logging.info("Configuration loaded successfully")
        except Exception as e:
            logging.error("Failed to load configuration: %s", e)
            self._log_buffer.append(f"Error loading configuration: {str(e)}")

        # GUI Elements
        self.status_label = ctk.CTkLabel(root, text="Status: Stopped", text_color="red", font=("Arial", 12))
        self.status_label.pack(pady=5)

        self.label = ctk.CTkLabel(root, text="File Organizer", font=("Arial", 16))
        self.label.pack(pady=5)

        self.theme_frame = ctk.CTkFrame(root)
        self.theme_frame.pack(pady=5)
        self.theme_label = ctk.CTkLabel(self.theme_frame, text="Theme:", font=("Arial", 12))
        self.theme_label.pack(side="left", padx=5)
        self.theme_option = ctk.CTkOptionMenu(
            self.theme_frame,
            values=["System", "Light", "Dark"],
            command=self.change_theme
        )
        self.theme_option.set(self.appearance_mode.capitalize())
        self.theme_option.pack(side="left", padx=5)

        self.folder_frame = ctk.CTkFrame(root)
        self.folder_frame.pack(pady=5, fill="x", padx=10)

        self.folder_listbox = tk.Listbox(self.folder_frame, height=5)
        self.folder_listbox.pack(side="left", fill="x", expand=True, padx=5)
        if self.monitored_folders:
            for folder in self.monitored_folders:
                self.folder_listbox.insert(tk.END, folder)

        self.folder_button_frame = ctk.CTkFrame(self.folder_frame)
        self.folder_button_frame.pack(side="right", padx=5)
        self.add_folder_button = ctk.CTkButton(self.folder_button_frame, text="Add Folder", command=self.add_folder)
        self.add_folder_button.pack(pady=2)
        self.remove_folder_button = ctk.CTkButton(self.folder_button_frame, text="Remove Folder", command=self.remove_folder)
        self.remove_folder_button.pack(pady=2)
        self.edit_folder_button = ctk.CTkButton(self.folder_button_frame, text="Edit Folder Settings", command=self.edit_folder_settings)
        self.edit_folder_button.pack(pady=2)

        self.category_button = ctk.CTkButton(root, text="Edit Categories", command=self.edit_categories)
        self.category_button.pack(pady=5)

        self.control_frame = ctk.CTkFrame(root)
        self.control_frame.pack(pady=5)
        self.start_button = ctk.CTkButton(self.control_frame, text="Start Watching", command=self.start_watching)
        self.start_button.pack(side="left", padx=5)
        self.pause_button = ctk.CTkButton(self.control_frame, text="Pause Watching", command=self.pause_watching, state="disabled")
        self.pause_button.pack(side="left", padx=5)
        self.stop_button = ctk.CTkButton(self.control_frame, text="Stop Watching", command=self.stop_watching, state="disabled")
        self.stop_button.pack(side="left", padx=5)

        self.options_frame = ctk.CTkFrame(root)
        self.options_frame.pack(pady=5)
        self.date_organize_checkbox = ctk.CTkCheckBox(
            self.options_frame,
            text="Organize by Date",
            command=self.toggle_date_organization,
            variable=tk.BooleanVar(value=self.organize_by_date)
        )
        self.date_organize_checkbox.pack(side="left", padx=5)
        self.update_button = ctk.CTkButton(self.options_frame, text="Check for Updates", command=self.check_for_updates)
        self.update_button.pack(side="left", padx=5)

        # Initialize log_text after all other widgets
        try:
            self.log_text = ctk.CTkTextbox(root, height=150, state="disabled")
            self.log_text.pack(pady=10, fill="both", expand=True, padx=10)
        except Exception as e:
            logging.error("Error creating log textbox: %s", e)
            self._log_buffer.append(f"Error creating log textbox: {str(e)}")

        self.log_button_frame = ctk.CTkFrame(root)
        self.log_button_frame.pack(pady=5)
        self.clear_log_button = ctk.CTkButton(self.log_button_frame, text="Clear Log", command=self.clear_log)
        self.clear_log_button.pack(side="left", padx=5)
        self.export_log_button = ctk.CTkButton(self.log_button_frame, text="Export Log to CSV", command=self.export_log_to_csv)
        self.export_log_button.pack(side="left", padx=5)

        self.startup_checkbox = ctk.CTkCheckBox(root, text="Start on Boot", command=self.toggle_startup, variable=tk.BooleanVar(value=self.startup_enabled))
        self.startup_checkbox.pack(pady=5)
        logging.info("Start on Boot checkbox initialized")

        # Set application window icon with fallback
        try:
            icon_path = get_resource_path('my_icon.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                logging.info("Successfully set window icon: %s", icon_path)
                self._log_buffer.append("Successfully set window icon")
            else:
                logging.warning("Icon file my_icon.ico not found at %s", icon_path)
                self._log_buffer.append("Warning: my_icon.ico not found, using default icon")
        except Exception as e:
            logging.error("Error setting window icon: %s", e)
            self._log_buffer.append(f"Error setting window icon: {str(e)}")

        # Apply appearance mode
        try:
            ctk.set_appearance_mode(self.appearance_mode)
            logging.info("Applied appearance mode: %s", self.appearance_mode)
        except Exception as e:
            logging.error("Error applying appearance mode: %s", e)
            self._log_buffer.append(f"Error applying appearance mode: {str(e)}")

        # Create system tray icon
        try:
            self.create_tray_icon()
            logging.info("System tray icon created")
        except Exception as e:
            logging.error("Error creating system tray icon: %s", e)
            self._log_buffer.append(f"Error creating system tray icon: {str(e)}")

        # Flush buffered log messages to GUI
        if self.log_text:
            for message in self._log_buffer:
                self._write_to_gui_log(message)
            self._log_buffer = []  # Clear buffer after flushing
            self._write_to_gui_log("Loaded configuration. Ready to start watching.")

        # Start watching if folders are configured
        if self.monitored_folders:
            try:
                self.start_watching()
                self._write_to_gui_log(f"Automatically started watching {len(self.monitored_folders)} folder(s): {', '.join(self.monitored_folders)}")
            except Exception as e:
                logging.error("Error starting watchers: %s", e)
                self._write_to_gui_log(f"Error starting watchers: {str(e)}")
        else:
            self._write_to_gui_log("No folders configured. Please add folders to monitor.")

    def log_to_gui(self, message):
        """Add a message to the log buffer or GUI log area if initialized."""
        if self.log_text is None:
            self._log_buffer.append(message)
            logging.info(message)  # Log to file as fallback
        else:
            self._write_to_gui_log(message)

    def _write_to_gui_log(self, message):
        """Write a message to the GUI log area."""
        if self.log_text:
            try:
                self.log_text.configure(state="normal")
                self.log_text.insert(tk.END, f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
                self.log_text.see(tk.END)
                self.log_text.configure(state="disabled")
            except Exception as e:
                logging.error("Error writing to GUI log: %s", e)
                logging.info(message)  # Fallback to file logging
        else:
            logging.info(message)  # Fallback to file logging

    def change_theme(self, choice):
        """Change the application theme and save to config."""
        try:
            self.appearance_mode = choice.lower()
            ctk.set_appearance_mode(self.appearance_mode)
            self.config['appearance_mode'] = self.appearance_mode
            save_config(self.config)
            self.log_to_gui(f"Changed theme to {choice}")
        except Exception as e:
            logging.error("Error changing theme: %s", e)
            self.log_to_gui(f"Error changing theme: {str(e)}")

    def toggle_date_organization(self):
        """Toggle organize by date and save to config."""
        try:
            self.organize_by_date = self.date_organize_checkbox.get()
            self.config['organize_by_date'] = self.organize_by_date
            save_config(self.config)
            self.log_to_gui(f"{'Enabled' if self.organize_by_date else 'Disabled'} organize by date")
            if self.is_watching:
                self.restart_watching()
        except Exception as e:
            logging.error("Error toggling date organization: %s", e)
            self.log_to_gui(f"Error toggling date organization: {str(e)}")

    def export_log_to_csv(self):
        """Export organizer.log to a CSV file."""
        try:
            output_file = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                title="Save Log as CSV"
            )
            if not output_file:
                return
            
            log_entries = []
            with open('organizer.log', 'r') as f:
                for line in f:
                    try:
                        timestamp, level, message = line.strip().split(' - ', 2)
                        log_entries.append({'Timestamp': timestamp, 'Level': level, 'Message': message})
                    except ValueError:
                        continue
            
            with open(output_file, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['Timestamp', 'Level', 'Message'])
                writer.writeheader()
                writer.writerows(log_entries)
            
            self.log_to_gui(f"Exported log to {output_file}")
            logging.info("Exported log to %s", output_file)
        except Exception as e:
            logging.error("Error exporting log to CSV: %s", e)
            self.log_to_gui(f"Error exporting log to CSV: {str(e)}")

    def check_for_updates(self):
        """Check for updates using GitHub API."""
        try:
            repo = "username/repo"  # Replace with your actual repo
            response = requests.get(f"https://api.github.com/repos/{repo}/releases/latest", timeout=5)
            response.raise_for_status()
            latest_version = response.json().get('tag_name', '0.0.0')
            if latest_version > APP_VERSION:
                self.log_to_gui(f"Update available: {latest_version}. Current version: {APP_VERSION}. Visit https://github.com/{repo} to download.")
                messagebox.showinfo("Update Available", f"A new version ({latest_version}) is available. Visit https://github.com/{repo} to download.")
            else:
                self.log_to_gui(f"No updates available. Current version: {APP_VERSION}")
                messagebox.showinfo("No Updates", "You are running the latest version.")
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                logging.error("Error checking for updates: Repository or release not found for %s", repo)
                self.log_to_gui("Error checking for updates: Repository or release not found. Please ensure the repository exists and has releases.")
                messagebox.showerror("Update Error", "Repository or release not found. Please ensure the repository exists and has releases.")
            else:
                logging.error("Error checking for updates: %s", e)
                self.log_to_gui(f"Error checking for updates: {str(e)}")
                messagebox.showerror("Update Error", f"Failed to check for updates: {str(e)}")
        except Exception as e:
            logging.error("Error checking for updates: %s", e)
            self.log_to_gui(f"Error checking for updates: {str(e)}")
            messagebox.showerror("Update Error", f"Failed to check for updates: {str(e)}")

    def add_folder(self):
        """Add a new folder to monitor."""
        try:
            folder = filedialog.askdirectory(title="Select Folder to Monitor")
            if folder and folder not in self.monitored_folders:
                self.monitored_folders.append(folder)
                self.folder_listbox.insert(tk.END, folder)
                self.folder_settings[folder] = {"recursive": True, "exclusions": []}
                self.config['monitored_folders'] = self.monitored_folders
                self.config['folder_settings'] = self.folder_settings
                save_config(self.config)
                self.log_to_gui(f"Added folder: {folder}")
                if self.is_watching:
                    self.restart_watching()
        except Exception as e:
            logging.error("Error adding folder: %s", e)
            self.log_to_gui(f"Error adding folder: {str(e)}")

    def remove_folder(self):
        """Remove a selected folder from monitoring."""
        try:
            selected = self.folder_listbox.curselection()
            if selected:
                folder = self.folder_listbox.get(selected[0])
                self.monitored_folders.remove(folder)
                self.folder_listbox.delete(selected[0])
                self.folder_settings.pop(folder, None)
                self.config['monitored_folders'] = self.monitored_folders
                self.config['folder_settings'] = self.folder_settings
                save_config(self.config)
                self.log_to_gui(f"Removed folder: {folder}")
                if self.is_watching:
                    self.restart_watching()
        except Exception as e:
            logging.error("Error removing folder: %s", e)
            self.log_to_gui(f"Error removing folder: {str(e)}")

    def edit_folder_settings(self):
        """Open a dialog to edit folder settings (recursive, exclusions)."""
        try:
            selected = self.folder_listbox.curselection()
            if not selected:
                self.log_to_gui("Error: No folder selected for editing settings.")
                messagebox.showerror("Error", "Please select a folder to edit settings.")
                return
            
            folder = self.folder_listbox.get(selected[0])
            dialog = ctk.CTkToplevel(self.root)
            dialog.title(f"Edit Settings for {folder}")
            dialog.geometry("400x300")
            dialog.attributes('-topmost', True)
            dialog.grab_set()

            recursive_var = tk.BooleanVar(value=self.folder_settings.get(folder, {}).get("recursive", False))
            ctk.CTkCheckBox(dialog, text="Monitor Subfolders (Recursive)", variable=recursive_var).pack(pady=5)

            ctk.CTkLabel(dialog, text="Excluded Subfolders:").pack(pady=5)
            exclusion_listbox = tk.Listbox(dialog, height=5)
            exclusion_listbox.pack(pady=5, fill="both", expand=True, padx=10)
            exclusions = self.folder_settings.get(folder, {}).get("exclusions", [])
            for excl in exclusions:
                exclusion_listbox.insert(tk.END, excl)

            def add_exclusion():
                excl_folder = filedialog.askdirectory(title="Select Subfolder to Exclude")
                if excl_folder and os.path.normpath(excl_folder).startswith(os.path.normpath(folder)) and excl_folder not in exclusions:
                    exclusions.append(excl_folder)
                    exclusion_listbox.insert(tk.END, excl_folder)
                    self.log_to_gui(f"Added exclusion {excl_folder} for {folder}")

            def remove_exclusion():
                selected_excl = exclusion_listbox.curselection()
                if selected_excl:
                    excl_folder = exclusion_listbox.get(selected_excl[0])
                    exclusions.remove(excl_folder)
                    exclusion_listbox.delete(selected_excl[0])
                    self.log_to_gui(f"Removed exclusion {excl_folder} for {folder}")

            button_frame = ctk.CTkFrame(dialog)
            button_frame.pack(pady=5)
            ctk.CTkButton(button_frame, text="Add Exclusion", command=add_exclusion).pack(side="left", padx=5)
            ctk.CTkButton(button_frame, text="Remove Exclusion", command=remove_exclusion).pack(side="left", padx=5)

            def save_settings():
                self.folder_settings[folder] = {
                    "recursive": recursive_var.get(),
                    "exclusions": exclusions
                }
                self.config['folder_settings'] = self.folder_settings
                save_config(self.config)
                self.log_to_gui(f"Updated settings for {folder}: recursive={recursive_var.get()}, exclusions={exclusions}")
                if self.is_watching:
                    self.restart_watching()
                dialog.attributes('-topmost', False)
                dialog.grab_release()
                dialog.destroy()

            ctk.CTkButton(dialog, text="Save", command=save_settings).pack(pady=5)

            def on_close():
                dialog.attributes('-topmost', False)
                dialog.grab_release()
                dialog.destroy()

            dialog.protocol("WM_DELETE_WINDOW", on_close)
        except Exception as e:
            logging.error("Error editing folder settings: %s", e)
            self.log_to_gui(f"Error editing folder settings: {str(e)}")

    def edit_categories(self):
        """Open a dialog to edit category mappings."""
        try:
            dialog = ctk.CTkToplevel(self.root)
            dialog.title("Edit Categories")
            dialog.geometry("400x300")
            dialog.attributes('-topmost', True)
            dialog.grab_set()
            logging.info("Edit Categories dialog opened and set to topmost")

            listbox = tk.Listbox(dialog, height=10)
            listbox.pack(pady=10, fill="both", expand=True, padx=10)
            for ext, category in self.categories.items():
                listbox.insert(tk.END, f"{ext} -> {category}")

            def add_category():
                ext = ctk.CTkInputDialog(text="Enter file extension (e.g., .xlsx):", title="Add Category").get_input()
                if ext and ext.startswith('.'):
                    category = ctk.CTkInputDialog(text=f"Enter category for {ext}:", title="Add Category").get_input()
                    if category:
                        self.categories[ext.lower()] = category
                        listbox.insert(tk.END, f"{ext.lower()} -> {category}")
                        self.config['categories'] = self.categories
                        save_config(self.config)
                        self.log_to_gui(f"Added category: {ext.lower()} -> {category}")
                        if self.is_watching:
                            self.restart_watching()

            def remove_category():
                selected = listbox.curselection()
                if selected:
                    ext_category = listbox.get(selected[0])
                    ext = ext_category.split(' -> ')[0]
                    del self.categories[ext]
                    listbox.delete(selected[0])
                    self.config['categories'] = self.categories
                    save_config(self.config)
                    self.log_to_gui(f"Removed category: {ext}")
                    if self.is_watching:
                        self.restart_watching()

            button_frame = ctk.CTkFrame(dialog)
            button_frame.pack(pady=5)
            ctk.CTkButton(button_frame, text="Add Category", command=add_category).pack(side="left", padx=5)
            ctk.CTkButton(button_frame, text="Remove Category", command=remove_category).pack(side="left", padx=5)

            def on_close():
                dialog.attributes('-topmost', False)
                dialog.grab_release()
                dialog.destroy()
                logging.info("Edit Categories dialog closed")

            dialog.protocol("WM_DELETE_WINDOW", on_close)
        except Exception as e:
            logging.error("Error editing categories: %s", e)
            self.log_to_gui(f"Error editing categories: {str(e)}")

    def start_watching(self):
        """Start the file watchers in a separate thread."""
        try:
            if not self.monitored_folders:
                self.log_to_gui("Error: No folders selected to monitor.")
                messagebox.showerror("Error", "No folders selected to monitor.")
                return
            
            self.is_watching = True
            self.start_button.configure(state="disabled")
            self.pause_button.configure(state="normal")
            self.stop_button.configure(state="normal")
            self.status_label.configure(text="Status: Watching", text_color="green")
            
            def run_watchers():
                try:
                    self.observers, self.handlers = start_watcher(
                        self.monitored_folders,
                        self.categories,
                        self.log_to_gui,
                        self.folder_settings,
                        self.organize_by_date
                    )
                    watched_folders = [handler.base_folder for handler in self.handlers]
                    configured_folders = set(self.monitored_folders)
                    missing_folders = configured_folders - set(watched_folders)
                    if missing_folders:
                        self.log_to_gui(f"Warning: Failed to start watchers for {', '.join(missing_folders)}")
                except Exception as e:
                    logging.error("Error in run_watchers: %s", e)
                    self.log_to_gui(f"Error starting watchers: {str(e)}")
            
            threading.Thread(target=run_watchers, daemon=True).start()
            self.log_to_gui("Started watching folders.")
        except Exception as e:
            logging.error("Error starting watching: %s", e)
            self.log_to_gui(f"Error starting watching: {str(e)}")

    def pause_watching(self):
        """Pause all file watchers."""
        try:
            for handler in self.handlers:
                handler.pause()
            self.pause_button.configure(text="Resume Watching", command=self.resume_watching)
            self.status_label.configure(text="Status: Paused", text_color="yellow")
            self.log_to_gui("Paused watching all folders.")
        except Exception as e:
            logging.error("Error pausing watching: %s", e)
            self.log_to_gui(f"Error pausing watching: {str(e)}")

    def resume_watching(self):
        """Resume all file watchers."""
        try:
            for handler in self.handlers:
                handler.resume()
            self.pause_button.configure(text="Pause Watching", command=self.pause_watching)
            self.status_label.configure(text="Status: Watching", text_color="green")
            self.log_to_gui("Resumed watching all folders.")
        except Exception as e:
            logging.error("Error resuming watching: %s", e)
            self.log_to_gui(f"Error resuming watching: {str(e)}")

    def stop_watching(self):
        """Stop all file watchers."""
        try:
            stop_watcher(self.observers, self.handlers)
            self.observers = []
            self.handlers = []
            self.is_watching = False
            self.start_button.configure(state="normal")
            self.pause_button.configure(state="disabled", text="Pause Watching", command=self.pause_watching)
            self.stop_button.configure(state="disabled")
            self.status_label.configure(text="Status: Stopped", text_color="red")
            self.log_to_gui("Stopped watching folders.")
        except Exception as e:
            logging.error("Error stopping watching: %s", e)
            self.log_to_gui(f"Error stopping watching: {str(e)}")

    def restart_watching(self):
        """Restart watchers after configuration changes."""
        try:
            if self.is_watching:
                self.stop_watching()
                self.start_watching()
        except Exception as e:
            logging.error("Error restarting watching: %s", e)
            self.log_to_gui(f"Error restarting watching: {str(e)}")

    def clear_log(self):
        """Clear the log text area and reset scroll position."""
        try:
            if self.log_text:
                self.log_text.configure(state="normal")
                self.log_text.delete("1.0", tk.END)
                self.log_text.insert(tk.END, f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Log cleared.\n")
                self.log_text.see(tk.END)
                self.log_text.configure(state="disabled")
                logging.info("Log cleared in GUI")
        except Exception as e:
            logging.error("Error clearing log: %s", e)
            self.log_to_gui(f"Error clearing log: {str(e)}")

    def minimize_to_tray(self):
        """Minimize the window to the system tray."""
        try:
            self.root.withdraw()
            if self.tray:
                self.tray.visible = True
            self.log_to_gui("Minimized to system tray.")
        except Exception as e:
            logging.error("Error minimizing to tray: %s", e)
            self.log_to_gui(f"Error minimizing to tray: {str(e)}")

    def restore_from_tray(self):
        """Restore the window from the system tray."""
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
            self.log_to_gui("Restored from system tray.")
        except Exception as e:
            logging.error("Error restoring from tray: %s", e)
            self.log_to_gui(f"Error restoring from tray: {str(e)}")

    def exit_app(self):
        """Exit the application and clean up."""
        try:
            if self.is_watching:
                self.stop_watching()
            if self.tray:
                self.tray.stop()
            self.root.quit()
            self.log_to_gui("Application exited.")
        except Exception as e:
            logging.error("Error exiting application: %s", e)
            self.log_to_gui(f"Error exiting application: {str(e)}")

    def create_tray_icon(self):
        """Create a system tray icon with a menu."""
        try:
            if self.tray:
                self.tray.stop()

            icon_path = get_resource_path("my_icon.png")
            if os.path.exists(icon_path):
                image = Image.open(icon_path)
                logging.info("Successfully loaded my_icon.png")
            else:
                image = Image.new("RGB", (16, 16), "white")
                logging.info("my_icon.png not found, using default white image")

            def on_start(icon, item):
                self.root.after(0, self.start_watching)

            def on_stop(icon, item):
                self.root.after(0, self.stop_watching)

            def on_show(icon, item):
                self.root.after(0, self.restore_from_tray)

            def on_exit(icon, item):
                self.root.after(0, self.exit_app)

            menu = pystray.Menu(
                pystray.MenuItem("Start Organizing", on_start),
                pystray.MenuItem("Stop Organizing", on_stop),
                pystray.MenuItem("Show", on_show),
                pystray.MenuItem("Exit", on_exit)
            )
            self.tray = pystray.Icon("File Organizer", image, "File Organizer", menu)
            threading.Thread(target=self.tray.run, daemon=True).start()
        except Exception as e:
            logging.error("Error creating tray icon: %s", e)
            self.log_to_gui(f"Error creating tray icon: {str(e)}")

    def toggle_startup(self):
        """Toggle startup on boot and update config."""
        try:
            self.startup_enabled = self.startup_checkbox.get()
            self.config['startup_enabled'] = self.startup_enabled
            save_config(self.config)
            if self.startup_enabled:
                self.set_startup(True)
                self.log_to_gui("Enabled startup on boot.")
            else:
                self.set_startup(False)
                self.log_to_gui("Disabled startup on boot.")
        except Exception as e:
            logging.error("Error toggling startup: %s", e)
            self.log_to_gui(f"Error toggling startup: {str(e)}")

    def set_startup(self, enable):
        """Add or remove a startup shortcut."""
        try:
            pythoncom.CoInitialize()
            shell = win32com.client.Dispatch("WScript.Shell")
            startup_path = os.path.join(winshell.startup(), "File Organizer.lnk")
            script_path = get_resource_path(sys.executable if getattr(sys, 'frozen', False) else __file__)

            if enable and not os.path.exists(startup_path):
                shortcut = shell.CreateShortCut(startup_path)
                shortcut.TargetPath = script_path
                shortcut.WorkingDirectory = os.path.dirname(script_path)
                shortcut.IconLocation = get_resource_path("my_icon.ico")
                shortcut.WindowStyle = 7
                shortcut.save()
                logging.info("Added startup shortcut at %s", startup_path)
            elif not enable and os.path.exists(startup_path):
                os.remove(startup_path)
                logging.info("Removed startup shortcut")
        except Exception as e:
            logging.error("Error setting startup: %s", e)
            self.log_to_gui(f"Error setting startup: {str(e)}")
        finally:
            pythoncom.CoUninitialize()

def start_watcher(folders, categories, log_callback, folder_settings, organize_by_date):
    """Start file system watchers for multiple folders with settings."""
    observers = []
    handlers = []
    
    for folder in folders:
        if not os.path.isabs(folder) or '\x0c' in folder:
            logging.error("Invalid folder path: %s", folder)
            log_callback(f"Invalid folder path: {folder}")
            continue
        
        if not os.path.exists(folder):
            try:
                os.makedirs(folder, exist_ok=True)
                logging.info("Created folder %s", folder)
                log_callback(f"Created folder {folder}")
            except Exception as e:
                logging.error("Error creating folder %s: %s", folder, e)
                log_callback(f"Error creating folder {folder}: {str(e)}")
                continue
        elif not os.access(folder, os.R_OK | os.W_OK):
            logging.error("No read/write access to folder %s", folder)
            log_callback(f"No read/write access to folder {folder}. Try running as administrator.")
            continue
        
        try:
            settings = folder_settings.get(folder, {"recursive": True, "exclusions": []})
            event_handler = FileOrganizerHandler(
                folder,
                categories,
                log_callback,
                recursive=settings["recursive"],
                exclusions=settings["exclusions"],
                organize_by_date=organize_by_date
            )
            observer = Observer()
            observer.schedule(event_handler, folder, recursive=settings["recursive"])
            observer.start()
            logging.info("Started file watcher for %s (recursive=%s)", folder, settings["recursive"])
            log_callback(f"Started file watcher for {folder} (recursive={settings['recursive']})")
            observers.append(observer)
            handlers.append(event_handler)
        except Exception as e:
            logging.error("Error starting watcher for %s: %s", folder, e)
            log_callback(f"Error starting watcher for {folder}: {str(e)}")
    
    if not observers:
        logging.warning("No watchers started for any folders")
        log_callback("Warning: No watchers started. Check folder paths and permissions.")
    
    return observers, handlers

def stop_watcher(observers, handlers):
    """Stop all file system watchers."""
    try:
        for observer, handler in zip(observers, handlers):
            observer.stop()
            observer.join()
            handler.stop()
        logging.info("Stopped all file watchers")
    except Exception as e:
        logging.error("Error stopping watchers: %s", e)

def main():
    """Run the GUI application."""
    try:
        ctk.set_appearance_mode("System")
        root = ctk.CTk()
        app = FileOrganizerApp(root)
        if app.startup_enabled and "--minimized" in sys.argv:
            app.minimize_to_tray()
        root.mainloop()
    except Exception as e:
        logging.error("Error in main: %s", e)
        if 'app' in locals():
            app.log_to_gui(f"Error in application: {str(e)}")
        else:
            logging.critical("Failed to initialize application: %s", e)

if __name__ == "__main__":
    main()