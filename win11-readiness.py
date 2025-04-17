import customtkinter as ctk
from PIL import Image
import platform
import psutil
import socket
import requests
import json
import threading
import queue
import time
import tempfile
import csv
import os
import sys

# --- Attempt Imports ---
# WMI
try:
    import wmi
    _wmi_available = True
except ImportError:
    _wmi_available = False
# WinReg
try:
    import winreg
    _winreg_available = True
except ImportError:
     _winreg_available = False
# Windows Update Agent COM API (for pending updates)
try:
    import win32com.client
    _wuapi_available = True
except ImportError:
    _wuapi_available = False

# --- Configuration ---
APP_NAME = "Win11 Readiness Check"
LOGO_FILENAME = "logo.png"
# --- !!! REPLACE WITH YOUR ACTUAL API ENDPOINT !!! ---
API_ENDPOINT_URL = "YOUR_API_ENDPOINT_HERE"
# --- !!! REPLACE WITH YOUR SUPPORT PHONE NUMBER !!! ---
SUPPORT_PHONE_NUMBER = "1-800-555-HELP"
# --- End Configuration ---

gui_queue = queue.Queue()

# --- Data Class ---
class SystemInfo:
    """ Holds all collected system information with default values. """
    def __init__(self):
        # User Input
        self.assessment_id = "Not Provided" # Added field for user ID
        # Basic Info
        self.hostname = "Undetermined"
        self.os_platform = "Undetermined"
        self.os_version = "Undetermined"
        self.os_release = "Undetermined"
        self.architecture = "Undetermined"
        self.processor = "Undetermined"
        # Timestamps
        self.timestamp_utc = "Undetermined"
        self.timestamp_local = "Undetermined"
        self.timezone_name = "Undetermined"
        self.timezone_offset_utc = "Undetermined"
        # Hardware Info
        self.ram_total_gb = 0.0
        self.disk_total_gb = 0.0
        self.disk_free_gb = 0.0
        # WMI Dependent Info
        self.tpm_present = "Check Not Run"
        self.tpm_version = "Check Not Run"
        self.tpm_enabled = "Check Not Run"
        self.secure_boot_enabled = "Check Not Run"
        self.graphics_card = "Check Not Run"
        self.wddm_version = "Check Not Run"
        # Windows Update Info
        self.pending_updates_count = -1 # Added field, -1 indicates check failed/not run
        # Status/Error fields
        self.collection_error = None

    def to_dict(self):
        """ Convert the object's attributes to a dictionary for serialization. """
        return self.__dict__

# --- Helper Functions ---

def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def update_status(message):
    gui_queue.put(f"STATUS: {message}")

def show_final_message(message_type, message):
    gui_queue.put(f"{message_type}: {message}")

# --- Data Collection Functions ---

def populate_wmi_info(info: SystemInfo):
    """ Attempts to get WMI info and updates the SystemInfo object. """
    if not _wmi_available:
        update_status("WMI module not found. Skipping WMI checks.")
        print("Python WMI module not installed (pip install wmi)")
        info.tpm_present = "WMI Module Missing"
        info.tpm_version = "WMI Module Missing"
        info.tpm_enabled = "WMI Module Missing"
        info.secure_boot_enabled = "WMI Module Missing"
        info.graphics_card = "WMI Module Missing"
        info.wddm_version = "WMI Module Missing"
        return

    permission_error_flag = "(Permissions?)"
    try:
        c = wmi.WMI()
    except Exception as e:
        update_status(f"WMI initialization failed: {e}. Skipping WMI checks.")
        print(f"WMI failed to initialize: {e}")
        info.tpm_present = f"WMI Init Failed {permission_error_flag}"
        # ... (set other WMI fields to Init Failed) ...
        info.secure_boot_enabled = f"WMI Init Failed {permission_error_flag}"
        info.graphics_card = f"WMI Init Failed {permission_error_flag}"
        info.wddm_version = f"WMI Init Failed {permission_error_flag}"
        return

    # --- TPM Check ---
    update_status("Querying WMI for TPM...")
    info.tpm_present = "Undetermined"
    info.tpm_version = "Undetermined"
    info.tpm_enabled = "Undetermined"
    try:
        tpm_info_list = c.Win32_Tpm()
        if tpm_info_list:
            tpm_info = tpm_info_list[0]
            info.tpm_present = True
            try: info.tpm_version = tpm_info.SpecVersion or "Unknown"
            except AttributeError: info.tpm_version = "Unknown"
            try:
                enabled = tpm_info.IsEnabled() if hasattr(tpm_info, 'IsEnabled') else None
                if enabled is not None: info.tpm_enabled = bool(enabled)
                else: info.tpm_enabled = f"Check Failed {permission_error_flag}"
            except Exception as e_detail:
                print(f"WMI TPM status check failed: {e_detail}")
                info.tpm_enabled = f"Check Failed {permission_error_flag}"
        else:
            info.tpm_present = False
            info.tpm_version = "N/A"
            info.tpm_enabled = "N/A"
    except Exception as e:
        print(f"WMI TPM query failed: {e}")
        info.tpm_present = f"Query Failed {permission_error_flag}"
        info.tpm_version = f"Query Failed {permission_error_flag}"
        info.tpm_enabled = f"Query Failed {permission_error_flag}"

    # --- Secure Boot Check ---
    update_status("Querying WMI for Secure Boot...")
    info.secure_boot_enabled = "Undetermined"
    try:
        sb_info_list = c.Win32_SecureBoot()
        if sb_info_list:
             if hasattr(sb_info_list[0], 'SecureBootEnabled'): info.secure_boot_enabled = bool(sb_info_list[0].SecureBootEnabled)
             elif hasattr(sb_info_list[0], 'IsEnabled'): info.secure_boot_enabled = bool(sb_info_list[0].IsEnabled)
             else: info.secure_boot_enabled = "Property Not Found"
        else: info.secure_boot_enabled = "Query Failed (Class Not Found?)"
    except Exception as e:
        print(f"WMI Secure Boot query failed: {e}")
        if _winreg_available:
            try:
                key_path = r"SYSTEM\CurrentControlSet\Control\SecureBoot\State"
                reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ)
                value, _ = winreg.QueryValueEx(reg_key, "UEFISecureBootEnabled")
                winreg.CloseKey(reg_key)
                info.secure_boot_enabled = bool(value)
            except FileNotFoundError: info.secure_boot_enabled = "Check Failed (Key Missing)"
            except PermissionError: info.secure_boot_enabled = f"Check Failed {permission_error_flag}"
            except Exception as reg_e:
                print(f"Registry Secure Boot check failed: {reg_e}")
                info.secure_boot_enabled = f"Check Failed {permission_error_flag}"
        else: info.secure_boot_enabled = f"Check Failed (WMI Error & WinReg Missing)"

    # --- Graphics Check ---
    update_status("Querying WMI for Graphics...")
    info.graphics_card = "Undetermined"
    info.wddm_version = "Undetermined"
    try:
        gpu_info = c.Win32_VideoController()[0]
        info.graphics_card = gpu_info.Name
        info.wddm_version = f"Driver: {gpu_info.DriverVersion}"
    except Exception as e:
        print(f"WMI Graphics query failed: {e}")
        info.graphics_card = "Query Failed"
        info.wddm_version = "Query Failed"


def check_pending_updates(info: SystemInfo):
    """ Checks for pending Windows Updates using WUA API. """
    update_status("Checking for pending Windows Updates...")
    info.pending_updates_count = -1 # Default to error/check failed

    if not _wuapi_available:
        update_status("Windows Update check skipped: pywin32 module missing.")
        print("Windows Update check skipped: pywin32 module missing (pip install pywin32).")
        info.pending_updates_count = -2 # Specific code for module missing
        return

    try:
        # Create COM objects
        update_session = win32com.client.Dispatch("Microsoft.Update.Session")
        update_searcher = update_session.CreateUpdateSearcher()

        # Search criteria: Updates that are not installed and not hidden
        # May require admin rights for a comprehensive search
        search_criteria = "IsInstalled=0 and IsHidden=0 and Type='Software'"
        update_status("Searching for available updates (this may take a moment)...")
        search_result = update_searcher.Search(search_criteria)
        update_status(f"Found {search_result.Updates.Count} available updates.")

        # Count updates ready for installation (already downloaded)
        # Note: This is a simplified check. A full check might look at specific
        # properties like IsDownloaded, IsMandatory, RebootRequired etc.
        # For simplicity, we count all applicable updates found by the search.
        # A count > 0 implies action might be needed.
        info.pending_updates_count = search_result.Updates.Count

    except pythoncom.com_error as com_err:
         # Handle COM errors (e.g., service not running, access denied)
         print(f"Windows Update check failed (COM Error): {com_err}")
         update_status(f"Windows Update check failed (COM Error: {com_err.hresult})")
         info.pending_updates_count = -3 # Specific code for COM error
    except Exception as e:
        print(f"Windows Update check failed: {e}")
        update_status(f"Windows Update check failed: {e}")
        info.pending_updates_count = -1 # General error

def collect_system_data(assessment_id: str) -> SystemInfo:
    """ Collects system information and returns a populated SystemInfo object. """
    info = SystemInfo()
    info.assessment_id = assessment_id # Store the user ID

    try:
        update_status("Collecting basic system info...")
        info.hostname = socket.gethostname()
        info.os_platform = platform.system()
        info.os_version = platform.version()
        info.os_release = platform.release()
        info.architecture = platform.machine()
        info.processor = platform.processor()

        # Timestamps
        current_time_utc = time.time()
        current_datetime_local = time.localtime(current_time_utc)
        timezone_name = time.tzname[current_datetime_local.tm_isdst]
        timezone_offset_seconds = -time.timezone if not current_datetime_local.tm_isdst else -time.altzone
        timezone_offset_hours = timezone_offset_seconds / 3600
        timezone_offset_str = f"{int(timezone_offset_hours):+03d}:00"
        info.timestamp_utc = time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime(current_time_utc))
        info.timestamp_local = time.strftime("%Y-%m-%d %H:%M:%S", current_datetime_local)
        info.timezone_name = timezone_name
        info.timezone_offset_utc = timezone_offset_str

        update_status("Collecting RAM info...")
        try:
            ram = psutil.virtual_memory()
            info.ram_total_gb = round(ram.total / (1024**3), 2)
        except Exception as e: info.ram_total_gb = -1.0

        update_status("Collecting Disk info (System Drive)...")
        try:
            system_drive = os.getenv("SystemDrive", "C:") + "\\"
            disk = psutil.disk_usage(system_drive)
            info.disk_total_gb = round(disk.total / (1024**3), 2)
            info.disk_free_gb = round(disk.free / (1024**3), 2)
        except Exception as e:
            info.disk_total_gb = -1.0
            info.disk_free_gb = -1.0

        # --- WMI Dependent Info ---
        populate_wmi_info(info)

        # --- Windows Update Check ---
        check_pending_updates(info) # Call the new check

        update_status("Data collection complete.")

    except Exception as e:
        print(f"CRITICAL ERROR during data collection: {e}")
        info.collection_error = str(e)
        update_status(f"Critical error during collection: {e}")

    return info

# --- Data Handling Functions ---

def send_data_to_api(info: SystemInfo):
    """ Sends collected data from SystemInfo object to the API endpoint """
    update_status(f"Sending data for {info.hostname} (Assessment ID: {info.assessment_id})...")
    headers = {'Content-Type': 'application/json'}
    data_dict = info.to_dict()
    try:
        response = requests.post(API_ENDPOINT_URL, headers=headers, json=data_dict, timeout=15)
        response.raise_for_status()
        update_status("Data sent successfully.")
        return True, "Success", info.hostname
    except requests.exceptions.Timeout:
        update_status("Error: Connection timed out.")
        return False, "Connection timed out.", info.hostname
    except requests.exceptions.RequestException as e:
        update_status(f"Error sending data: {e}")
        return False, f"Could not connect to server: {e}", info.hostname
    except Exception as e:
        update_status(f"An unexpected error occurred during sending: {e}")
        return False, f"An unexpected error occurred: {e}", info.hostname

def save_data_to_csv(info: SystemInfo):
    """ Saves collected data from SystemInfo object to a CSV file """
    data_dict = info.to_dict()
    if not data_dict: return
    try:
        temp_dir = tempfile.gettempdir()
        hostname = info.hostname if info.hostname != "Undetermined" else "unknown"
        assessment_id_part = info.assessment_id.replace(" ", "_") if info.assessment_id else "no_id"
        filename = os.path.join(temp_dir, f"readiness_check_{hostname}_{assessment_id_part}_{time.strftime('%Y%m%d_%H%M%S')}.csv")
        update_status(f"Saving data locally to {filename}...")
        fieldnames = sorted(list(data_dict.keys()))
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerow(data_dict)
        update_status(f"Data saved locally to {filename}")
        print(f"Data saved to {filename}")
    except Exception as e:
        update_status(f"Error saving data locally: {e}")
        print(f"Failed to save CSV: {e}")

# --- Background Task ---

def worker_thread_task(assessment_id: str): # Accept assessment_id
    """ The main task performed in the background thread """
    system_info = None
    try:
        system_info = collect_system_data(assessment_id) # Pass assessment_id

        if system_info.collection_error:
             show_final_message("ERROR", f"A critical error occurred during data collection: {system_info.collection_error}\nPlease call us at {SUPPORT_PHONE_NUMBER}.")
             save_data_to_csv(system_info)
             return

        success, status_msg, hostname = send_data_to_api(system_info)

        if success:
            show_final_message("SUCCESS", f"Assessment data for '{hostname}' (Assessment ID: {assessment_id}) has been submitted successfully.")
        else:
            show_final_message("ERROR", f"There was an issue submitting the data ({status_msg}).\nPlease call us at {SUPPORT_PHONE_NUMBER}.\n\nResults are being saved locally.")
            save_data_to_csv(system_info)

    except Exception as e:
        print(f"Critical error in worker thread: {e}")
        update_status(f"Critical error: {e}")
        show_final_message("ERROR", f"A critical error occurred: {e}\nPlease call us at {SUPPORT_PHONE_NUMBER}.")
        if system_info: save_data_to_csv(system_info)
    finally:
        gui_queue.put("TASK_COMPLETE")


# --- GUI Class (Modified for ID Prompt) ---

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(APP_NAME)
        self.geometry("450x200")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.logo_image = None
        self.status_label = None
        self.worker_thread = None
        self.assessment_id = "Not Provided" # Store user ID here

        self.show_splash_screen()

    def show_splash_screen(self):
        # (Splash screen code remains the same as before)
        self.withdraw()
        self.splash = ctk.CTkToplevel(self)
        self.splash.title("Loading...")
        self.splash.geometry("300x200")
        self.splash.resizable(False, False)
        self.splash.overrideredirect(True)

        screen_width = self.splash.winfo_screenwidth()
        screen_height = self.splash.winfo_screenheight()
        x_co = int((screen_width / 2) - (300 / 2))
        y_co = int((screen_height / 2) - (200 / 2))
        self.splash.geometry(f"300x200+{x_co}+{y_co}")

        try:
            logo_path = get_resource_path(LOGO_FILENAME)
            pil_image = Image.open(logo_path)
            self.logo_image = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(128, 128))
            logo_label = ctk.CTkLabel(self.splash, image=self.logo_image, text="")
            logo_label.pack(pady=20, padx=20, expand=True, fill="both")
        except FileNotFoundError:
            print(f"Error: Logo file '{LOGO_FILENAME}' not found at '{get_resource_path(LOGO_FILENAME)}'.")
            error_label = ctk.CTkLabel(self.splash, text=f"Error: Logo not found!\nPlace {LOGO_FILENAME}\n in the application folder.", text_color="red")
            error_label.pack(pady=20, padx=20, expand=True, fill="both")
        except Exception as e:
             print(f"Error loading logo: {e}")
             error_label = ctk.CTkLabel(self.splash, text=f"Error loading logo:\n{e}", text_color="red")
             error_label.pack(pady=20, padx=20, expand=True, fill="both")

        # Changed the function called after splash delay
        self.splash.after(3000, self.close_splash_and_prompt_id)

    def close_splash_and_prompt_id(self):
        """ Close splash and show the ID prompt """
        if self.splash:
            self.splash.destroy()
            self.splash = None
        # Don't show main window yet, show prompt first
        self.prompt_for_assessment_id()

    def prompt_for_assessment_id(self):
        """ Prompts the user for an ID using CTkInputDialog """
        dialog = ctk.CTkInputDialog(text="Please enter a Assessment ID:", title="Enter ID")
        # Center the dialog (approximation)
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        dialog_width = 300 # Estimated width
        dialog_height = 200 # Estimated height
        x = int((screen_width / 2) - (dialog_width / 2))
        y = int((screen_height / 2) - (dialog_height / 2))
        dialog.geometry(f"+{x}+{y}")

        entered_id = dialog.get_input()

        if entered_id:
            self.assessment_id = entered_id
            print(f"Assessment ID: {self.assessment_id}")
        else:
            self.assessment_id = "Not Provided" # Handle cancel or empty input
            print("User cancelled or provided no ID.")

        # Now show the main window and start checks
        self.deiconify() # Show main window
        self.lift()
        self.setup_main_window()
        self.start_checks(self.assessment_id) # Pass the collected ID

    def setup_main_window(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(self, text="Initializing...", wraplength=400, justify="center")
        self.status_label.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        self.process_gui_queue()

    def start_checks(self, assessment_id: str): # Accept assessment_id
        """ Starts the background check thread, passing the user ID """
        self.status_label.configure(text=f"Starting readiness checks for ID: {assessment_id}...")
        # Pass assessment_id to the worker thread
        self.worker_thread = threading.Thread(target=worker_thread_task, args=(assessment_id,), daemon=True)
        self.worker_thread.start()

    def process_gui_queue(self):
        # (This function remains the same as before)
        try:
            message = gui_queue.get_nowait()
            print(f"GUI Queue Received: {message}")

            if message.startswith("STATUS:"):
                self.status_label.configure(text=message[len("STATUS:"):].strip())
            elif message.startswith("SUCCESS:"):
                self.show_popup("Success", message[len("SUCCESS:"):].strip(), exit_on_close=True)
                # self.after(5000, self.on_closing) # Auto-close handled by popup button now
            elif message.startswith("ERROR:"):
                self.show_popup("Error", message[len("ERROR:"):].strip(), error=True, exit_on_close=True)
                # self.after(8000, self.on_closing) # Auto-close handled by popup button now
            elif message == "TASK_COMPLETE":
                pass

        except queue.Empty:
            pass
        finally:
            self.after(100, self.process_gui_queue)

    def show_popup(self, title, message, error=False, exit_on_close=False):
        # (This function remains mostly the same, but OK button action changes)
        popup = ctk.CTkToplevel(self)
        popup.title(title)
        popup.geometry("400x150")
        popup.resizable(False, False)
        popup.grab_set()
        popup.attributes("-topmost", True)

        main_x, main_y = self.winfo_x(), self.winfo_y()
        main_w, main_h = self.winfo_width(), self.winfo_height()
        p_w, p_h = 400, 150
        x = main_x + (main_w // 2) - (p_w // 2)
        y = main_y + (main_h // 2) - (p_h // 2)
        popup.geometry(f"{p_w}x{p_h}+{x}+{y}")

        icon_text = "!" if error else "i"
        icon_color = "red" if error else "green"

        icon_label = ctk.CTkLabel(popup, text=icon_text, font=("Arial", 24), text_color=icon_color)
        icon_label.pack(side="left", padx=15, anchor="n", pady=15)

        message_label = ctk.CTkLabel(popup, text=message, wraplength=300, justify="left")
        message_label.pack(side="left", padx=(0,10), pady=15, expand=True, fill="both")

        button_frame = ctk.CTkFrame(popup, fg_color="transparent")
        button_frame.pack(side="bottom", fill="x", pady=(0,10))

        # --- Modified Button Action ---
        if exit_on_close:
            # If it's a final message, OK button closes the whole app
            ok_button = ctk.CTkButton(button_frame, text="OK", width=80, command=self.on_closing)
        else:
            # Otherwise, just close the popup
            ok_button = ctk.CTkButton(button_frame, text="OK", width=80, command=popup.destroy)

        ok_button.pack()


    def on_closing(self):
        print("Closing application...")
        self.destroy()

# --- Main Execution ---
if __name__ == "__main__":
    # Need to initialize COM library for the WUA API if running in threads
    # Do this once at the start if pywin32 is available
    if _wuapi_available:
        try:
            import pythoncom
            pythoncom.CoInitialize()
            # Note: Ideally CoUninitialize should be called on exit,
            # but in a simple script like this, it might be omitted.
            # For robustness, consider adding it to on_closing or using atexit.
        except Exception as e:
            print(f"Failed to initialize COM: {e}. Update check might fail.")


    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    app = App()
    app.mainloop()

