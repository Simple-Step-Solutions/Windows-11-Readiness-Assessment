import customtkinter as ctk
from PIL import Image, ImageTk
from config import *
from system_info import SystemInfo
import platform
import psutil # Already imported, needed for service check
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
# Windows Update Agent COM API & COM utilities
try:
    import win32com.client
    import pythoncom # Needed for CoInitialize/CoUninitialize
    _wuapi_available = True
except ImportError:
    _wuapi_available = False
    # Define pythoncom as None if import fails to avoid NameError later
    pythoncom = None


gui_queue = queue.Queue()

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

def check_wmi_service() -> (bool, str):
    """ Checks if the WMI service ('Winmgmt') is running. """
    try:
        service = psutil.win_service_get('Winmgmt')
        status = service.status()
        if status == 'running':
            return True, "Running"
        else:
            return False, f"Service status: {status}"
    except psutil.NoSuchProcess:
        return False, "Service not found (NoSuchProcess)"
    except Exception as e:
        return False, f"Error checking service: {e}"

def populate_wmi_info(info: SystemInfo):
    """ Attempts to get WMI info and updates the SystemInfo object. """
    update_status("Checking WMI Service Status...")
    wmi_service_ok, wmi_service_status = check_wmi_service()
    if not wmi_service_ok:
        error_msg = f"WMI Service ('Winmgmt') not running or inaccessible. Status: {wmi_service_status}"
        update_status(error_msg)
        print(error_msg)
        info.wmi_error_details = error_msg
        # Set all WMI fields to indicate service error
        info.tpm_present = "WMI Service Error"
        info.tpm_version = "WMI Service Error"
        info.tpm_enabled = "WMI Service Error"
        info.secure_boot_enabled = "WMI Service Error"
        info.graphics_card = "WMI Service Error"
        info.wddm_version = "WMI Service Error"
        info.ram_type = "WMI Service Error"
        info.system_drive_type = "WMI Service Error"
        info.ram_speed = "WMI Service Error"
        return

    if not _wmi_available:
        update_status("WMI module not found. Skipping WMI checks.")
        info.tpm_present = "WMI Module Missing"
        info.tpm_version = "WMI Module Missing"
        info.tpm_enabled = "WMI Module Missing"
        info.secure_boot_enabled = "WMI Module Missing"
        info.graphics_card = "WMI Module Missing"
        info.wddm_version = "WMI Module Missing"
        info.ram_type = "WMI Module Missing"
        info.system_drive_type = "WMI Module Missing"
        info.ram_speed = "WMI Module Missing"
        return

    permission_error_flag = "(Permissions?)"
    wmi_errors = []

    # --- Open WMI Namespaces ---
    try:
        # Default WMI connection (usually ROOT\cimv2)
        c = wmi.WMI()
        # Connection for Storage namespace (needed for MSFT_PhysicalDisk)
        c_storage = None
        try:
            # This namespace might require admin rights
            c_storage = wmi.WMI(namespace="root/Microsoft/Windows/Storage")
            update_status("Connected to WMI Storage namespace.")
        except Exception as e_storage_con:
            err = f"WMI Storage Namespace Connect Failed: {type(e_storage_con).__name__}: {e_storage_con}"
            print(err)
            wmi_errors.append(err)
            info.system_drive_type = f"WMI Storage Connect Failed {permission_error_flag}"
    except Exception as e:
        error_msg = f"WMI initialization failed: {type(e).__name__}: {e}"
        update_status(error_msg)
        print(error_msg)
        info.wmi_error_details = error_msg
        # Set all WMI fields to indicate init error
        info.tpm_present = f"WMI Init Failed {permission_error_flag}"
        info.tpm_version = f"WMI Init Failed {permission_error_flag}"
        info.tpm_enabled = f"WMI Init Failed {permission_error_flag}"
        info.secure_boot_enabled = f"WMI Init Failed {permission_error_flag}"
        info.graphics_card = f"WMI Init Failed {permission_error_flag}"
        info.wddm_version = f"WMI Init Failed {permission_error_flag}"
        info.ram_type = f"WMI Init Failed {permission_error_flag}"
        info.system_drive_type = f"WMI Init Failed {permission_error_flag}"
        return

    # --- TPM Check (using default connection 'c') ---
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
                err = f"TPM Status Check Failed: {type(e_detail).__name__}: {e_detail}"
                print(err)
                wmi_errors.append(err)
                info.tpm_enabled = f"Check Failed {permission_error_flag}"
        else:
            info.tpm_present = False
            info.tpm_version = "N/A"
            info.tpm_enabled = "N/A"
    except Exception as e:
        err = f"TPM Query Failed: {type(e).__name__}: {e}"
        print(err)
        wmi_errors.append(err)
        info.tpm_present = f"Query Failed {permission_error_flag}"
        info.tpm_version = f"Query Failed {permission_error_flag}"
        info.tpm_enabled = f"Query Failed {permission_error_flag}"


    # --- Secure Boot Check (using default connection 'c') ---
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
        err = f"Secure Boot WMI Query Failed: {type(e).__name__}: {e}"
        print(err)
        wmi_errors.append(err)
        # Try registry fallback
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
                err_reg = f"Secure Boot Registry Check Failed: {type(reg_e).__name__}: {reg_e}"
                print(err_reg)
                wmi_errors.append(err_reg)
                info.secure_boot_enabled = f"Check Failed {permission_error_flag}"
        else: info.secure_boot_enabled = f"Check Failed (WMI Error & WinReg Missing)"


    # --- Graphics Check (using default connection 'c') ---
    update_status("Querying WMI for Graphics...")
    info.graphics_card = "Undetermined"
    info.wddm_version = "Undetermined"
    try:
        gpu_info = c.Win32_VideoController()[0]
        info.graphics_card = gpu_info.Name
        info.wddm_version = f"Driver: {gpu_info.DriverVersion}"
    except Exception as e:
        err = f"Graphics Query Failed: {type(e).__name__}: {e}"
        print(err)
        wmi_errors.append(err)
        info.graphics_card = "Query Failed"
        info.wddm_version = "Query Failed"

    # --- RAM Type Check (using default connection 'c') ---
    update_status("Querying WMI for RAM Type...")
    info.ram_type = "Undetermined"
    try:
        for mem in c.Win32_PhysicalMemory():
            smbios_memory_type = mem.SMBIOSMemoryType
            # https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/win32-physicalmemory
            type_mapping = {
                0: "Unknown",
                1: "Other",
                2: "DRAM",
                3: "Synchronous DRAM",
                4: "Cache DRAM",
                5: "EDO",
                6: "EDRAM",
                7: "VRAM",
                8: "SRAM",
                9: "RAM",
                10: "ROM",
                11: "Flash",
                12: "EEPROM",
                13: "FEPROM",
                14: "EPROM",
                15: "CDRAM",
                16: "3DRAM",
                17: "SDRAM",
                18: "SGRAM",
                19: "RDRAM",
                20: "DDR",
                21: "DDR-2",
                22: "BRAM",
                23: "FB-DIMM",
                24: "DDR3",
                25: "FBD2",
                26: "DDR4",
                27: "LPDDR",
                28: "LPDDR2",
                29: "LPDDR3",
                30: "LPDDR4",
                31: "Logical non-volatile device",
                32: "HBM (High Bandwidth Memory)",
                33: "HBM2 (High Bandwidth Memory Generation 2)",
                34: "DDR5",
                35: "LPDDR5",
                36: "HBM3 (High Bandwidth Memory Generation 3)",
            }
            info.ram_type = type_mapping.get(smbios_memory_type, "Undetermined")
    except Exception as e:
        err = f"RAM Type Query Failed: {type(e).__name__}: {e}"
        print(err); wmi_errors.append(err)
        info.ram_type = f"Query Failed {permission_error_flag}"

    # --- RAM Speed Check (using default connection 'c') ---
    info.ram_speed = "Undetermined"
    try:
        c = wmi.WMI()
        speed = None
        speed_unit = "MHz"  # Default unit
        for mem in c.Win32_PhysicalMemory():
            speed = mem.Speed
            if speed is not None:
                info.ram_speed = f"{speed}{speed_unit}"
    except Exception as e:
        err = f"RAM Type Query Failed: {type(e).__name__}: {e}"
        print(err); wmi_errors.append(err)
        info.ram_type = f"Query Failed {permission_error_flag}"

    # --- Drive Type Check (using storage connection 'c_storage') ---
    update_status("Querying WMI for System Drive Type...")
    info.system_drive_type = "Undetermined"
    if c_storage: # Only proceed if connection to storage namespace succeeded
        try:
            system_drive_letter = os.getenv("SystemDrive", "C:")
            # Query physical disks
            # MediaType: 3=HDD, 4=SSD, 5=SCM, 0=Unspecified
            disk_types = {3: "HDD", 4: "SSD", 5: "SCM", 0: "Unspecified"}

            # Find the physical disk associated with the system drive letter
            # This mapping can be complex. Simplified approach: check all physical disks.
            # A robust solution might involve Win32_LogicalDiskToPartition etc.
            # Assumption: System drive is usually on the first or primary physical disk reported.

            model = None

            try:
                c = wmi.WMI()
                for drive in c.Win32_DiskDrive():
                    for partition in c.Win32_DiskPartition(DiskIndex=drive.Index):
                        # Correctly link using the partition's index and the disk's DeviceID
                        for logical_disk in c.Win32_LogicalDisk():
                            if logical_disk.DeviceID == system_drive_letter:
                                model = drive.Model
            except wmi.x_wmi as e:
                print(f"WMI Error: {e}")
            except Exception as e:
                print(f"An unexpected error occurred: {e}")

            if model:
                for d in c_storage.MSFT_PhysicalDisk():
                    if model == d.Model:
                        media_type_code = d.MediaType
                        info.system_drive_type = disk_types.get(media_type_code, f"Unknown Code ({media_type_code})")

        except Exception as e:
            print("ERR", e)
            err = f"Drive Type Query Failed (MSFT_PhysicalDisk): {type(e).__name__}: {e}"
            print(err); wmi_errors.append(err)
            # Fallback or set specific error
            info.system_drive_type = f"Query Failed {permission_error_flag}"
    elif info.system_drive_type != f"WMI Init Failed {permission_error_flag}" and \
         info.system_drive_type != "WMI Service Error": # Avoid overwriting previous errors
        # If storage connection failed earlier, reflect that
        info.system_drive_type = "WMI Storage Connect Failed"


    # Store collected WMI errors if any occurred
    if wmi_errors:
        # Append to existing details if WMI init failed earlier
        existing_details = info.wmi_error_details if info.wmi_error_details else ""
        new_details = "; ".join(wmi_errors)
        info.wmi_error_details = f"{existing_details}; {new_details}".strip("; ")


def check_pending_updates(info: SystemInfo):
    """ Checks for pending Windows Updates using WUA API. """
    update_status("Checking for pending Windows Updates...")
    info.pending_updates_count = -1 # Default to general error
    info.update_check_error_details = None

    if not _wuapi_available:
        update_status("Windows Update check skipped: pywin32 module missing.")
        info.pending_updates_count = -2
        info.update_check_error_details = "pywin32 module not found"
        return

    try:
        update_session = win32com.client.Dispatch("Microsoft.Update.Session")
        update_searcher = update_session.CreateUpdateSearcher()
        search_criteria = "IsInstalled=0 and IsHidden=0 and Type='Software'"
        update_status("Searching for available updates (this may take a moment)...")
        search_result = update_searcher.Search(search_criteria)
        count = search_result.Updates.Count
        update_status(f"Found {count} applicable updates.")
        info.pending_updates_count = count

    except pythoncom.com_error as com_err:
         err_msg = f"COM Error HRESULT={com_err.hresult}: {com_err}"
         print(f"Windows Update check failed (COM Error): {err_msg}")
         update_status(f"Windows Update check failed (COM Error)")
         info.pending_updates_count = -3
         info.update_check_error_details = err_msg
    except Exception as e:
        err_msg = f"{type(e).__name__}: {e}"
        print(f"Windows Update check failed: {err_msg}")
        update_status(f"Windows Update check failed: {e}")
        info.pending_updates_count = -1
        info.update_check_error_details = err_msg

def collect_system_data(assessment_id: str) -> SystemInfo:
    """ Collects system information and returns a populated SystemInfo object. """
    info = SystemInfo()
    info.assessment_id = assessment_id

    try:
        update_status("Collecting basic system info...")
        info.hostname = socket.gethostname()
        info.os_platform = platform.system()
        info.os_version = platform.version()
        info.os_release = platform.release()
        info.architecture = platform.machine()
        info.processor = platform.processor() # Basic processor name

        # CPU Details
        update_status("Collecting CPU details...")
        try:
            info.cpu_physical_cores = psutil.cpu_count(logical=False)
            info.cpu_logical_cores = psutil.cpu_count(logical=True)
            cpu_freq = psutil.cpu_freq()
            # Use max freq if available, otherwise current freq
            max_freq = cpu_freq.max if cpu_freq.max > 0 else cpu_freq.current
            info.cpu_max_speed_ghz = round(max_freq / 1000, 2) if max_freq > 0 else 0.0
        except Exception as e:
            print(f"CPU detail collection failed: {e}")
            # Defaults remain 0

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
        populate_wmi_info(info) # Includes RAM Type and Drive Type now

        # --- Windows Update Check ---
        check_pending_updates(info)

        update_status("Data collection complete.")

    except Exception as e:
        print(f"CRITICAL ERROR during data collection: {e}")
        info.collection_error = str(e)
        update_status(f"Critical error during collection: {e}")

    return info

# --- Data Handling Functions ---
# (send_data_to_api and save_data_to_csv remain the same)
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
        assessment_id_part = info.assessment_id.replace(" ", "_").replace("/", "-").replace("\\", "-") if info.assessment_id else "no_id" # Sanitize ID for filename
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

def worker_thread_task(assessment_id: str):
    """ The main task performed in the background thread """
    system_info = None
    coinitialized = False # Flag to track COM initialization
    try:
        # --- Initialize COM for this thread ---
        if _wuapi_available and pythoncom:
            try:
                pythoncom.CoInitialize()
                coinitialized = True
            except Exception as com_init_e:
                print(f"Failed to CoInitialize COM for thread: {com_init_e}")
                update_status("Error initializing COM for Update check.")

        system_info = collect_system_data(assessment_id)

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
        # --- Uninitialize COM for this thread ---
        if coinitialized and pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception as com_uninit_e:
                print(f"Failed to CoUninitialize COM for thread: {com_uninit_e}")
        gui_queue.put("TASK_COMPLETE")


# --- GUI Class ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(APP_NAME)
        # Set window icon using iconphoto (requires Pillow)
        self.app_icon = None # Keep reference if needed elsewhere, though not strictly for iconbitmap
        try:
            icon_path = get_resource_path(ICON_FILENAME) # Use logo.ico
            self.iconbitmap(icon_path) # Set taskbar/window icon using ICO
            print(f"Successfully set window icon from {icon_path}")
        except Exception as e:
            print(f"Error setting window icon using iconbitmap: {e}")
            # Fallback attempt using iconphoto with PNG (less reliable)
            try:
                print("Attempting fallback icon using iconphoto with PNG...")
                png_icon_path = get_resource_path(LOGO_FILENAME)
                pil_icon = Image.open(png_icon_path)
                self.app_icon = ImageTk.PhotoImage(pil_icon) # Need to store PhotoImage
                self.iconphoto(True, self.app_icon)
                print(f"Successfully set window icon using iconphoto fallback.")
            except Exception as e_photo:
                 print(f"Error setting window icon using iconphoto fallback: {e_photo}")

        self.geometry("450x200")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.logo_image = None
        self.status_label = None
        self.worker_thread = None
        self.assessment_id = "Not Provided" # Store user ID here

        self.show_splash_screen()

    def show_splash_screen(self):
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
        dialog = ctk.CTkInputDialog(text="Please enter an assessment ID:", title="Enter ID")
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
            print(f"User entered ID: {self.assessment_id}")
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
            elif message.startswith("ERROR:"):
                self.show_popup("Error", message[len("ERROR:"):].strip(), error=True, exit_on_close=True)
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
            ok_button = ctk.CTkButton(button_frame, text="OK", width=80, command=self.on_closing)
        else:
            ok_button = ctk.CTkButton(button_frame, text="OK", width=80, command=popup.destroy)

        ok_button.pack()


    def on_closing(self):
        print("Closing application...")
        self.destroy()

# --- Main Execution ---
if __name__ == "__main__":
    # Removed CoInitialize from main thread - it's now handled per-thread
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    app = App()
    app.mainloop()
