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

# --- Configuration ---
APP_NAME = "Win11 Readiness Check"
LOGO_FILENAME = "logo.png"
# --- !!! REPLACE WITH YOUR ACTUAL API ENDPOINT !!! ---
API_ENDPOINT_URL = "YOUR_API_ENDPOINT_HERE"
# --- !!! REPLACE WITH YOUR SUPPORT PHONE NUMBER !!! ---
SUPPORT_PHONE_NUMBER = "1-800-555-HELP"
# --- End Configuration ---

gui_queue = queue.Queue()

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

def get_wmi_info():
    """ Attempts to get WMI information (TPM, Secure Boot, Graphics) """
    wmi_data = {
        "tpm_present": "Undetermined",
        "tpm_version": "Undetermined",
        "tpm_enabled": "Undetermined",
        "secure_boot_enabled": "Undetermined",
        "graphics_card": "Undetermined",
        "wddm_version": "Undetermined"
    }
    permission_error_flag = "(Permissions?)"

    try:
        import wmi # (pip install wmi)
        c = wmi.WMI()
        update_status("Querying WMI for TPM...")
        try:
            # This initial query usually works even without admin
            tpm_info_list = c.Win32_Tpm()
            if tpm_info_list:
                tpm_info = tpm_info_list[0]
                wmi_data["tpm_present"] = True
                try:
                    wmi_data["tpm_version"] = tpm_info.SpecVersion or "Unknown"
                except AttributeError:
                    wmi_data["tpm_version"] = "Unknown"

                # Checking specific enabled status often requires admin
                try:
                    # Attempt to check properties that might require elevation
                    # Use hasattr for safety, as properties might not exist
                    enabled = tpm_info.IsEnabled() if hasattr(tpm_info, 'IsEnabled') else None
                    activated = tpm_info.IsActivated() if hasattr(tpm_info, 'IsActivated') else None
                    # Simple logic: if we can read IsEnabled, report it. Otherwise, suspect permissions.
                    # A full implementation might need more checks.
                    if enabled is not None:
                         wmi_data["tpm_enabled"] = bool(enabled)
                    else:
                         # Could not read IsEnabled, likely permissions or property missing
                         wmi_data["tpm_enabled"] = f"Check Failed {permission_error_flag}"
                except Exception as e:
                    # Catch potential COM errors or access denied
                    print(f"WMI TPM status check failed: {e}")
                    wmi_data["tpm_enabled"] = f"Check Failed {permission_error_flag}"
            else:
                wmi_data["tpm_present"] = False
        except Exception as e:
            # Handle failure to query Win32_Tpm class itself
            print(f"WMI TPM query failed: {e}")
            wmi_data["tpm_present"] = f"Query Failed {permission_error_flag}"
            wmi_data["tpm_enabled"] = f"Query Failed {permission_error_flag}" # If base query fails, status also failed


        update_status("Querying WMI for Secure Boot...")
        try:
            # This query often requires admin rights for definitive results
            sb_info_list = c.Win32_SecureBoot()
            if sb_info_list:
                 # Property names can vary, check common ones
                 if hasattr(sb_info_list[0], 'SecureBootEnabled'):
                     wmi_data["secure_boot_enabled"] = bool(sb_info_list[0].SecureBootEnabled)
                 elif hasattr(sb_info_list[0], 'IsEnabled'): # Fallback check
                      wmi_data["secure_boot_enabled"] = bool(sb_info_list[0].IsEnabled)
                 else:
                      wmi_data["secure_boot_enabled"] = "Property Not Found"
            else:
                 wmi_data["secure_boot_enabled"] = "Query Failed (Class Not Found?)"
        except Exception as e:
            # Catch permission errors or other WMI failures
            print(f"WMI Secure Boot query failed: {e}")
            # Try registry as fallback (also often needs admin)
            try:
                import winreg
                key_path = r"SYSTEM\CurrentControlSet\Control\SecureBoot\State"
                # Use KEY_READ access explicitly
                reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ)
                value, _ = winreg.QueryValueEx(reg_key, "UEFISecureBootEnabled")
                winreg.CloseKey(reg_key)
                wmi_data["secure_boot_enabled"] = bool(value)
            except FileNotFoundError:
                 print("Secure Boot registry key not found.")
                 wmi_data["secure_boot_enabled"] = "Check Failed (Key Missing)"
            except PermissionError:
                 print("Permission denied reading Secure Boot registry key.")
                 wmi_data["secure_boot_enabled"] = f"Check Failed {permission_error_flag}"
            except Exception as reg_e:
                print(f"Registry Secure Boot check failed: {reg_e}")
                wmi_data["secure_boot_enabled"] = f"Check Failed {permission_error_flag}"


        update_status("Querying WMI for Graphics...")
        try:
            # Graphics info is usually readable by standard users
            gpu_info = c.Win32_VideoController()[0]
            wmi_data["graphics_card"] = gpu_info.Name
            # Simplified WDDM check using driver version
            wmi_data["wddm_version"] = f"Driver: {gpu_info.DriverVersion}"
        except Exception as e:
            print(f"WMI Graphics query failed: {e}")
            wmi_data["graphics_card"] = "Query Failed"
            wmi_data["wddm_version"] = "Query Failed"

    except ImportError:
        update_status("WMI module not found. Skipping WMI checks.")
        print("Python WMI module not installed (pip install wmi)")
        for key in wmi_data.keys(): wmi_data[key] = "WMI Module Missing"
    except Exception as e:
        # Catch potential COM errors during WMI initialization
        update_status(f"WMI initialization failed: {e}. Skipping WMI checks.")
        print(f"WMI failed to initialize: {e}")
        for key in wmi_data.keys(): wmi_data[key] = f"WMI Init Failed {permission_error_flag}"

    return wmi_data


def collect_system_data():
    """ Collects system information relevant to Win11 compatibility """
    data = {} # Initialize empty dictionary
    update_status("Collecting basic system info...")
    data["hostname"] = socket.gethostname()
    data["os_platform"] = platform.system()
    data["os_version"] = platform.version()
    data["os_release"] = platform.release()
    data["architecture"] = platform.machine()
    data["processor"] = platform.processor()
    # Add current time and timezone using current standard library features
    # Based on current time Wednesday, April 16, 2025 at 3:17:24 PM EDT
    # EDT is UTC-4
    current_time_utc = time.time()
    current_datetime_local = time.localtime(current_time_utc)
    timezone_name = time.tzname[current_datetime_local.tm_isdst] # Get current timezone name (e.g., 'EDT')
    timezone_offset_seconds = -time.timezone if not current_datetime_local.tm_isdst else -time.altzone
    timezone_offset_hours = timezone_offset_seconds / 3600
    timezone_offset_str = f"{int(timezone_offset_hours):+03d}:00" # Format as +/-HH:00

    data["timestamp_utc"] = time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime(current_time_utc))
    data["timestamp_local"] = time.strftime("%Y-%m-%d %H:%M:%S", current_datetime_local)
    data["timezone_name"] = timezone_name
    data["timezone_offset_utc"] = timezone_offset_str


    update_status("Collecting RAM info...")
    try:
        ram = psutil.virtual_memory()
        data["ram_total_gb"] = round(ram.total / (1024**3), 2)
    except Exception as e:
        print(f"RAM collection failed: {e}")
        data["ram_total_gb"] = "Error"

    update_status("Collecting Disk info (System Drive)...")
    try:
        system_drive = os.getenv("SystemDrive", "C:") + "\\"
        disk = psutil.disk_usage(system_drive)
        data["disk_total_gb"] = round(disk.total / (1024**3), 2)
        data["disk_free_gb"] = round(disk.free / (1024**3), 2)
    except Exception as e:
        print(f"Disk collection failed: {e}")
        data["disk_total_gb"] = "Error"
        data["disk_free_gb"] = "Error"

    update_status("Collecting WMI specific info (TPM, Secure Boot, Graphics)...")
    wmi_results = get_wmi_info()
    data.update(wmi_results)

    update_status("Data collection complete.")
    return data

def send_data_to_api(data):
    """ Sends collected data to the API endpoint """
    update_status(f"Sending data for {data.get('hostname', 'Unknown Host')}...")
    headers = {'Content-Type': 'application/json'}
    try:
        response = requests.post(API_ENDPOINT_URL, headers=headers, json=data, timeout=15)
        response.raise_for_status()
        update_status("Data sent successfully.")
        return True, "Success", data.get('hostname', 'Unknown Host')
    except requests.exceptions.Timeout:
        update_status("Error: Connection timed out.")
        return False, "Connection timed out.", data.get('hostname', 'Unknown Host')
    except requests.exceptions.RequestException as e:
        update_status(f"Error sending data: {e}")
        print(f"API Send Error: {e}")
        return False, f"Could not connect to server: {e}", data.get('hostname', 'Unknown Host')
    except Exception as e:
        update_status(f"An unexpected error occurred during sending: {e}")
        print(f"Unexpected Send Error: {e}")
        return False, f"An unexpected error occurred: {e}", data.get('hostname', 'Unknown Host')

def save_data_to_csv(data):
    """ Saves collected data to a CSV file in the temp directory """
    if not data:
        print("No data to save.")
        return
    try:
        temp_dir = tempfile.gettempdir()
        # Include timestamp in filename for uniqueness
        filename = os.path.join(temp_dir, f"readiness_check_{data.get('hostname', 'unknown')}_{time.strftime('%Y%m%d_%H%M%S')}.csv")
        update_status(f"Saving data locally to {filename}...")

        # Dynamically create headers from collected data keys
        fieldnames = sorted(list(data.keys()))

        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            # Get method safely handles if a key somehow wasn't populated
            row_data = {key: data.get(key, '') for key in fieldnames}
            writer.writerow(row_data)
        update_status(f"Data saved locally to {filename}")
        print(f"Data saved to {filename}")
    except Exception as e:
        update_status(f"Error saving data locally: {e}")
        print(f"Failed to save CSV: {e}")


def worker_thread_task():
    """ The main task performed in the background thread (no admin flag needed) """
    collected_data = {}
    try:
        collected_data = collect_system_data()
        success, status_msg, hostname = send_data_to_api(collected_data)

        if success:
            show_final_message("SUCCESS", f"Assessment data for '{hostname}' has been submitted successfully.")
        else:
            show_final_message("ERROR", f"There was an issue submitting the data ({status_msg}).\nPlease call us at {SUPPORT_PHONE_NUMBER}.\n\nResults are being saved locally.")
            save_data_to_csv(collected_data)

    except Exception as e:
        print(f"Critical error in worker thread: {e}")
        update_status(f"Critical error: {e}")
        show_final_message("ERROR", f"A critical error occurred: {e}\nPlease call us at {SUPPORT_PHONE_NUMBER}.")
        if collected_data:
             save_data_to_csv(collected_data)

    gui_queue.put("TASK_COMPLETE")


# --- GUI Class (Simplified) ---

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(APP_NAME)
        self.geometry("450x200") # Reduced height slightly
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.logo_image = None
        self.status_label = None
        self.worker_thread = None

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
            print(f"Loading logo from: {logo_path}")
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

        self.splash.after(3000, self.close_splash_and_start) # 3 seconds delay

    def close_splash_and_start(self):
        if self.splash:
            self.splash.destroy()
            self.splash = None
        self.deiconify() # Show main window
        self.lift()
        self.setup_main_window() # Setup the main window UI elements
        self.start_checks() # Immediately start the checks

    def setup_main_window(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(self, text="Initializing...", wraplength=400, justify="center")
        self.status_label.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        self.process_gui_queue()

    def start_checks(self):
        self.status_label.configure(text="Starting readiness checks...")
        self.worker_thread = threading.Thread(target=worker_thread_task, daemon=True) # Pass no args
        self.worker_thread.start()

    def process_gui_queue(self):
        try:
            message = gui_queue.get_nowait()
            print(f"GUI Queue Received: {message}")

            if message.startswith("STATUS:"):
                self.status_label.configure(text=message[len("STATUS:"):].strip())
            elif message.startswith("SUCCESS:"):
                self.show_popup("Success", message[len("SUCCESS:"):].strip())
                self.after(5000, self.on_closing)
            elif message.startswith("ERROR:"):
                self.show_popup("Error", message[len("ERROR:"):].strip(), error=True)
                self.after(8000, self.on_closing)
            elif message == "TASK_COMPLETE":
                pass # Handled by SUCCESS/ERROR

        except queue.Empty:
            pass
        finally:
            self.after(100, self.process_gui_queue)

    def show_popup(self, title, message, error=False): # TODO - Add 'close on ok' param to exit the program and remove auto close
        popup = ctk.CTkToplevel(self)
        popup.title(title)
        # Adjust size based on message length? For now, fixed.
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
        icon_label.pack(side="left", padx=15, anchor="n", pady=15) # Anchor top

        message_label = ctk.CTkLabel(popup, text=message, wraplength=300, justify="left")
        message_label.pack(side="left", padx=(0,10), pady=15, expand=True, fill="both")

        # Button frame to keep button at bottom
        button_frame = ctk.CTkFrame(popup, fg_color="transparent")
        button_frame.pack(side="bottom", fill="x", pady=(0,10))

        ok_button = ctk.CTkButton(button_frame, text="OK", width=80, command=popup.destroy)
        ok_button.pack() # Center button at the bottom


    def on_closing(self):
        print("Closing application...")
        self.destroy()

# --- Main Execution ---
if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    app = App()
    app.mainloop()

# TODO - Check if there are any pending installs