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
import subprocess
import logging
import re

# create logger with 'spam_application'
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
# create file handler which logs even debug messages
fh = logging.FileHandler(os.path.join(tempfile.gettempdir(), f"win11_readiness__{time.strftime('%Y%m%d_%H%M%S')}.log"))
fh.setLevel(logging.INFO)
# create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.ERROR)
# create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)
# add the handlers to the logger
logger.addHandler(fh)
logger.addHandler(ch)

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

def run_command(command, timeout=15):
    """Runs a command and returns its stdout, stderr, and return code."""
    try:
        result = subprocess.run(command, capture_output=True, text=True, shell=True, timeout=timeout, check=False)
        # Be cautious with shell=True if command parts come from untrusted input.
        # For fixed commands here, it's generally okay and simplifies things.
        return result.stdout, result.stderr, result.returncode
    except FileNotFoundError:
        logger.error(f"Command not found: {command.split()[0]}")
        return None, f"Command not found: {command.split()[0]}", -1
    except subprocess.TimeoutExpired:
        logger.error(f"Command timed out: {command}")
        return None, f"Command timed out after {timeout} seconds", -1
    except Exception as e:
        logger.error(f"Error running command '{command}': {e}")
        return None, f"Unexpected error: {e}", -1


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
        logger.error(error_msg)
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
            logger.error(err)
            wmi_errors.append(err)
            info.system_drive_type = f"WMI Storage Connect Failed {permission_error_flag}"

        # Connection for TPM namespace
        c_tpm = None
        try:
            # This namespace might require admin rights
            c_tpm = wmi.WMI(namespace="root/cimv2/security/microsofttpm")
            update_status("Connected to WMI TPM namespace.")
        except Exception as e_storage_con:
            err = f"WMI TPM Namespace Connect Failed: {type(e_storage_con).__name__}: {e_storage_con}"
            logger.error(err)
            wmi_errors.append(err)
            info.tpm_present = f"WMI TPM Connect Failed {permission_error_flag}"
            info.tpm_enabled = f"WMI TPM Connect Failed {permission_error_flag}"
            info.tpm_version = f"WMI TPM Connect Failed {permission_error_flag}"
    except Exception as e:
        error_msg = f"WMI initialization failed: {type(e).__name__}: {e}"
        update_status(error_msg)
        logger.error(error_msg)
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

    # --- Manufacturer Check (using default connection 'c') ---
    update_status("Querying WMI for Manufacturer...")
    info.manufacturer = "Undetermined"
    try:
        info.manufacturer = c.Win32_ComputerSystem()[0].Manufacturer
    except Exception as e:
        err = f"Manufacturer Query Failed: {type(e).__name__}: {e}"
        logger.error(err)
        wmi_errors.append(err)
        info.manufacturer = f"Query Failed {permission_error_flag}"

    # --- TPM Check (using tpm connection 'c_tpm') ---
    update_status("Querying WMI for TPM...")
    info.tpm_present = "Undetermined"
    info.tpm_version = "Undetermined"
    info.tpm_enabled = "Undetermined"
    try:
        tpm_info_list = c_tpm.Win32_Tpm()
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
                logger.error(err)
                wmi_errors.append(err)
                info.tpm_enabled = f"Check Failed {permission_error_flag}"
        else:
            info.tpm_present = False
            info.tpm_version = "N/A"
            info.tpm_enabled = "N/A"
    except Exception as e:
        err = f"TPM Query Failed: {type(e).__name__}: {e}"
        logger.error(err)
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
        logger.error(err)
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
                logger.error(err_reg)
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
        logger.error(err)
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
        logger.error(err); wmi_errors.append(err)
        info.ram_type = f"Query Failed {permission_error_flag}"

    # --- RAM Speed Check (using default connection 'c') ---
    update_status("Querying WMI for RAM Speed...")
    info.ram_speed_mhz = "Undetermined"
    try:
        c = wmi.WMI()
        speed = None
        for mem in c.Win32_PhysicalMemory():
            speed = mem.Speed
            if speed is not None:
                info.ram_speed_mhz = speed
    except Exception as e:
        err = f"RAM Type Query Failed: {type(e).__name__}: {e}"
        logger.error(err); wmi_errors.append(err)
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
                logger.error(f"WMI Error: {e}")
            except Exception as e:
                logger.error(f"An unexpected error occurred: {e}")

            if model:
                for d in c_storage.MSFT_PhysicalDisk():
                    if model == d.Model:
                        media_type_code = d.MediaType
                        info.system_drive_type = disk_types.get(media_type_code, f"Unknown Code ({media_type_code})")

        except Exception as e:
            logger.error("ERR", e)
            err = f"Drive Type Query Failed (MSFT_PhysicalDisk): {type(e).__name__}: {e}"
            logger.error(err); wmi_errors.append(err)
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
         logger.error(f"Windows Update check failed (COM Error): {err_msg}")
         update_status(f"Windows Update check failed (COM Error)")
         info.pending_updates_count = -3
         info.update_check_error_details = err_msg
    except Exception as e:
        err_msg = f"{type(e).__name__}: {e}"
        logger.error(f"Windows Update check failed: {err_msg}")
        update_status(f"Windows Update check failed: {e}")
        info.pending_updates_count = -1
        info.update_check_error_details = err_msg

def check_bitlocker_status(info: SystemInfo):
    """Checks BitLocker status for fixed drives using manage-bde."""
    update_status("Checking BitLocker status...")
    info.hipaa_bitlocker_status = "Error" # Default to error
    info.hipaa_bitlocker_details = {}
    overall_encrypted = True
    overall_unencrypted = True
    found_drives = False

    stdout, stderr, retcode = run_command("manage-bde -status")

    if retcode != 0 or stdout is None:
        error_msg = f"manage-bde failed. Return code: {retcode}. Stderr: {stderr}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['BitLocker'] = error_msg
        return

    try:
        # Process stdout - this parsing is complex and might need adjustment based on OS language/version
        current_drive = None
        details = {}
        lines = stdout.strip().split('\n')
        drive_pattern = re.compile(r"Volume ([A-Z]:)") # Simple pattern for drive letter

        for line in lines:
            line = line.strip()
            match = drive_pattern.search(line)
            if match:
                current_drive = match.group(1)
                if current_drive:
                    details[current_drive] = {
                        "ProtectionStatus": "Unknown",
                        "EncryptionMethod": "Unknown",
                        "EncryptionPercentage": -1,
                        "ConversionStatus": "Unknown", # Added for clarity
                        "LockStatus": "Unknown"
                    }
                    found_drives = True

            if current_drive and details.get(current_drive):
                if "Protection Status:" in line:
                     # Make case-insensitive and handle variations
                    status_val = line.split(":", 1)[1].strip().lower()
                    if "on" in status_val:
                       details[current_drive]["ProtectionStatus"] = "On"
                    elif "off" in status_val:
                       details[current_drive]["ProtectionStatus"] = "Off"

                elif "Conversion Status:" in line:
                    # E.g., "Fully Encrypted", "Used Space Only Encrypted", "Encryption in Progress"
                    status_val = line.split(":", 1)[1].strip()
                    details[current_drive]["ConversionStatus"] = status_val
                    if "encrypt" in status_val.lower() and "%" in status_val:
                        try:
                             # Extract percentage if present, e.g., "Encryption in Progress 55.5%"
                            percentage_match = re.search(r'(\d+(\.\d+)?)\s*%', status_val)
                            if percentage_match:
                                details[current_drive]["EncryptionPercentage"] = float(percentage_match.group(1))
                            else:
                                details[current_drive]["EncryptionPercentage"] = -2 # Indicate in progress but % not parsed
                        except ValueError:
                             details[current_drive]["EncryptionPercentage"] = -3 # Parse error

                    elif "fully encrypted" in status_val.lower() or "used space only encrypted" in status_val.lower():
                         details[current_drive]["EncryptionPercentage"] = 100.0

                    elif "fully decrypted" in status_val.lower():
                         details[current_drive]["EncryptionPercentage"] = 0.0
                    else:
                         details[current_drive]["EncryptionPercentage"] = -4 # Unknown status

                elif "Encryption Method:" in line:
                    details[current_drive]["EncryptionMethod"] = line.split(":", 1)[1].strip()

                elif "Lock Status:" in line:
                    details[current_drive]["LockStatus"] = line.split(":", 1)[1].strip()


        info.hipaa_bitlocker_details = details

        # Determine overall status based on Protection Status
        if not found_drives:
             info.hipaa_bitlocker_status = "No Fixed Drives Found/Parsed"
             return

        any_on = False
        any_off = False
        for drive, data in details.items():
            # Only consider fixed drives for overall status (heuristically check method)
            if data.get("EncryptionMethod", "Unknown") != "Unknown" and data.get("EncryptionMethod", "None") != "None":
                 if data.get("ProtectionStatus") == "On":
                      any_on = True
                 elif data.get("ProtectionStatus") == "Off":
                      any_off = True
                 else: # If status is unknown for any relevant drive, mark overall as mixed/error
                      any_off = True # Treat unknowns as potentially not secure


        if any_on and not any_off:
            info.hipaa_bitlocker_status = "Encrypted"
        elif any_off and not any_on:
            info.hipaa_bitlocker_status = "Not Encrypted"
        elif any_on and any_off:
            info.hipaa_bitlocker_status = "Mixed"
        elif not any_on and not any_off: # Should only happen if only non-fixed drives parsed or error
             info.hipaa_bitlocker_status = "Unknown/Error Parsing Status"
        else:
             info.hipaa_bitlocker_status = "Unknown State"


    except Exception as e:
        error_msg = f"Error parsing manage-bde output: {e}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['BitLocker'] = error_msg
        info.hipaa_bitlocker_status = "Error Parsing Output"

def check_audit_policy(info: SystemInfo):
    """Checks key audit policy settings using auditpol."""
    update_status("Checking audit policy...")
    stdout, stderr, retcode = run_command("auditpol /get /category:*")

    if retcode != 0 or stdout is None:
        error_msg = f"auditpol failed. Return code: {retcode}. Stderr: {stderr}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['AuditPolicy'] = error_msg
        info.hipaa_audit_logon_events = "Error"
        info.hipaa_audit_account_mgmt = "Error"
        info.hipaa_audit_policy_change = "Error"
        info.hipaa_audit_object_access = "Error"
        return

    policy_map = {
        "Logon/Logoff": "hipaa_audit_logon_events",
        "Account Management": "hipaa_audit_account_mgmt",
        "Policy Change": "hipaa_audit_policy_change",
        "Object Access": "hipaa_audit_object_access", # Note: This category has many subcategories
    }

    try:
        # Initialize defaults
        for field in policy_map.values():
            setattr(info, field, "No Auditing") # Default if category not found or not set

        # Parse auditpol output (example, may need refinement)
        lines = stdout.strip().split('\n')
        current_category = None
        for line in lines:
            line = line.strip()
            # Identify category lines (heuristic: check indentation or known names)
            if line and not line.startswith("  ") and line.endswith(":"): # Basic category detection
                 current_category_name = line[:-1] # Remove trailing colon
                 if current_category_name in policy_map:
                      current_category = current_category_name
                 else:
                      current_category = None # Reset if not a category we track

            # Check subcategory settings within tracked categories
            if current_category and line.startswith("  "):
                setting = "No Auditing" # Default for the subcategory line
                if "Success and Failure" in line:
                    setting = "Success and Failure"
                elif "Success" in line: # Must check Success *after* S+F
                    setting = "Success"
                elif "Failure" in line:
                    setting = "Failure"

                # Update the main category field - take the highest level found
                # (e.g., if any subcategory has S+F, mark the main category S+F)
                current_setting = getattr(info, policy_map[current_category])
                if setting == "Success and Failure":
                    setattr(info, policy_map[current_category], setting)
                elif setting == "Success" and current_setting != "Success and Failure":
                     setattr(info, policy_map[current_category], setting)
                elif setting == "Failure" and current_setting not in ["Success and Failure", "Success"]:
                     setattr(info, policy_map[current_category], setting)
                # Else: keep "No Auditing" or the existing higher setting

    except Exception as e:
        error_msg = f"Error parsing auditpol output: {e}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['AuditPolicyParse'] = error_msg
        # Reset fields to Error on parse failure
        for field in policy_map.values():
             setattr(info, field, "Error Parsing")

def check_security_log_settings(info: SystemInfo):
    """Checks Security event log size and retention policy."""
    update_status("Checking Security event log settings...")
    # Use wevtutil for reliable info
    stdout, stderr, retcode = run_command('wevtutil gl Security')

    if retcode != 0 or stdout is None:
        error_msg = f"wevtutil failed. Return code: {retcode}. Stderr: {stderr}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['SecurityLog'] = error_msg
        info.hipaa_audit_log_max_size_mb = -1
        info.hipaa_audit_log_retention = "Error"
        return

    try:
        max_size_bytes = -1
        retention = "Error"
        for line in stdout.strip().split('\n'):
            line = line.strip()
            if line.startswith("maxSize:"):
                try:
                    max_size_bytes = int(line.split(":")[1].strip())
                    info.hipaa_audit_log_max_size_mb = max_size_bytes // (1024 * 1024)
                except ValueError:
                    logger.error(f"Could not parse maxSize: {line}")
                    info.hipaa_audit_log_max_size_mb = -2 # Indicate parse error
            elif line.startswith("retention:"):
                retention_flag = line.split(":")[1].strip().lower()
                # Possible values: True (archive), False (overwrite)
                # Need to also check autoBackupLogFiles flag for "Do not overwrite"
                if retention_flag == 'true':
                     retention = "Archive the log when full" # Tentative, check autoBackup below
                elif retention_flag == 'false':
                     retention = "Overwrite as needed"
            elif line.startswith("autoBackup:"): # Older name? Check wevtutil docs if needed
                 pass # Often combined with retention
            elif line.startswith("autoBackupLogFiles:"): # Seems to be the key for "Do not overwrite"
                if line.split(":")[1].strip().lower() == 'true' and retention == "Archive the log when full":
                     # This combination usually means "Do not overwrite events (Clear log manually)" in the GUI
                     retention = "Do not overwrite events" # Or "Archive + AutoBackup" ? Needs verification.


        info.hipaa_audit_log_retention = retention
        if info.hipaa_audit_log_max_size_mb == -1 and 'maxSize' not in info.hipaa_checks_error_details.get('SecurityLog',''):
             info.hipaa_audit_log_max_size_mb = -3 # Indicate value not found in output
        if info.hipaa_audit_log_retention == "Error" and 'retention' not in info.hipaa_checks_error_details.get('SecurityLog',''):
             info.hipaa_audit_log_retention = "Value Not Found"


    except Exception as e:
        error_msg = f"Error parsing wevtutil output: {e}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['SecurityLogParse'] = error_msg
        info.hipaa_audit_log_max_size_mb = -1
        info.hipaa_audit_log_retention = "Error Parsing"

def check_account_policies(info: SystemInfo):
    """Checks password complexity, length, and lockout settings using net accounts."""
    update_status("Checking account policies (net accounts)...")
    stdout, stderr, retcode = run_command("net accounts")

    if retcode != 0 or stdout is None:
        error_msg = f"net accounts failed. Return code: {retcode}. Stderr: {stderr}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['AccountPolicies'] = error_msg
        info.hipaa_password_complexity = "Error"
        info.hipaa_min_password_length = -1
        info.hipaa_account_lockout_threshold = -1
        info.hipaa_account_lockout_duration_min = -1
        return

    try:
        # Set defaults for parsing
        info.hipaa_password_complexity = "Disabled" # Assume disabled unless found enabled
        info.hipaa_min_password_length = 0 # Default if not found or 'None'
        info.hipaa_account_lockout_threshold = 0 # Default is 'Never' which maps to 0 attempts
        info.hipaa_account_lockout_duration_min = -1 # Indicate not applicable if threshold is Never

        lines = stdout.strip().split('\n')
        for line in lines:
            # Password Complexity (heuristic check)
            # Note: 'Password complexity' might not appear directly. Check for related policy enforce lines.
            # A better check might involve 'secedit /export' which is more complex to parse.
            # 'net accounts' sometimes shows 'Lockout threshold: Never' but complexity might still be GPO-enforced.
            # This check is thus potentially unreliable for complexity via 'net accounts'.
            # Let's check the "Password must meet complexity requirements" line if present
            if "Password must meet complexity requirements" in line:
                 if "Enabled" in line or "Yes" in line: # Adapt based on actual output
                     info.hipaa_password_complexity = "Enabled"
                 elif "Disabled" in line or "No" in line:
                      info.hipaa_password_complexity = "Disabled"


            # Minimum password length
            elif "Minimum password length:" in line:
                try:
                    length_str = line.split(":")[1].strip()
                    if length_str.lower() == 'none':
                         info.hipaa_min_password_length = 0
                    else:
                         info.hipaa_min_password_length = int(length_str)
                except (IndexError, ValueError):
                    logger.warning(f"Could not parse min password length: {line}")
                    info.hipaa_min_password_length = -2 # Parse error

            # Lockout threshold
            elif "Lockout threshold:" in line:
                try:
                    threshold_str = line.split(":")[1].strip()
                    if threshold_str.lower() == 'never':
                        info.hipaa_account_lockout_threshold = 0
                    else:
                        # Attempt to parse as number, assuming it means attempts
                         threshold_val = int(re.sub(r'\D', '', threshold_str)) # Extract digits
                         info.hipaa_account_lockout_threshold = threshold_val

                except (IndexError, ValueError):
                    logger.warning(f"Could not parse lockout threshold: {line}")
                    info.hipaa_account_lockout_threshold = -2 # Parse error


            # Lockout duration
            elif "Lockout duration (minutes):" in line:
                 try:
                      duration_str = line.split(":")[1].strip()
                      if info.hipaa_account_lockout_threshold == 0:
                           info.hipaa_account_lockout_duration_min = 0 # Duration is irrelevant if threshold is Never
                      else:
                          info.hipaa_account_lockout_duration_min = int(duration_str)

                 except (IndexError, ValueError):
                      logger.warning(f"Could not parse lockout duration: {line}")
                      if info.hipaa_account_lockout_threshold != 0:
                         info.hipaa_account_lockout_duration_min = -2 # Parse error only if relevant
                      else:
                           info.hipaa_account_lockout_duration_min = 0


        # Post-processing check for complexity if not explicitly found
        if info.hipaa_password_complexity == "Check Not Run":
             # If complexity wasn't explicitly found, mark as unknown rather than assuming disabled
             # GPO settings are more reliable for complexity.
             info.hipaa_password_complexity = "Unknown (Check GPO)"
             logger.warning("Password complexity state couldn't be reliably determined via 'net accounts'. Check Group Policy.")


    except Exception as e:
        error_msg = f"Error parsing net accounts output: {e}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['AccountPoliciesParse'] = error_msg
        # Reset relevant fields
        info.hipaa_password_complexity = "Error Parsing"
        info.hipaa_min_password_length = -1
        info.hipaa_account_lockout_threshold = -1
        info.hipaa_account_lockout_duration_min = -1

def check_lock_timeout_settings(info: SystemInfo):
    """
    Checks display off timeout, inactivity lock policy, and password requirement
    using registry and powercfg.
    """
    logger.info("Checking display off timeout and lock settings (Registry/powercfg)...")
    info.hipaa_display_off_timeout_ac_min = -1 # Error default
    info.hipaa_display_off_timeout_dc_min = -1 # Error default
    info.hipaa_inactivity_lock_timeout_sec = -1 # Error default
    info.hipaa_require_password_on_wakeup = "Error" # Error default

    active_scheme_guid = None
    password_required = "Unknown" # Intermediate status
    inactivity_timeout_policy_sec = -1 # Default to not set/error

    # --- Priority 1: Check Policy Registry Keys ---
    # Check if lock screen is disabled by policy
    try:
        policy_key_path = r"SOFTWARE\Policies\Microsoft\Windows\Personalization"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, policy_key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            no_lock_screen_val, _ = winreg.QueryValueEx(key, "NoLockScreen")
            if int(no_lock_screen_val) == 1:
                logger.warning("Lock screen is explicitly DISABLED by HKLM Policy (NoLockScreen=1).")
                info.hipaa_require_password_on_wakeup = "Disabled by Policy"
                # We can still check display timeout, but lock won't happen
                password_required = "Disabled by Policy" # Set intermediate status
            else:
                 logger.info("HKLM Policy NoLockScreen is not set to 1 (or doesn't exist).")
    except FileNotFoundError:
        logger.info("HKLM Policy NoLockScreen key/value not found.")
    except Exception as e:
        logger.warning(f"Could not read HKLM Policy NoLockScreen: {e}")
        info.hipaa_checks_error_details['LockTimeoutPolicyCheck'] = f"Error checking NoLockScreen policy: {e}"

    # Check inactivity timeout policy (most direct setting for lock)
    try:
        policy_key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, policy_key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            inactivity_timeout_policy_sec = int(winreg.QueryValueEx(key, "InactivityTimeoutSecs")[0])
            info.hipaa_inactivity_lock_timeout_sec = inactivity_timeout_policy_sec
            logger.info(f"Found HKLM Policy InactivityTimeoutSecs: {inactivity_timeout_policy_sec} seconds.")
            # If this policy is set and non-zero, password is implicitly required
            if inactivity_timeout_policy_sec > 0 and password_required != "Disabled by Policy":
                password_required = "Yes"
                logger.info(" -> Inactivity policy implies password required on wakeup.")
            elif inactivity_timeout_policy_sec == 0:
                 logger.info(" -> Inactivity policy is set to 0 (disabled).")
                 # Don't assume password isn't required, other settings might apply
            # If < 0 (error reading), leave password_required as is
    except FileNotFoundError:
        logger.info("HKLM Policy InactivityTimeoutSecs key/value not found.")
        info.hipaa_inactivity_lock_timeout_sec = 0 # Treat as not set
    except ValueError:
         logger.error(f"Could not parse HKLM Policy InactivityTimeoutSecs value.")
         info.hipaa_checks_error_details['LockTimeoutPolicyParse'] = "Failed to parse InactivityTimeoutSecs"
         info.hipaa_inactivity_lock_timeout_sec = -2 # Indicate parse error
    except Exception as e:
        logger.warning(f"Could not read HKLM Policy InactivityTimeoutSecs: {e}")
        info.hipaa_checks_error_details['LockTimeoutPolicyCheck'] = f"Error checking InactivityTimeoutSecs policy: {e}"
        info.hipaa_inactivity_lock_timeout_sec = -1 # General error

    # --- Check Power Config Settings (Display Timeout and Console Lock) ---
    # Get active scheme first
    stdout_guid, stderr_guid, retcode_guid = run_command("powercfg /getactivescheme")
    if retcode_guid == 0 and stdout_guid:
        match = re.search(r"GUID: ([a-f0-9-]+)", stdout_guid, re.IGNORECASE)
        if match:
            active_scheme_guid = match.group(1)
            logger.info(f"Active power scheme GUID: {active_scheme_guid}")
        else:
            logger.error("Could not parse active power scheme GUID from powercfg output.")
            info.hipaa_checks_error_details['LockTimeout'] = "Failed to parse active scheme GUID"
            active_scheme_guid = None # Ensure it's None so subsequent queries fail gracefully
    else:
        logger.error(f"Failed to get active power scheme. RetCode: {retcode_guid}, Stderr: {stderr_guid}")
        info.hipaa_checks_error_details['LockTimeout'] = f"powercfg /getactivescheme failed (Code: {retcode_guid})"
        active_scheme_guid = None

    # Proceed only if we have an active scheme GUID
    if active_scheme_guid:
        # GUIDs for settings
        display_subgroup_guid = "7516b95f-f776-4464-8c53-06167f40cc99"
        display_off_setting_guid = "3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e"
        session_state_subgroup_guid = "7648efa3-dd9c-4e3e-b566-50f929386280"
        console_lock_timeout_guid = "8ec4b3a5-6868-48c2-be75-4f3044be88a7"

        # Query Display Off Timeout
        cmd_display = f"powercfg /query {active_scheme_guid} {display_subgroup_guid} {display_off_setting_guid}"
        stdout_disp, stderr_disp, retcode_disp = run_command(cmd_display)
        if retcode_disp == 0 and stdout_disp:
            ac_match = re.search(r"Current AC Power Setting Index:\s*0x([0-9a-f]+)", stdout_disp, re.IGNORECASE)
            dc_match = re.search(r"Current DC Power Setting Index:\s*0x([0-9a-f]+)", stdout_disp, re.IGNORECASE)
            if ac_match:
                try: ac_seconds = int(ac_match.group(1), 16); info.hipaa_display_off_timeout_ac_min = ac_seconds // 60 if ac_seconds > 0 else 0
                except ValueError: logger.error(f"Could not parse AC display timeout hex: {ac_match.group(1)}"); info.hipaa_checks_error_details['LockTimeoutACParse'] = "Failed to parse AC hex"
            else: logger.warning("Could not find AC display timeout setting."); info.hipaa_display_off_timeout_ac_min = -2
            if dc_match:
                try: dc_seconds = int(dc_match.group(1), 16); info.hipaa_display_off_timeout_dc_min = dc_seconds // 60 if dc_seconds > 0 else 0
                except ValueError: logger.error(f"Could not parse DC display timeout hex: {dc_match.group(1)}"); info.hipaa_checks_error_details['LockTimeoutDCParse'] = "Failed to parse DC hex"
            else: logger.warning("Could not find DC display timeout setting."); info.hipaa_display_off_timeout_dc_min = -2
        else:
            logger.error(f"Failed to query display off setting. RetCode: {retcode_disp}, Stderr: {stderr_disp}")
            info.hipaa_checks_error_details['LockTimeoutDisplayQuery'] = f"powercfg query display failed (Code: {retcode_disp})"

        # Query Console Lock Timeout (only if password requirement not already determined by policy)
        if password_required not in ["Yes", "Disabled by Policy"]:
            cmd_console_lock = f"powercfg /query {active_scheme_guid} {session_state_subgroup_guid} {console_lock_timeout_guid}"
            stdout_lock, stderr_lock, retcode_lock = run_command(cmd_console_lock)
            if retcode_lock == 0 and stdout_lock:
                ac_lock_match = re.search(r"Current AC Power Setting Index:\s*0x([0-9a-f]+)", stdout_lock, re.IGNORECASE)
                dc_lock_match = re.search(r"Current DC Power Setting Index:\s*0x([0-9a-f]+)", stdout_lock, re.IGNORECASE)
                lock_enabled = False
                parsed_lock_setting = False
                if ac_lock_match: parsed_lock_setting = True; lock_enabled = lock_enabled or (int(ac_lock_match.group(1), 16) != 0)
                if dc_lock_match: parsed_lock_setting = True; lock_enabled = lock_enabled or (int(dc_lock_match.group(1), 16) != 0)

                if parsed_lock_setting:
                    password_required = "Yes" if lock_enabled else "No"
                    logger.info(f"Console lock display off timeout enabled via powercfg: {lock_enabled} -> Password Required: {password_required}")
                else:
                    logger.warning("Could not parse Console lock display off timeout setting. Falling back to registry.")
                    password_required = check_screensaver_secure_registry() # Fallback 1
            else:
                error_detail = f"powercfg query console lock failed (Code: {retcode_lock})"
                if "does not exist" in (stderr_lock or ""): error_detail += " - Setting/Subgroup GUID likely invalid."
                logger.warning(f"Failed to query console lock setting ({error_detail}). Falling back to registry.")
                info.hipaa_checks_error_details['LockTimeoutConsoleLockQuery'] = error_detail
                password_required = check_screensaver_secure_registry() # Fallback 2
    else:
        # If getting active scheme failed, we can't check powercfg settings
        logger.warning("Cannot check powercfg timeouts because active scheme query failed.")
        # Rely on registry checks already performed or fallback
        if password_required not in ["Yes", "Disabled by Policy"]:
             password_required = check_screensaver_secure_registry() # Fallback 3

    # --- Final Fallback (if still Unknown/Error) ---
    if password_required in ["Unknown", "Error"]:
         logger.warning(f"Password requirement still '{password_required}' after checks, performing final registry fallback.")
         password_required = check_screensaver_secure_registry()

    # Final assignment for password requirement
    info.hipaa_require_password_on_wakeup = password_required
    logger.info(f"Final determination for Password Required on Wakeup: {info.hipaa_require_password_on_wakeup}")

def check_screensaver_secure_registry():
    """
    Fallback check for password requirement using ScreenSaverIsSecure registry keys.
    Returns: "Yes", "No", "Unknown"
    """
    logger.info("Performing fallback check via ScreenSaverIsSecure registry...")
    is_secure = "Unknown" # Default if not found or error
    gpo_secure_flag = -1
    user_secure_flag = -1

    # Check HKLM Policy first
    try:
        policy_key_path = r"SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, policy_key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            policy_secure_str, _ = winreg.QueryValueEx(key, "ScreenSaverIsSecure")
            gpo_secure_flag = int(policy_secure_str)
            logger.info(f"Found HKLM Policy ScreenSaverIsSecure: {gpo_secure_flag}")
    except FileNotFoundError: logger.info("No HKLM Policy ScreenSaverIsSecure found.")
    except Exception as e: logger.warning(f"Could not read HKLM Policy ScreenSaverIsSecure: {e}")

    # Check HKCU if HKLM doesn't dictate
    if gpo_secure_flag == -1: # If GPO didn't set it
        try:
            key_path = r"Control Panel\Desktop"
            # Use try-except for OpenKey in case user profile isn't loaded or accessible
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ) as key:
                    user_secure_str, _ = winreg.QueryValueEx(key, "ScreenSaverIsSecure")
                    user_secure_flag = int(user_secure_str)
                    logger.info(f"Found HKCU ScreenSaverIsSecure: {user_secure_flag}")
            except OSError as e_os:
                 logger.warning(f"Could not open HKCU key '{key_path}': {e_os}. Assuming default.")
                 user_secure_flag = -1 # Treat as not found if HKCU inaccessible
        except FileNotFoundError: logger.info("No HKCU ScreenSaverIsSecure value found.")
        except Exception as e: logger.warning(f"Could not read HKCU ScreenSaverIsSecure: {e}")

    # Determine final state
    if gpo_secure_flag != -1: is_secure = "Yes" if gpo_secure_flag == 1 else "No"
    elif user_secure_flag != -1: is_secure = "Yes" if user_secure_flag == 1 else "No"
    # else: is_secure remains "Unknown"

    logger.info(f"Registry fallback check determined Password Required: {is_secure}")
    return is_secure

def check_antivirus_status_cim(info: SystemInfo):
    """
    Checks installed Antivirus products and their states using PowerShell Get-CimInstance.
    Stores the results as a formatted string.
    """
    logger.info("Checking Antivirus products (PowerShell Get-CimInstance)...")
    # Command to get AV products and convert to JSON
    command = 'powershell -NoProfile -Command "Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct -ErrorAction SilentlyContinue | Select-Object -Property displayName, productState | ConvertTo-Json -Compress"'
    stdout, stderr, retcode = run_command(command)

    info.hipaa_antivirus_products_details = "Error" # Default

    if retcode != 0 or stdout is None or not stdout.strip():
        error_msg = f"Get-CimInstance AntivirusProduct command returned no output or failed. Stderr: {stderr or 'No stderr'}. RetCode: {retcode}."
        if stderr and ("Get-CimInstance" in stderr or "Invalid namespace" in stderr or "Invalid class" in stderr):
             error_msg = f"Get-CimInstance failed: {stderr.strip()}"
        elif not stdout or stdout.strip() == "[]" or stdout.strip() == "":
             # Command ran but found no products - this is not an error state for this check
             error_msg = "No Antivirus products found via Get-CimInstance in root/SecurityCenter2."
             logger.info(error_msg)
             info.hipaa_antivirus_products_details = "None Found"
             return # Successfully determined none found

        # Only log as error if command failed or returned unexpected stderr
        logger.error(error_msg)
        info.hipaa_checks_error_details['AntivirusCIM'] = error_msg
        return

    try:
        output_text = stdout.strip()
        # Handle empty JSON array output
        if not output_text or output_text == "[]":
             logger.info("Get-CimInstance reported no AntiVirusProduct (empty JSON).")
             info.hipaa_antivirus_products_details = "None Found"
             return

        # Handle single JSON object output
        if not output_text.startswith('['):
             output_text = f"[{output_text}]"

        av_products = json.loads(output_text)

        if not av_products: # Double check after parsing
            info.hipaa_antivirus_products_details = "None Found"
            logger.info("Parsed JSON but no AntiVirusProduct objects found.")
            return

        product_details_list = []
        for product in av_products:
            display_name = product.get('displayName', 'Unknown Provider')
            product_state_raw = product.get('productState')
            product_state = 0 # Default state if parsing fails
            state_hex = "N/A"
            try:
                if isinstance(product_state_raw, (int, float)):
                    product_state = int(product_state_raw)
                    state_hex = f"0x{product_state:X}" # Format as hex
                elif isinstance(product_state_raw, str) and product_state_raw.isdigit():
                    product_state = int(product_state_raw)
                    state_hex = f"0x{product_state:X}" # Format as hex
                else:
                     logger.warning(f"Could not parse productState '{product_state_raw}' for {display_name} as integer.")
                     product_state = -1 # Indicate parse error state
                     state_hex = "Parse Error"
            except Exception as parse_e:
                 logger.warning(f"Error formatting productState '{product_state_raw}' for {display_name}: {parse_e}")
                 product_state = -2 # Indicate format error state
                 state_hex = "Format Error"


            detail_string = f"{display_name}, State: {product_state} ({state_hex})"
            product_details_list.append(detail_string)
            logger.info(f"Found AV: {detail_string}")

        # Join the details into a single string
        info.hipaa_antivirus_products_details = ", ".join(product_details_list)

    except json.JSONDecodeError as e:
        error_msg = f"Error parsing Get-CimInstance JSON output: {e}. Output: {stdout[:500]}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['AntivirusCIMParse'] = error_msg
        info.hipaa_antivirus_products_details = "Error Parsing Output"
    except Exception as e:
        error_msg = f"General error processing Get-CimInstance output: {e}"
        logger.exception("Error processing AV CIM output") # Log full traceback
        info.hipaa_checks_error_details['AntivirusCIMProcess'] = error_msg
        info.hipaa_antivirus_products_details = "Error Processing"

def check_firewall_status(info: SystemInfo):
    """Checks Windows Firewall status for Domain, Private, and Public profiles using netsh."""
    update_status("Checking Firewall status (netsh)...")
    stdout, stderr, retcode = run_command("netsh advfirewall show allprofiles")

    if retcode != 0 or stdout is None:
        error_msg = f"netsh advfirewall failed. Return code: {retcode}. Stderr: {stderr}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['Firewall'] = error_msg
        info.hipaa_firewall_status = "Error"
        return

    try:
        # Check if 'State' is 'ON' for all relevant profiles found
        profiles_on = {}
        current_profile = None
        profile_pattern = re.compile(r"^(Domain|Private|Public) Profile Settings:$")
        state_pattern = re.compile(r"^State\s+(ON|OFF)", re.IGNORECASE) # Case insensitive state

        for line in stdout.strip().split('\n'):
            line = line.strip()
            profile_match = profile_pattern.match(line)
            if profile_match:
                current_profile = profile_match.group(1)
                profiles_on[current_profile] = False # Assume off until State ON found
                continue

            if current_profile:
                state_match = state_pattern.match(line)
                if state_match:
                    if state_match.group(1).upper() == 'ON':
                        profiles_on[current_profile] = True
                    # Reset current_profile once state is found for it? Optional.

        if not profiles_on: # No profiles parsed
             info.hipaa_firewall_status = "Error Parsing Output"
             logger.error("Could not parse any firewall profiles from netsh output.")
             return

        # Determine overall status
        all_on = True
        any_on = False
        for profile, status in profiles_on.items():
            if status:
                any_on = True
            else:
                all_on = False

        if all_on:
            info.hipaa_firewall_status = "Enabled (All Profiles)"
        elif any_on:
            info.hipaa_firewall_status = "Partially Enabled"
        else:
            info.hipaa_firewall_status = "Disabled"

    except Exception as e:
        error_msg = f"Error parsing netsh firewall output: {e}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['FirewallParse'] = error_msg
        info.hipaa_firewall_status = "Error Parsing Output"

def check_usb_storage_restriction(info: SystemInfo):
    """Checks if USB Mass Storage driver (USBSTOR) start type is disabled via registry."""
    update_status("Checking USB Storage restriction (Registry)...")
    key_path = r"SYSTEM\CurrentControlSet\Services\USBSTOR"
    value_name = "Start"
    restricted_value = 4 # 4 means disabled

    try:
        # Requires HKLM access, needs admin privileges
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            start_value, _ = winreg.QueryValueEx(key, value_name)
            if int(start_value) == restricted_value:
                info.hipaa_usb_storage_restricted = "Restricted"
            else:
                info.hipaa_usb_storage_restricted = "Allowed"

    except FileNotFoundError:
        # If the USBSTOR key or Start value doesn't exist, it's likely not explicitly restricted (default is enabled)
        info.hipaa_usb_storage_restricted = "Allowed (Default)"
        update_status(f"USBSTOR registry key or '{value_name}' not found. Assuming allowed.")
    except PermissionError:
        error_msg = "Permission denied accessing HKLM registry for USBSTOR. Run as Administrator."
        logger.error(error_msg)
        info.hipaa_checks_error_details['USBStorage'] = error_msg
        info.hipaa_usb_storage_restricted = "Error (Permission Denied)"
    except Exception as e:
        error_msg = f"Error reading registry for USBSTOR: {e}"
        logger.error(error_msg)
        info.hipaa_checks_error_details['USBStorage'] = error_msg
        info.hipaa_usb_storage_restricted = "Error"

def check_domain_membership(info: SystemInfo):
    """Checks if the computer is joined to a domain using PowerShell."""
    logging.info("Checking domain membership (PowerShell Get-CimInstance)...")
    update_status("Checking domain membership (PowerShell Get-CimInstance)...")
    # Use Get-CimInstance for modern approach, more robust than gwmi alias
    command = 'powershell -NoProfile -Command "(Get-CimInstance -ClassName Win32_ComputerSystem).PartOfDomain"'
    stdout, stderr, retcode = run_command(command)

    info.domain_joined = "Error" # Default

    if retcode != 0:
        error_msg = f"Get-CimInstance Win32_ComputerSystem command failed. RetCode: {retcode}. Stderr: {stderr or 'No stderr'}."
        # Check specific PowerShell errors if possible
        if stderr and ("Get-CimInstance" in stderr or "Invalid class" in stderr):
             error_msg = f"Get-CimInstance failed: {stderr.strip()}"
        logging.error(error_msg)
        info.hipaa_checks_error_details['DomainCheck'] = error_msg
        return

    if stdout is not None:
        result_str = stdout.strip().lower()
        if result_str == 'true':
            info.domain_joined = True
            logging.info("Computer is joined to a domain.")
        elif result_str == 'false':
            info.domain_joined = False
            logging.info("Computer is NOT joined to a domain.")
        else:
            logging.warning(f"Unexpected output from domain check command: {stdout.strip()}")
            info.hipaa_checks_error_details['DomainCheckParse'] = f"Unexpected output: {stdout.strip()}"
            info.domain_joined = "Error (Parse)"
    else:
        # Should not happen if retcode is 0, but handle defensively
        logging.error("Domain check command succeeded but returned no output.")
        info.hipaa_checks_error_details['DomainCheck'] = "Command succeeded but no output"
        info.domain_joined = "Error (No Output)"


# --- Data Collection Wrappers ---
def run_hipaa_technical_checks(info: SystemInfo):
    """Runs all automatable HIPAA-related technical checks."""
    update_status("Starting HIPAA technical checks...")

    # Ensure platform is Windows before running Windows-specific commands
    if platform.system() != "Windows":
        logger.error("HIPAA technical checks are designed for Windows only.")
        info.hipaa_checks_error_details['Platform'] = "Not Windows"
        # Set all checks to an appropriate status like 'Not Applicable' or 'Error'
        for attr in dir(info):
             if attr.startswith("hipaa_"):
                 # Check if it's one of the main status fields (not details dict or error dict)
                  if isinstance(getattr(info, attr), (str, int, float)) and 'details' not in attr.lower() and 'error' not in attr.lower():
                      setattr(info, attr, "Not Applicable (Non-Windows)")
        return

    # Run checks - consider adding checks for required privileges here if needed
    # (e.g., using ctypes.windll.shell32.IsUserAnAdmin())

    check_bitlocker_status(info)
    check_audit_policy(info)
    check_security_log_settings(info)
    check_account_policies(info)
    check_lock_timeout_settings(info)
    # Use PowerShell AV check first, WMI is fallback inside the PS function
    check_antivirus_status_cim(info)
    check_firewall_status(info)
    check_usb_storage_restriction(info)
    # Add future checks here (e.g., SMB Encryption check if feasible/desired)

    update_status("Finished HIPAA technical checks.")

def collect_system_data(assessment_id: str, hipaa_compliant: bool) -> SystemInfo:
    """ Collects system information and returns a populated SystemInfo object. """
    info = SystemInfo()
    info.assessment_id = assessment_id
    info.hipaa_compliant = hipaa_compliant

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
            logger.error(f"CPU detail collection failed: {e}")
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

        # --- Check if system is joined to domain ---
        check_domain_membership(info)

        # --- Windows Update Check ---
        run_hipaa_technical_checks(info)

        update_status("Data collection complete.")

    except Exception as e:
        logger.error(f"CRITICAL ERROR during data collection: {e}")
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
        logger.info(f"Data saved to {filename}")
    except Exception as e:
        update_status(f"Error saving data locally: {e}")
        logger.error(f"Failed to save CSV: {e}")


# --- Background Task ---
def worker_thread_task(assessment_id: str, hipaa_compliant: bool):
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
                logger.error(f"Failed to CoInitialize COM for thread: {com_init_e}")
                update_status("Error initializing COM for Update check.")

        system_info = collect_system_data(assessment_id, hipaa_compliant)

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
        logger.error(f"Critical error in worker thread: {e}")
        update_status(f"Critical error: {e}")
        show_final_message("ERROR", f"A critical error occurred: {e}\nPlease call us at {SUPPORT_PHONE_NUMBER}.")
        if system_info: save_data_to_csv(system_info)
    finally:
        # --- Uninitialize COM for this thread ---
        if coinitialized and pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception as com_uninit_e:
                logger.error(f"Failed to CoUninitialize COM for thread: {com_uninit_e}")
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
            logger.info(f"Successfully set window icon from {icon_path}")
        except Exception as e:
            logger.error(f"Error setting window icon using iconbitmap: {e}")
            # Fallback attempt using iconphoto with PNG (less reliable)
            try:
                logger.info("Attempting fallback icon using iconphoto with PNG...")
                png_icon_path = get_resource_path(LOGO_FILENAME)
                pil_icon = Image.open(png_icon_path)
                self.app_icon = ImageTk.PhotoImage(pil_icon) # Need to store PhotoImage
                self.iconphoto(True, self.app_icon)
                logger.info(f"Successfully set window icon using iconphoto fallback.")
            except Exception as e_photo:
                 logger.error(f"Error setting window icon using iconphoto fallback: {e_photo}")

        self.geometry("450x200")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.logo_image = None
        self.status_label = None
        self.worker_thread = None
        self.assessment_id = "Not Provided"
        self.hipaa_compliant = False
        self.hipaa_prompt_window = None

        self.show_splash_screen()

    def _center_window(self, window, width, height):
        """Helper to center a top-level window."""
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x_co = int((screen_width / 2) - (width / 2))
        y_co = int((screen_height / 2) - (height / 2))
        window.geometry(f"{width}x{height}+{x_co}+{y_co}")
    
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
            logger.error(f"Error: Logo file '{LOGO_FILENAME}' not found at '{get_resource_path(LOGO_FILENAME)}'.")
            error_label = ctk.CTkLabel(self.splash, text=f"Error: Logo not found!\nPlace {LOGO_FILENAME}\n in the application folder.", text_color="red")
            error_label.pack(pady=20, padx=20, expand=True, fill="both")
        except Exception as e:
             logger.error(f"Error loading logo: {e}")
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
        self._center_window(dialog, dialog_width, dialog_height)

        entered_id = dialog.get_input()

        if entered_id:
            self.assessment_id = entered_id
            logger.info(f"User entered ID: {self.assessment_id}")
        else:
            self.assessment_id = "Not Provided" # Handle cancel or empty input
            logger.warning("User cancelled or provided no ID.")

        # Prompt the user to see if HIPAA applies
        self.prompt_for_hipaa()

    def prompt_for_hipaa(self):
        """ Shows a modal Yes/No dialog to check for HIPAA compliance """
        if self.hipaa_prompt_window is not None: # Prevent multiple prompts
             self.hipaa_prompt_window.focus()
             return

        self.hipaa_prompt_window = ctk.CTkToplevel(self)
        dialog = self.hipaa_prompt_window # Use the instance variable for clarity
        dialog.title("HIPAA Confirmation")
        prompt_w, prompt_h = 350, 150
        dialog.geometry(f"{prompt_w}x{prompt_h}")
        dialog.resizable(False, False)
        dialog.grab_set()  # Make modal
        dialog.attributes("-topmost", True)
        self._center_window(dialog, prompt_w, prompt_h)

        # Make closing the window equivalent to clicking "No"
        # Note: We use a lambda here to ensure the correct dialog reference is passed
        dialog.protocol("WM_DELETE_WINDOW", lambda: self._handle_hipaa_response(False, dialog))

        question = "Is this business regulated by HIPAA?"
        label = ctk.CTkLabel(dialog, text=question, wraplength=prompt_w - 40)
        label.pack(pady=20, padx=20, expand=True, fill="x")

        button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        button_frame.pack(pady=(0, 15), padx=20, fill="x")
        button_frame.columnconfigure((0, 1), weight=1) # Center buttons

        # Assign separate class methods as commands
        # Use lambda to pass the response value and dialog reference easily
        yes_button = ctk.CTkButton(button_frame, text="Yes", width=100,
                                   command=lambda: self._handle_hipaa_response(True, dialog))
        yes_button.grid(row=0, column=0, padx=(0, 5))

        no_button = ctk.CTkButton(button_frame, text="No", width=100,
                                  command=lambda: self._handle_hipaa_response(False, dialog),
                                  fg_color="#D32F2F", hover_color="#B71C1C")
        no_button.grid(row=0, column=1, padx=(5, 0))

    # --- New Handler Method (Refactored) ---
    def _handle_hipaa_response(self, is_hipaa: bool, dialog_window: ctk.CTkToplevel):
        """ Handles the response from the HIPAA prompt and proceeds. """
        self.hipaa_compliant = is_hipaa
        logger.info(f"HIPAA regulated: {self.hipaa_compliant}")

        if dialog_window:
            dialog_window.destroy()
        self.hipaa_prompt_window = None # Clear the reference

        # Now proceed to the main application window and checks
        self.deiconify() # Show main window
        self.lift() # Bring main window to front
        self.setup_main_window()
        self.start_checks() # Pass the collected ID

    def setup_main_window(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(self, text="Initializing...", wraplength=400, justify="center")
        self.status_label.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        self.process_gui_queue()

    def start_checks(self):
        """ Starts the background check thread, passing the user ID """
        self.status_label.configure(text=f"Starting readiness checks for ID: {self.assessment_id}...")
        # Pass assessment_id to the worker thread
        self.worker_thread = threading.Thread(target=worker_thread_task, args=(self.assessment_id,self.hipaa_compliant,), daemon=True)
        self.worker_thread.start()

    def process_gui_queue(self):
        # (This function remains the same as before)
        try:
            message = gui_queue.get_nowait()
            logger.info(f"GUI Queue Received: {message}")

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
        logger.info("Closing application...")
        self.destroy()

# --- Main Execution ---
if __name__ == "__main__":
    # Removed CoInitialize from main thread - it's now handled per-thread
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    app = App()
    app.mainloop()
