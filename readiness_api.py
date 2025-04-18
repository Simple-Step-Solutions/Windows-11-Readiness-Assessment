def assess_windows11_readiness(system_info):
    """
    Determines the Windows 11 readiness level of a system based on its specifications.

    Args:
        system_info (dict): A dictionary containing system information.

    Returns:
        str: The readiness level ("Does not meet requirements", "Required", or "Recommended").
    """

    os_version = system_info.get("os_version", "")
    ram_total_gb = system_info.get("ram_total_gb", 0)
    disk_total_gb = system_info.get("disk_total_gb", 0)
    secure_boot_enabled = system_info.get("secure_boot_enabled", False)
    tpm_present_str = str(system_info.get("tpm_present", False)).lower()
    processor = system_info.get("processor", "").lower()
    architecture = system_info.get("architecture", "").lower()
    graphics_card = system_info.get("graphics_card", "").lower()
    wddm_version_str = system_info.get("wddm_version", "")

    meets_windows_min = True
    reasons_not_meeting_windows = []

    # Check OS Version (Windows 11 requires at least version 10.0.22000)
    major, minor, build = map(int, os_version.split('.')) if "." in os_version and os_version.split('.')[0] == '10' else (0, 0, 0)
    if major < 10 or (major == 10 and build < 19041):
        meets_windows_min = False
        reasons_not_meeting_windows.append(f"OS Version ({os_version}) is below Windows 11 minimum.")
    elif major > 10:
        # Assuming all future major versions meet the requirement for this check
        pass

    # Check RAM (Windows 11 requires 4 GB)
    if ram_total_gb < 4:
        meets_windows_min = False
        reasons_not_meeting_windows.append(f"RAM ({ram_total_gb:.2f} GB) is below Windows 11 minimum (4 GB).")

    # Check Disk Space (Windows 11 requires 64 GB)
    if disk_total_gb < 64:
        meets_windows_min = False
        reasons_not_meeting_windows.append(f"Total Disk Space ({disk_total_gb:.2f} GB) is below Windows 11 minimum (64 GB).")

    # Check Secure Boot
    if not secure_boot_enabled:
        meets_windows_min = False
        reasons_not_meeting_windows.append("Secure Boot is not enabled.")

    # Check TPM (Windows 11 requires TPM 2.0)
    if "failed" in tpm_present_str or "not present" in tpm_present_str:
        meets_windows_min = False
        reasons_not_meeting_windows.append("TPM 2.0 is not present or could not be verified.")
    elif "1.2" in system_info.get("tpm_version", ""):
        meets_windows_min = False
        reasons_not_meeting_windows.append(f"TPM version ({system_info.get('tpm_version')}) is below Windows 11 minimum (2.0).")

    # Check Processor (This is complex and requires a list. For now, just check architecture)
    if "64" not in architecture:
        meets_windows_min = False
        reasons_not_meeting_windows.append(f"Architecture ({architecture}) is not 64-bit.")
    # In a real scenario, you'd have a list of supported Intel and AMD processors.

    # Check Graphics Card and WDDM (Windows 11 requires WDDM 2.0)
    if "basic display adapter" in graphics_card:
        meets_windows_min = False
        reasons_not_meeting_windows.append(f"Graphics card ({graphics_card}) may not meet Windows 11 requirements.")
    if wddm_version_str:
        try:
            wddm_major = float(wddm_version_str.split(':')[1].strip().split('.')[0])
            if wddm_major < 2.0:
                meets_windows_min = False
                reasons_not_meeting_windows.append(f"WDDM version ({wddm_version_str}) is below Windows 11 minimum (2.0).")
        except (IndexError, ValueError):
            pass # Unable to parse WDDM version, assume it might be an issue

    if not meets_windows_min:
        return "Does not meet requirements", "\n".join(reasons_not_meeting_windows)
    else:
        # Meets Windows 11 minimum requirements, now check our "recommended" specs
        meets_our_recommended = True
        reasons_not_meeting_recommended = []

        # Our Recommended RAM (e.g., 8 GB)
        if ram_total_gb < 8:
            meets_our_recommended = False
            reasons_not_meeting_recommended.append(f"RAM ({ram_total_gb:.2f} GB) is below our recommended (8 GB).")

        # Our Recommended Disk Space (e.g., 256 GB SSD - let's just check total for simplicity)
        if disk_total_gb < 256:
            meets_our_recommended = False
            reasons_not_meeting_recommended.append(f"Total Disk Space ({disk_total_gb:.2f} GB) is below our recommended (256 GB).")

        # Our Recommended Processor (e.g., specific Intel i5 or better - again, complex, just a placeholder)
        if "intel" in processor and "family 6 model 63" in processor: # Example: Older Intel
            meets_our_recommended = False
            reasons_not_meeting_recommended.append(f"Processor ({processor}) is below our recommended level.")

        # Our Recommended Graphics (e.g., dedicated GPU with certain VRAM - placeholder)
        if "basic display adapter" in graphics_card:
            meets_our_recommended = False
            reasons_not_meeting_recommended.append(f"Graphics card ({graphics_card}) is below our recommended level.")

        if meets_our_recommended:
            return "Recommended", "\n".join(reasons_not_meeting_windows)
        else:
            return "Required", "\n".join(reasons_not_meeting_windows)


if __name__ == "__main__":
    sample_data = {
    "assessment_id": "Not Provided",
    "hostname": "HQ-Office-01",
    "os_platform": "Windows",
    "os_version": "10.0.19045",
    "os_release": "10",
    "architecture": "AMD64",
    "processor": "AMD64 Family 25 Model 97 Stepping 2, AuthenticAMD",
    "cpu_physical_cores": 6,
    "cpu_logical_cores": 12,
    "cpu_max_speed_ghz": 4.7,
    "timestamp_utc": "2025-04-18 18:00:35",
    "timestamp_local": "2025-04-18 14:00:35",
    "timezone_name": "Eastern Daylight Time",
    "timezone_offset_utc": "-04:00",
    "ram_total_gb": 31.15,
    "ram_speed": "5600MHz",
    "ram_type": "DDR5",
    "disk_total_gb": 464.66,
    "disk_free_gb": 98.92,
    "system_drive_type": "SSD",
    "tpm_present": True,
    "tpm_version": "2.0, 0, 1.59",
    "tpm_enabled": True,
    "secure_boot_enabled": False,
    "graphics_card": "Intel(R) Arc(TM) A380 Graphics",
    "wddm_version": "Driver: 32.0.101.6651",
    "pending_updates_count": 1,
    "collection_error": None,
    "wmi_error_details": "Secure Boot WMI Query Failed: AttributeError: winmgmts:.Win32_SecureBoot",
    "update_check_error_details": None,
    "os_version_name": "Windows 10 22H2"
    }

    print(assess_windows11_readiness(sample_data))