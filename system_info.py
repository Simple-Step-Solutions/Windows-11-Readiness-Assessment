# --- Data Class ---
class SystemInfo:
    """ Holds all collected system information with default values. """
    def __init__(self):
        # User Input
        self.assessment_id = "Not Provided"
        self.hipaa_compliant = False
        
        # Basic Info
        self.hostname = "Undetermined"
        self.manufacturer = "Undetermined"
        self.os_platform = "Undetermined"
        self.os_version = "Undetermined"
        self.os_release = "Undetermined"
        self.architecture = "Undetermined"
        self.domain_joined = "Undetermined"
        
        # CPU Info
        self.processor = "Undetermined"
        self.cpu_physical_cores = 0
        self.cpu_logical_cores = 0
        self.cpu_max_speed_ghz = 0.0
        
        # Timestamps
        self.timestamp_utc = "Undetermined"
        self.timestamp_local = "Undetermined"
        self.timezone_name = "Undetermined"
        self.timezone_offset_utc = "Undetermined"
        
        # Hardware Info
        self.ram_total_gb = 0.0
        self.ram_speed_mhz = 0
        self.ram_type = "Undetermined"
        self.disk_total_gb = 0.0
        self.disk_free_gb = 0.0
        self.system_drive_type = "Undetermined"
        
        # WMI Dependent Info
        self.tpm_present = "Check Not Run"
        self.tpm_version = "Check Not Run"
        self.tpm_enabled = "Check Not Run"
        self.secure_boot_enabled = "Check Not Run"
        self.graphics_card = "Check Not Run"
        self.wddm_version = "Check Not Run"
        
        # Windows Update Info
        self.pending_updates_count = -1 # -1: Error, -2: Module Missing, -3: COM Error, -4: Service Error
        
        # Encryption at Rest (HIPAA §164.312(a)(2)(iv))
        self.hipaa_bitlocker_status = "Check Not Run" # Overall status: "Encrypted", "Not Encrypted", "Mixed", "Error"
        self.hipaa_bitlocker_details = {} # Dict: {DriveLetter: {"ProtectionStatus": "On/Off", "EncryptionMethod": "XTS-AES 128", "EncryptionPercentage": 100}}

        # Audit Controls (HIPAA §164.312(b))
        self.hipaa_audit_logon_events = "Check Not Run" # Value: "Success and Failure", "Success", "Failure", "No Auditing", "Error"
        self.hipaa_audit_account_mgmt = "Check Not Run" # Value: "Success and Failure", "Success", "Failure", "No Auditing", "Error"
        self.hipaa_audit_policy_change = "Check Not Run" # Value: "Success and Failure", "Success", "Failure", "No Auditing", "Error"
        self.hipaa_audit_object_access = "Check Not Run" # Value: "Success and Failure", "Success", "Failure", "No Auditing", "Error" (Note: Requires specific SACLs on objects)
        self.hipaa_audit_log_max_size_mb = -1 # Value: Size in MB or -1 if error
        self.hipaa_audit_log_retention = "Check Not Run" # Value: "Overwrite as needed", "Archive the log when full", "Do not overwrite events", "Error"

        # Access Control - Technical Aspects (HIPAA §164.312(a)(1), §164.308(a)(5)(ii)(D))
        self.hipaa_password_complexity = "Check Not Run" # Value: "Enabled", "Disabled", "Error"
        self.hipaa_min_password_length = -1 # Value: Length or -1 if error
        self.hipaa_account_lockout_threshold = -1 # Value: Attempts (0=Off) or -1 if error
        self.hipaa_account_lockout_duration_min = -1 # Value: Minutes or -1 if error
        self.hipaa_display_off_timeout_ac_min = -1 # Minutes, -1 Error, 0 Never
        self.hipaa_display_off_timeout_dc_min = -1 # Minutes, -1 Error, 0 Never
        self.hipaa_inactivity_lock_timeout_sec = -1 # Seconds from Policy, -1 Error, 0 Disabled/Not Set
        self.hipaa_require_password_on_wakeup = "Check Not Run" # Values: "Yes", "No", "Error", "Unknown", "Disabled by Policy"

        # Integrity - Technical Aspects (HIPAA §164.312(c)(1))
        self.hipaa_antivirus_products_details = "Check Not Run"

        # Transmission Security / Workstation Security - Technical Aspects (HIPAA §164.312(e)(1), §164.310(c))
        self.hipaa_firewall_status = "Check Not Run" # Value: "Enabled", "Disabled", "Error" (Checks Domain, Private, Public profiles)
        # self.hipaa_smb_encryption_required = "Check Not Run" # Value: "Enabled", "Disabled", "Error" # Requires PowerShell 5.0+ typically
        self.hipaa_usb_storage_restricted = "Check Not Run" # Value: "Restricted", "Allowed", "Check Not Run", "Error"

        # Status/Error fields
        self.collection_error = None
        self.wmi_error_details = None
        self.update_check_error_details = None
        self.hipaa_checks_error_details = {} # Dict: {CheckName: ErrorMessage}


    def to_dict(self):
        """ Convert the object's attributes to a dictionary for serialization. """
        # Ensure all attributes are serializable
        d = {}
        for k, v in self.__dict__.items():
             if isinstance(v, (str, int, float, bool, list, dict, type(None))):
                 d[k] = v
             else:
                 d[k] = str(v) # Convert non-standard types to string
        return d
