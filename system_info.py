# --- Data Class ---
class SystemInfo:
    """ Holds all collected system information with default values. """
    def __init__(self):
        # User Input
        self.assessment_id = "Not Provided"
        # Basic Info
        self.hostname = "Undetermined"
        self.manufacturer = "Undetermined"
        self.os_platform = "Undetermined"
        self.os_version = "Undetermined"
        self.os_release = "Undetermined"
        self.architecture = "Undetermined"
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
        # Status/Error fields
        self.collection_error = None
        self.wmi_error_details = None
        self.update_check_error_details = None

    def to_dict(self):
        """ Convert the object's attributes to a dictionary for serialization. """
        return self.__dict__
