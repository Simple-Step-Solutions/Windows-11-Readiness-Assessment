# --- Data Class ---
class SystemInfo:
    """ Holds all collected system information with default values. """
    def __init__(self):
        # User Input
        self.assessment_id = "Not Provided"
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
        self.pending_updates_count = -1 # -1: Error, -2: Module Missing, -3: COM Error, -4: WMI Service Error
        # Status/Error fields
        self.collection_error = None
        self.wmi_error_details = None # Added for specific WMI errors
        self.update_check_error_details = None # Added for specific Update check errors

    def to_dict(self):
        """ Convert the object's attributes to a dictionary for serialization. """
        # Filter out None values if desired, otherwise return full dict
        # return {k: v for k, v in self.__dict__.items() if v is not None}
        return self.__dict__
