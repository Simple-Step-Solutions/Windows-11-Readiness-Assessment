<#
.SYNOPSIS
    Checks Windows 11 readiness requirements and sends data to an API.
.DESCRIPTION
    This script gathers system information (OS, CPU, RAM, Disk, TPM, Secure Boot, Graphics, Pending Updates),
    prompts for an Assessment ID, and attempts to POST the collected data as JSON to a specified API endpoint.
    Designed as a fallback for the GUI application. Requires PowerShell 5.1 or later.
    Some checks (TPM, Secure Boot, Updates) may require administrative privileges for accurate results.
.NOTES
    Requires: PowerShell 5.1+
    Run as Administrator for best results.
#>
param() # No parameters needed for direct execution

# --- Configuration ---
$Script:ApiEndpointUrl = "https://n8n.simplestep.tech/webhook/aecfd3e4-659b-4e2c-a87e-3b13b34267a0"
$Script:SupportPhoneNumber = "1-914-250-9190"
# --- End Configuration ---

# --- Helper Function for WMI/CIM Queries with Error Handling ---
function Get-SafeCimInstance {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ClassName,
        [string]$Namespace,
        [string]$Filter,
        [switch]$FirstOrDefault
    )
    $params = @{ ClassName = $ClassName }
    if ($PSBoundParameters.ContainsKey('Namespace')) { $params.Namespace = $Namespace }
    if ($PSBoundParameters.ContainsKey('Filter')) { $params.Filter = $Filter }

    try {
        # Use -ErrorVariable to capture non-terminating errors if needed, but Stop is generally better here
        $result = Get-CimInstance @params -ErrorAction Stop
        if ($FirstOrDefault -and $result) {
            return $result[0]
        } else {
            return $result
        }
    } catch {
        # Log the specific error to the script-level error list
        $errMsg = "Failed to query CIM/WMI Class '$ClassName': $($_.Exception.Message)"
        Write-Warning $errMsg
        $script:WmiErrors.Add($errMsg) # Add error to the list
        return $null # Return null on failure
    }
}

# --- Main Script Logic ---
Write-Host "Starting Windows 11 Readiness Check..." -ForegroundColor Cyan

# 1. Prompt for Assessment ID (Renamed)
$assessmentId = Read-Host -Prompt "Please enter an Assessment ID"
if ([string]::IsNullOrWhiteSpace($assessmentId)) {
    $assessmentId = "Not Provided"
    Write-Warning "No Assessment ID provided."
} else {
    Write-Host "Using Assessment ID: $assessmentId"
}

# 2. Initialize Data Hashtable (Added manufacturer, ram_speed_mhz; Renamed user_provided_id)
$data = [ordered]@{
    assessment_id           = $assessmentId # Renamed
    hostname                = $env:COMPUTERNAME
    manufacturer            = "Undetermined" # Added
    os_platform             = $null
    os_version              = $null
    os_release              = $null # Build Number
    os_display_version      = $null # e.g., 22H2 (Requires registry read)
    architecture            = $env:PROCESSOR_ARCHITECTURE
    processor               = $null
    cpu_physical_cores      = 0
    cpu_logical_cores       = 0
    cpu_max_speed_ghz       = 0.0
    timestamp_utc           = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    timestamp_local         = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    timezone_name           = (Get-TimeZone).Id
    timezone_offset_utc     = (Get-TimeZone).BaseUtcOffset.ToString() # Format: +/-hh:mm:ss
    ram_total_gb            = 0.0
    ram_speed_mhz           = 0 # Added
    ram_type                = "Undetermined"
    disk_total_gb           = 0.0
    disk_free_gb            = 0.0
    system_drive_type       = "Undetermined"
    tpm_present             = "Undetermined"
    tpm_version             = "Undetermined"
    tpm_enabled             = "Undetermined"
    secure_boot_enabled     = "Undetermined"
    graphics_card           = "Undetermined"
    wddm_version            = "Undetermined"
    pending_updates_count   = -1 # Error/Not Run Default
    collection_error        = $null
    wmi_error_details       = $null
    update_check_error_details = $null
}

# Initialize script-level error list
$script:WmiErrors = [System.Collections.Generic.List[string]]::new()

# 3. Collect Data
try {
    # Basic OS Info
    Write-Host "Collecting Basic OS Info..."
    $osInfo = Get-SafeCimInstance -ClassName Win32_OperatingSystem -FirstOrDefault
    if ($osInfo) {
        $data.os_platform = $osInfo.Caption
        $data.os_version = $osInfo.Version
        $data.os_release = $osInfo.BuildNumber
    } # Error handled in Get-SafeCimInstance

    # Manufacturer
    $csInfo = Get-SafeCimInstance -ClassName Win32_ComputerSystem -FirstOrDefault
    if ($csInfo) {
        $data.manufacturer = $csInfo.Manufacturer
    } # Error handled in Get-SafeCimInstance

    # OS Display Version (Registry)
    try {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $data.os_display_version = (Get-ItemProperty -Path $regPath -Name DisplayVersion -ErrorAction SilentlyContinue).DisplayVersion
        if (-not $data.os_display_version) {
             $data.os_display_version = (Get-ItemProperty -Path $regPath -Name ReleaseId -ErrorAction SilentlyContinue).ReleaseId
        }
    } catch { Write-Warning "Could not read OS DisplayVersion/ReleaseId from registry: $($_.Exception.Message)" }

    # CPU Info
    Write-Host "Collecting CPU Info..."
    $cpuInfo = Get-SafeCimInstance -ClassName Win32_Processor -FirstOrDefault
    if ($cpuInfo) {
        $data.processor = $cpuInfo.Name.Trim()
        $data.cpu_physical_cores = $cpuInfo.NumberOfCores
        $data.cpu_logical_cores = $cpuInfo.NumberOfLogicalProcessors
        $data.cpu_max_speed_ghz = [math]::Round($cpuInfo.MaxClockSpeed / 1000, 2)
    } # Error handled in Get-SafeCimInstance

    # RAM Info
    Write-Host "Collecting RAM Info..."
    try {
        # Use Get-CimInstance directly here as Get-SafeCimInstance doesn't handle collections for Measure-Object well
        $totalRamBytes = (Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction Stop | Measure-Object -Property Capacity -Sum).Sum
        $data.ram_total_gb = [math]::Round($totalRamBytes / 1GB, 2)
    } catch {
         Write-Warning "Failed to calculate total RAM: $($_.Exception.Message)"
         $script:WmiErrors.Add("Failed Win32_PhysicalMemory sum: $($_.Exception.Message)")
    }
    # RAM Type & Speed (Get first module)
    $memModule = Get-SafeCimInstance -ClassName Win32_PhysicalMemory -FirstOrDefault
    if ($memModule) {
         # RAM Type Mapping
         $memTypes = @{ 0="Unknown"; 1="Other"; 2="Unknown"; 3="DRAM"; 4="EDRAM"; 5="VRAM"; 6="SRAM"; 7="RAM"; 8="ROM"; 9="SDRAM"; 10="SGRAM"; 11="RDRAM"; 12="EDO"; 17="FPM"; 18="EDO"; 19="FPM"; 20="DDR"; 21="DDR2"; 22="DDR2 FB-DIMM"; 24="DDR3"; 25="FBD2"; 26="DDR4"; 27="LPDDR"; 28="LPDDR2"; 29="LPDDR3"; 30="LPDDR4"; 31="Logical NV-DIMM"; 32="HBM"; 33="HBM2"; 34="DDR5"; 35="LPDDR5" }
         $typeCode = $memModule.MemoryType
         $data.ram_type = $memTypes.Item($typeCode) # Use .Item() for safer access
         if (-not $data.ram_type) { $data.ram_type = "Unknown Code ($typeCode)" }

         # RAM Speed
         if ($memModule.Speed -ne $null) { # Check if Speed property exists and is not null
             $data.ram_speed_mhz = $memModule.Speed
         } else {
             $data.ram_speed_mhz = 0 # Or "Undetermined" if preferred
             Write-Warning "Could not determine RAM speed from WMI."
         }
    } # Error for query handled in Get-SafeCimInstance

    # Disk Info (System Drive)
    Write-Host "Collecting System Disk Info..."
    try {
        $sysDriveLetter = $env:SystemDrive[0]
        $sysVolume = Get-Volume -DriveLetter $sysDriveLetter -ErrorAction Stop
        $data.disk_total_gb = [math]::Round($sysVolume.Size / 1GB, 2)
        $data.disk_free_gb = [math]::Round($sysVolume.SizeRemaining / 1GB, 2)

        # Drive Type (Primary Method - Needs Admin)
        try {
            # Get partition number associated with the drive letter
            $partition = Get-Partition | Where-Object { $_.DriveLetter -eq $sysDriveLetter } | Select-Object -First 1 -ErrorAction Stop
            if ($partition) {
                $physDisk = Get-PhysicalDisk -Number $partition.DiskNumber -ErrorAction Stop
                if ($physDisk) {
                    $diskTypes = @{ 3 = "HDD"; 4 = "SSD"; 5 = "SCM"; 0 = "Unspecified" }
                    $mediaCode = $physDisk.MediaType
                    if ($diskTypes.ContainsKey($mediaCode)) {
                        $data.system_drive_type = $diskTypes[$mediaCode]
                    } else {
                        $data.system_drive_type = "Unknown Code ($mediaCode)"
                    }
                } else { throw "Could not find physical disk for partition." }
            } else { throw "Could not find partition for system drive." }
        } catch {
            Write-Warning "Primary drive type check failed (needs Admin?): $($_.Exception.Message)"
            $script:WmiErrors.Add("Failed Get-PhysicalDisk/Get-Partition query: $($_.Exception.Message)")
            # Drive Type (Fallback Method)
            Write-Host "Attempting fallback drive type check..."
            # Need DiskNumber from partition, try getting it again safely
            $diskNumber = $null
            try { $diskNumber = (Get-Partition | Where-Object { $_.DriveLetter -eq $sysDriveLetter }).DiskNumber } catch {}

            if ($diskNumber -ne $null) {
                 $diskDrive = Get-SafeCimInstance -ClassName Win32_DiskDrive -Filter "Index=$diskNumber" -FirstOrDefault
                 if ($diskDrive -and $diskDrive.Model) {
                     if ($diskDrive.Model -ilike "*SSD*") { $data.system_drive_type = "SSD (Heuristic)" }
                     else { $data.system_drive_type = "HDD (Heuristic)" }
                 } else {
                      $data.system_drive_type = "Fallback Failed (No Win32_DiskDrive)"
                      $script:WmiErrors.Add("Failed Win32_DiskDrive fallback query")
                 }
            } else {
                 $data.system_drive_type = "Fallback Failed (No Partition Info)"
                 $script:WmiErrors.Add("Failed Get-Partition for fallback")
            }
        }
    } catch {
        Write-Warning "Failed to get system volume info: $($_.Exception.Message)"
        $script:WmiErrors.Add("Failed Get-Volume query: $($_.Exception.Message)")
    }

    # Graphics Card
    Write-Host "Collecting Graphics Info..."
    $gpuInfo = Get-SafeCimInstance -ClassName Win32_VideoController -FirstOrDefault
    if ($gpuInfo) {
        $data.graphics_card = $gpuInfo.Name
        $data.wddm_version = "Driver: $($gpuInfo.DriverVersion)" # Simplified
    } # Error handled in Get-SafeCimInstance

    # TPM Info (Needs Admin)
    Write-Host "Collecting TPM Info (requires Admin)..."
    try {
        if (Get-Command Get-Tpm -ErrorAction SilentlyContinue) {
            $tpm = Get-Tpm -ErrorAction Stop
            $data.tpm_present = $tpm.TpmPresent
            $data.tpm_enabled = $tpm.TpmReady -and $tpm.TpmEnabled
            if ($tpm.ManufacturerVersion -match '(\d+)\.(\d+)') { $data.tpm_version = "$($matches[1]).$($matches[2])" }
            elseif ($tpm.ManufacturerId -ne 0) { $data.tpm_version = "Present (Version Unknown)" }
            else { $data.tpm_version = "N/A" }
        } else {
             Write-Warning "Get-Tpm cmdlet not found."
             $data.tpm_present = "Check Failed (Cmdlet Missing)"; $data.tpm_version = $data.tpm_present; $data.tpm_enabled = $data.tpm_present
        }
    } catch {
        Write-Warning "TPM check failed (needs Admin?): $($_.Exception.Message)"
        $script:WmiErrors.Add("Failed Get-Tpm: $($_.Exception.Message)")
        $data.tpm_present = "Check Failed (Error/Permissions?)"; $data.tpm_version = $data.tpm_present; $data.tpm_enabled = $data.tpm_present
    }

    # Secure Boot (Needs Admin, UEFI)
    Write-Host "Checking Secure Boot Status (requires Admin, UEFI)..."
    try {
        if (Get-Command Confirm-SecureBootUEFI -ErrorAction SilentlyContinue) {
            $sbStatus = Confirm-SecureBootUEFI -ErrorAction Stop
            $data.secure_boot_enabled = $sbStatus
        } else {
             Write-Warning "Confirm-SecureBootUEFI cmdlet not found."
             $data.secure_boot_enabled = "Check Failed (Cmdlet Missing)"
        }
    } catch {
        if ($_.Exception.Message -match 'not supported on computers that do not support UEFI') {
             Write-Warning "Secure Boot check failed: System is not UEFI."
             $data.secure_boot_enabled = "Not Applicable (Non-UEFI)"
             $script:WmiErrors.Add("Secure Boot: Not UEFI")
        } else {
            Write-Warning "Secure Boot check failed (needs Admin?): $($_.Exception.Message)"
            $script:WmiErrors.Add("Failed Confirm-SecureBootUEFI: $($_.Exception.Message)")
            $data.secure_boot_enabled = "Check Failed (Error/Permissions?)"
        }
    }

    # Pending Updates (Needs Admin, COM API)
    Write-Host "Checking Pending Updates (requires Admin)..."
    try {
        $updateSession = New-Object -ComObject "Microsoft.Update.Session" -ErrorAction Stop
        $updateSearcher = $updateSession.CreateUpdateSearcher()
        $searchCriteria = "IsInstalled=0 and IsHidden=0 and Type='Software'"
        Write-Host "Searching for available updates (this may take a moment)..."
        $searchResult = $updateSearcher.Search($searchCriteria)
        $data.pending_updates_count = $searchResult.Updates.Count
        Write-Host "Found $($data.pending_updates_count) applicable updates."
    } catch {
        Write-Warning "Pending update check failed: $($_.Exception.Message)"
        $data.pending_updates_count = -3 # Indicate COM/API error
        $data.update_check_error_details = "$($_.Exception.GetType().FullName): $($_.Exception.Message)"
    }

    # Store combined WMI errors if any occurred
    if ($script:WmiErrors.Count -gt 0) {
        $data.wmi_error_details = $script:WmiErrors -join "; "
    }

} catch {
    # Catch major script errors during collection
    Write-Error "A critical error occurred during data collection: $($_.Exception.Message)"
    $data.collection_error = "$($_.Exception.GetType().FullName): $($_.Exception.Message)"
}

# 4. Send Data to API
Write-Host "Attempting to send data to API: $Script:ApiEndpointUrl"
try {
    $jsonBody = $data | ConvertTo-Json -Depth 5 -Compress
    # Write-Host $jsonBody # Uncomment for debugging payload

    $response = Invoke-RestMethod -Method Post -Uri $Script:ApiEndpointUrl -ContentType 'application/json' -Body $jsonBody -TimeoutSec 20 -ErrorAction Stop

    Write-Host "Data submitted successfully for Host: $($data.hostname), Assessment ID: $($data.assessment_id)." -ForegroundColor Green # Updated ID name

} catch {
    Write-Error "Failed to send data to API: $($_.Exception.Message)"
    Write-Host "Please contact support at $Script:SupportPhoneNumber" -ForegroundColor Yellow

    # Save data locally on failure
    try {
        $tempDir = $env:TEMP
        $fileName = "readiness_check_$($data.hostname)_$($data.assessment_id -replace '[^a-zA-Z0-9_-]','_')_$(Get-Date -Format 'yyyyMMdd_HHmmss').json" # Save as JSON, Updated ID name
        $filePath = Join-Path -Path $tempDir -ChildPath $fileName
        Write-Host "Saving data locally to: $filePath"
        $data | ConvertTo-Json -Depth 5 | Out-File -FilePath $filePath -Encoding utf8 -ErrorAction Stop
        Write-Host "Data saved successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to save data locally: $($_.Exception.Message)"
    }
}

Write-Host "Readiness Check Complete." -ForegroundColor Cyan