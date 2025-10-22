<# : hybrid batch + powershell script
@powershell -noprofile -ExecutionPolicy Bypass -Window Normal -c "$param='%*';$ScriptPath='%~f0';iex((Get-Content('%~f0') | Out-String))"&exit/b
#>

Write-Host ""
Write-Host "Gathering system information, please wait..."

# --------------------------------------------------------------------------------------
# 1. GATHER BASIC INFO AND WMI/CIM OBJECTS
# --------------------------------------------------------------------------------------

$computerName = $env:COMPUTERNAME
$userName = $env:USERNAME
# --- DEÐÝÞÝKLÝK: Get-CimInstance -> Get-WmiObject ve -ClassName -> -Class olarak deðiþtirildi ---
$osInfo = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue
$mbInfo = Get-WmiObject -Class Win32_BaseBoard -ErrorAction SilentlyContinue
$cpuInfo = Get-WmiObject -Class Win32_Processor -ErrorAction SilentlyContinue
$diskDrives = Get-WmiObject -Class Win32_DiskDrive -ErrorAction SilentlyContinue
$physMem = Get-WmiObject -Class Win32_PhysicalMemory -ErrorAction SilentlyContinue
$memArray = Get-WmiObject -Class Win32_PhysicalMemoryArray -ErrorAction SilentlyContinue
$videoControllers = Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue
$netAdaptersConfig = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue

# --------------------------------------------------------------------------------------
# 2. PROCESS RAM INFORMATION
# --------------------------------------------------------------------------------------

$totalRamGB_Display = 0
$maxSupportedRamGB = 0
$ramDetails = ""
$totalSlots = 0
$occupiedSlots = 0
$emptySlots = 0

if ($memArray) {
    # MaxCapacity is in KB, convert to GB
    $maxSupportedRamGB = [Math]::Round($memArray.MaxCapacity / 1MB) 
    $totalSlots = $memArray.MemoryDevices
}
if ($physMem) {
    $i = 1
    # Convert to array even if only one slot is found
    $allMemSlots = @($physMem) 
    
    foreach ($mem in $allMemSlots) {
        $memSizeGB = [Math]::Round($mem.Capacity / 1GB)
        $ramSpeed = $mem.Speed
        $ramType = ""

        # Determine RAM Type based on speed (same logic as VBS)
        if ($ramSpeed -ge 1600 -and $ramSpeed -lt 2133) { $ramType = "DDR3" }
        elseif ($ramSpeed -ge 2133 -and $ramSpeed -lt 2933) { $ramType = "DDR4" }
        elseif ($ramSpeed -ge 2933) { $ramType = "DDR5" }
        else { $ramType = "Unknown" }
        
        # `n = Newline (vbCrLf)
        # ${i} = To separate the variable from the ':'
        $ramDetails += "Slot ${i}: $memSizeGB GB, Speed: $ramSpeed MHz, Type: $ramType`n"
        $occupiedSlots++
        $i++
    }
    
    # Calculate total RAM more accurately
    $totalRamBytes = ($allMemSlots | Measure-Object -Property Capacity -Sum).Sum
    $totalRamGB_Display = [Math]::Round($totalRamBytes / 1GB)
}
$emptySlots = $totalSlots - $occupiedSlots

# --------------------------------------------------------------------------------------
# 3. ARCHITECTURE INFORMATION
# --------------------------------------------------------------------------------------

if ($cpuInfo) {
    $allCpus = @($cpuInfo) 

    switch ($allCpus[0].Architecture) {
        0 { $cpuArchitecture = "x86" }
        9 { $cpuArchitecture = "x64" }
        5 { $cpuArchitecture = "ARM" }
        6 { $cpuArchitecture = "Itanium" }
        12 { $cpuArchitecture = "ARM64" }
        default { 
            $archValue = $allCpus[0].Architecture
            if ($null -eq $archValue) {
                $cpuArchitecture = "Unknown (null)"
            } else {
                $cpuArchitecture = "Unknown ($archValue)" 
            }
        }
    }
} else {
    $cpuArchitecture = "Unknown (No CPU information)"
}

# --------------------------------------------------------------------------------------
# 4. GRAPHICS CARD (GPU) INFORMATION
# --------------------------------------------------------------------------------------

$gpuDetails = ""
if ($videoControllers) {
    # Handle multiple GPUs
    foreach ($gpu in $videoControllers) {
        $gpuDetails += "Model     : $($gpu.Description)`n"
    }
}

# --------------------------------------------------------------------------------------
# 5. NETWORK INFORMATION (LOCAL IP, WAN IP, DNS)
# --------------------------------------------------------------------------------------

$ipDetails = ""
$counter = 1

$validAdapters = $netAdaptersConfig | Where-Object { $_.IPEnabled -eq $true -and $_.Description -notmatch "WAN Miniport" -and $_.Description -notmatch "Microsoft" }

if ($validAdapters) {
    foreach ($adapter in $validAdapters) {
        $ipAddr = $adapter.IPAddress | Select-Object -First 1; if (-not $ipAddr) { $ipAddr = "Not found" }
        $macAddr = $adapter.MACAddress; if (-not $macAddr) { $macAddr = "Not found" }
        $dhcpServer = $adapter.DHCPServer; if (-not $dhcpServer) { $dhcpServer = "Not found" }

        # ${counter} = To separate the variable from the ':'
        $ipDetails += "Network Card ${counter}:`n"
        $ipDetails += "$($adapter.Description)`n"
        $ipDetails += "MAC Address`t: $macAddr`n"
        $ipDetails += "IP Address`t: $ipAddr`n"
        $ipDetails += "DHCP Server`t: $dhcpServer`n`n"
        
        $counter++
    }
}

# --- WAN IP and Ping Test ---
$WANIP = "Not found"
$pingSuccess = $false

Write-Host "Checking internet connection..."
if (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue) { $pingSuccess = $true }

if ($pingSuccess) {
    Write-Host "Getting WAN IP address..."
    try {
        # Invoke-RestMethod is the modern way (replaces MSXML2.XMLHTTP)
        $WANIP = Invoke-RestMethod -Uri "http://api.ipify.org" -TimeoutSec 5 
    }
    catch { $WANIP = "Not found (Service unavailable)" }
}

$ipDetails += "WAN IP Address`t: $WANIP`n"

# --- DNS Servers ---
$dnsServers = $null
$adapterWithDns = $validAdapters | Where-Object { $_.DNSServerSearchOrder } | Select-Object -First 1

if ($adapterWithDns) {
    # Join the array of DNS servers with " / "
    $dnsServers = $adapterWithDns.DNSServerSearchOrder -join " / " 
} else {
    $dnsServers = "Not found"
}

$ipDetails += "DNS Server`t: $dnsServers`n"

# --------------------------------------------------------------------------------------
# 6. OS INSTALLATION DATE
# --------------------------------------------------------------------------------------

$formattedInstallDate = "Not found"
if ($osInfo) {
    # Get-CimInstance auto-converts WMI date to [datetime]
    $wmiDate = $osInfo.InstallDate
    try {
        $installDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($wmiDate)
        $formattedInstallDate = $installDate.ToString("dd.MM.yyyy HH:mm:ss")
    } catch {
        $formattedInstallDate = "Date could not be read"
    }
}

# --------------------------------------------------------------------------------------
# 7. DISK INFORMATION 
# --------------------------------------------------------------------------------------

function Get-DriveMediaTypeFromVBSLogic {
    param ([psobject]$DiskDrive)
    try {
        if ($null -ne $DiskDrive) {
            $model = $DiskDrive.Model.ToLower()
            $interfaceType = ""; if ($DiskDrive.PSObject.Properties.Name -contains 'InterfaceType') { $interfaceType = $DiskDrive.InterfaceType.ToLower() }
            
            # Match VBS logic
            if ($model -match "nvme" -or $model -match "nvm") { return "NVMe" }
            elseif ($model -match "ssd" -or $model -match "sd") { return "SSD" }
            elseif ($interfaceType -eq "usb") { return "USB" }
            else { return "HDD" }
        } else { return "Unknown" }
    } catch { return "Unknown (Error)" }
}

$diskInfo = "Disk Summary   :`n"
if ($diskDrives) {
    foreach ($drive in $diskDrives) {
        try {
            $diskSizeGB = [Math]::Round($drive.Size / 1GB, 2)
            $driveType = Get-DriveMediaTypeFromVBSLogic -DiskDrive $drive
            $diskInfo += "$($drive.DeviceID) - $driveType - Capacity: $diskSizeGB GB`n"
        } catch { 
            $diskInfo += "$($drive.DeviceID) - Disk status unreadable`n" 
        }
    }
} else { 
    $diskInfo += "Disk information not found.`n" 
}

# --------------------------------------------------------------------------------------
# 8. COMBINE ALL INFORMATION
# --------------------------------------------------------------------------------------

$dividerLine = "-----------------------------------------------------------------------------"

$tMessage = "Operating System`t`t: $($($osInfo.Caption -replace 'Microsoft ', '').Trim())`n"
$tMessage += "OS Version`t`t: $($osInfo.Version)`n"
$tMessage += "Windows Architecture`t: $cpuArchitecture`n"
$tMessage += "User Name`t`t: $userName`n"
$tMessage += "Computer Name`t`t: $computerName`n"
$tMessage += "Last Format Date`t`t: $formattedInstallDate`n"
$tMessage += "$dividerLine`n" 
$tMessage += "Motherboard Manufacturer`t: $($mbInfo.Manufacturer)`n"
$tMessage += "Motherboard Model`t: $($mbInfo.Product)`n"
$tMessage += "Processor`t`t: $($cpuInfo.Manufacturer)`n"
$tMessage += "Processor Model`t`t: $($cpuInfo.Name)`n"
$tMessage += "CPU Architecture`t`t: $cpuArchitecture`n"
$tMessage += "Total RAM`t`t: $totalRamGB_Display GB`n"
$tMessage += "Max Supported RAM`t: $maxSupportedRamGB GB`n"
$tMessage += "Empty RAM Slots`t`t: $emptySlots`n"
$tMessage += "RAM Slots`t`t: `n$ramDetails`n" 
$tMessage += "Graphics Card(s)`t`t: `n$gpuDetails"
$tMessage += "$dividerLine`n" 
$tMessage += "Network Card(s) and IP Address(es) :`n`n$ipDetails"
$tMessage += "$dividerLine`n" 
$tMessage += "$diskInfo" 

# --------------------------------------------------------------------------------------
# 9. SHOW POPUP AND SAVE
# --------------------------------------------------------------------------------------

# Create the WScript.Shell COM object
$WshShell = New-Object -ComObject WScript.Shell
Write-Host "Displaying information..."
$WshShell.Popup($tMessage, 0, "Hardware Information |  | by Abdullah ERTÜRK", 0 + 64 + 4096)
$userResponse = $WshShell.Popup("Save system information to Desktop?", 0, "Confirmation", 4 + 32 + 4096)

if ($userResponse -eq 6) { # 6 = Yes
    try {
        $formattedDate = (Get-Date).ToString("yyyy_MM_dd")
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        # Note: File name uses translated text
        $fileName = "${computerName}_System_Information_${formattedDate}.txt"
        $filePath = Join-Path -Path $desktopPath -ChildPath $fileName

        # Set-Content is the easiest way to write a file
        Set-Content -Path $filePath -Value $tMessage -Encoding UTF8
        
        # 64 (Info Icon) + 4096 (Always on Top)
        $WshShell.Popup("System information saved to $filePath", 0, "Information", 64 + 4096)
    
    } catch {
        # 16 (Error Icon) + 4096 (Always on Top)
        $WshShell.Popup("Failed to save file: $($_.Exception.Message)", 0, "Error", 16 + 4096)
    }
}

Write-Host "Operation complete."