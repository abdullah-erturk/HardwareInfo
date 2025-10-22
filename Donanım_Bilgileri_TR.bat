<# : hybrid batch + powershell script
@powershell -noprofile -ExecutionPolicy Bypass -Window Normal -c "$param='%*';$ScriptPath='%~f0';iex((Get-Content('%~f0') | Out-String))"&exit/b
#>

Write-Host ""
Write-Host "Sistem bilgileri toplan�yor, l�tfen bekleyin..."

# --------------------------------------------------------------------------------------
# 1. TEMEL B�LG�LER� VE WMI/CIM NESNELER�N� TOPLAMA (Windows 7 uyumlu)
# --------------------------------------------------------------------------------------

$computerName = $env:COMPUTERNAME
$userName = $env:USERNAME
# Get-CimInstance yerine Get-WmiObject kullan�l�yor
$osInfo = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue
$mbInfo = Get-WmiObject -Class Win32_BaseBoard -ErrorAction SilentlyContinue
$cpuInfo = Get-WmiObject -Class Win32_Processor -ErrorAction SilentlyContinue
$diskDrives = Get-WmiObject -Class Win32_DiskDrive -ErrorAction SilentlyContinue
$physMem = Get-WmiObject -Class Win32_PhysicalMemory -ErrorAction SilentlyContinue
$memArray = Get-WmiObject -Class Win32_PhysicalMemoryArray -ErrorAction SilentlyContinue
$videoControllers = Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue
$netAdaptersConfig = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue

# --------------------------------------------------------------------------------------
# 2. RAM B�LG�LER�N� ��LEME
# --------------------------------------------------------------------------------------

$totalRamGB_Display = 0
$maxSupportedRamGB = 0
$ramDetails = ""
$totalSlots = 0
$occupiedSlots = 0
$emptySlots = 0

if ($memArray) {
    $maxSupportedRamGB = [Math]::Round($memArray.MaxCapacity / 1MB) 
    $totalSlots = $memArray.MemoryDevices
}
if ($physMem) {
    $i = 1
    $allMemSlots = @($physMem) # Bu @() kullan�m�, tek RAM'de bile koleksiyon olmas�n� sa�lar
    
    foreach ($mem in $allMemSlots) {
        $memSizeGB = [Math]::Round($mem.Capacity / 1GB)
        $ramSpeed = $mem.ConfiguredClockSpeed # Get-WmiObject'te �zellik ad� farkl� olabilir
        if (-not $ramSpeed) { $ramSpeed = $mem.Speed } # Alternatif �zellik

        $ramType = ""
        if ($ramSpeed -ge 1600 -and $ramSpeed -lt 2133) { $ramType = "DDR3" }
        elseif ($ramSpeed -ge 2133 -and $ramSpeed -lt 2933) { $ramType = "DDR4" }
        elseif ($ramSpeed -ge 2933) { $ramType = "DDR5" }
        else { $ramType = "Bilinmiyor" }
        
        $ramDetails += "Slot ${i}: $memSizeGB GB, H�z: $ramSpeed MHz, T�r: $ramType`n"
        $occupiedSlots++
        $i++
    }
    
    $totalRamBytes = ($allMemSlots | Measure-Object -Property Capacity -Sum).Sum
    $totalRamGB_Display = [Math]::Round($totalRamBytes / 1GB)
}
$emptySlots = $totalSlots - $occupiedSlots

# --------------------------------------------------------------------------------------
# 3. M�MAR� B�LG�S� (D�ZELT�LM�� B�L�M)
# --------------------------------------------------------------------------------------

$cpuArchitecture = ""
if ($cpuInfo) {
    # --- D�ZELTME BURADA ---
    # $cpuInfo'nun tek bir nesne (koleksiyon de�il) olabilme ihtimaline kar��
    # @() ile bir dizi/koleksiyon olmaya zorluyoruz.
    $allCpus = @($cpuInfo) 

    # Art�k $allCpus[0] g�venle kullan�labilir.
    switch ($allCpus[0].Architecture) {
        0 { $cpuArchitecture = "x86" }
        9 { $cpuArchitecture = "x64" }
        5 { $cpuArchitecture = "ARM" }
        6 { $cpuArchitecture = "Itanium" }
        12 { $cpuArchitecture = "ARM64" }
        default { 
            $archValue = $allCpus[0].Architecture
            if ($null -eq $archValue) {
                $cpuArchitecture = "Bilinmiyor (null)"
            } else {
                $cpuArchitecture = "Bilinmiyor ($archValue)" 
            }
        }
    }
} else {
    $cpuArchitecture = "Bilinmiyor (CPU bilgisi yok)"
}


# --------------------------------------------------------------------------------------
# 4. GRAF�K KARTI (GPU) B�LG�S�
# --------------------------------------------------------------------------------------

$gpuDetails = ""
if ($videoControllers) {
    # foreach d�ng�s� tekil nesnelerde veya koleksiyonlarda sorunsuz �al���r
    foreach ($gpu in $videoControllers) {
        $gpuDetails += "Modeli    : $($gpu.Description)`n"
    }
}

# --------------------------------------------------------------------------------------
# 5. A� B�LG�LER� (YEREL IP, WAN IP, DNS)
# --------------------------------------------------------------------------------------

$ipDetails = ""
$counter = 1

$validAdapters = $netAdaptersConfig | Where-Object { $_.IPEnabled -eq $true -and $_.Description -notmatch "WAN Miniport" -and $_.Description -notmatch "Microsoft" }

if ($validAdapters) {
    foreach ($adapter in $validAdapters) {
        $ipAddr = $adapter.IPAddress | Select-Object -First 1; if (-not $ipAddr) { $ipAddr = "Bulunamad�" }
        $macAddr = $adapter.MACAddress; if (-not $macAddr) { $macAddr = "Bulunamad�" }
        $dhcpServer = $adapter.DHCPServer; if (-not $dhcpServer) { $dhcpServer = "Bulunamad�" }

        $ipDetails += "A� Kart� ${counter}:`n"
        $ipDetails += "$($adapter.Description)`n"
        $ipDetails += "MAC Adresi`t: $macAddr`n"
        $ipDetails += "IP Adresi`t`t: $ipAddr`n"
        $ipDetails += "DHCP Sunucu`t: $dhcpServer`n`n"
        
        $counter++
    }
}

# --- WAN IP ve Ping Testi (Windows 7 uyumlu) ---
$WANIP = "Bulunamad�"
$pingSuccess = $false

Write-Host "�nternet ba�lant�s� kontrol ediliyor..."
if (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue) { $pingSuccess = $true }

if ($pingSuccess) {
    Write-Host "WAN IP adresi al�n�yor..."
    try {
        $WebClient = New-Object System.Net.WebClient
        $WANIP = $WebClient.DownloadString("http://api.ipify.org")
        $WebClient.Dispose()
    }
    catch { $WANIP = "Bulunamad� (Servis eri�ilemiyor)" }
}

$ipDetails += "WAN IP Adresi`t: $WANIP`n"

# --- DNS Sunucular� ---
$dnsServers = $null
$adapterWithDns = $validAdapters | Where-Object { $_.DNSServerSearchOrder } | Select-Object -First 1

if ($adapterWithDns) {
    $dnsServers = $adapterWithDns.DNSServerSearchOrder -join " / " 
} else {
    $dnsServers = "Bulunamad�"
}

$ipDetails += "DNS Sunucu`t: $dnsServers`n"

# --------------------------------------------------------------------------------------
# 6. ��LET�M S�STEM� KURULUM TAR�H� (Windows 7 uyumlu)
# --------------------------------------------------------------------------------------

$formattedInstallDate = "Bulunamad�"
if ($osInfo) {
    # Get-WmiObject tarihi WMI format�nda (string) d�nd�r�r, d�n��t�rme gerekir.
    $wmiDate = $osInfo.InstallDate
    try {
        $installDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($wmiDate)
        $formattedInstallDate = $installDate.ToString("dd.MM.yyyy HH:mm:ss")
    } catch {
        $formattedInstallDate = "Tarih okunamad�"
    }
}

# --------------------------------------------------------------------------------------
# 7. D�SK B�LG�LER� (FONKS�YON �LE)
# --------------------------------------------------------------------------------------

function Get-DriveMediaTypeFromVBSLogic {
    param ([psobject]$DiskDrive)
    try {
        if ($null -ne $DiskDrive) {
            $model = $DiskDrive.Model.ToLower()
            $interfaceType = ""; if ($DiskDrive.PSObject.Properties.Name -contains 'InterfaceType') { $interfaceType = $DiskDrive.InterfaceType.ToLower() }
            
            if ($model -match "nvme" -or $model -match "nvm") { return "NVMe" }
            elseif ($model -match "ssd" -or $model -match "sd") { return "SSD" }
            elseif ($interfaceType -eq "usb") { return "USB" }
            else { return "HDD" }
        } else { return "Bilinmiyor" }
    } catch { return "Bilinmiyor (Hata)" }
}

$diskInfo = "Disk �zeti   :`n"
if ($diskDrives) {
    foreach ($drive in $diskDrives) {
        try {
            $diskSizeGB = [Math]::Round($drive.Size / 1GB, 2)
            $driveType = Get-DriveMediaTypeFromVBSLogic -DiskDrive $drive
            $diskInfo += "$($drive.DeviceID) - $driveType - Kapasite: $diskSizeGB GB`n"
        } catch { 
            $diskInfo += "$($drive.DeviceID) - Disk durumu okunamad�`n" 
        }
    }
} else { 
    $diskInfo += "Disk bilgisi bulunamad�.`n" 
}

# --------------------------------------------------------------------------------------
# 8. T�M B�LG�LER� B�RLE�T�RME
# --------------------------------------------------------------------------------------

$dividerLine = "-----------------------------------------------------------------------------"

# Win32_BaseBoard (mbInfo) ve Win32_OperatingSystem (osInfo) her zaman tek nesne d�ner,
# bu y�zden [0] indeksi KULLANILMAMALIDIR. Do�rudan �zelliklere eri�iyoruz.
$tMessage = "��letim Sistemi`t`t: $($($osInfo.Caption -replace 'Microsoft ', '').Trim())`n"
$tMessage += "��letim Sistemi Versiyonu`t: $($osInfo.Version)`n"
$tMessage += "Windows Mimari Yap�s�`t: $cpuArchitecture`n" # B�l�m 3'teki d�zeltilmi� de�i�ken
$tMessage += "Kullan�c� Ad�`t`t: $userName`n"
$tMessage += "Bilgisayar Ad�`t`t: $computerName`n"
$tMessage += "Son Format Tarihi`t`t: $formattedInstallDate`n"
$tMessage += "$dividerLine`n" 
$tMessage += "Anakart �reticisi`t`t: $($mbInfo.Manufacturer)`n" # D�zeltme: [0] indeksi kald�r�ld�
$tMessage += "Anakart Modeli`t`t: $($mbInfo.Product)`n" # D�zeltme: [0] indeksi kald�r�ld�
$tMessage += "��lemci`t`t`t: $($($allCpus = @($cpuInfo); $allCpus[0].Manufacturer))`n" # CPU i�in de @() garantisi
$tMessage += "��lemci Modeli`t`t: $($($allCpus = @($cpuInfo); $allCpus[0].Name))`n" # CPU i�in de @() garantisi
$tMessage += "CPU Mimarisi`t`t: $cpuArchitecture`n" # B�l�m 3'teki d�zeltilmi� de�i�ken
$tMessage += "Toplam RAM`t`t: $totalRamGB_Display GB`n"
$tMessage += "Desteklenen Toplam RAM`t: $maxSupportedRamGB GB`n"
$tMessage += "Bo� RAM Slotlar�`t`t: $emptySlots`n"
$tMessage += "RAM Yuvalar�`t`t: `n$ramDetails`n" 
$tMessage += "Grafik Kart(lar)�`t`t: `n$gpuDetails"
$tMessage += "$dividerLine`n" 
$tMessage += "A� Kart(lar)� ve IP Adres(ler)i :`n`n$ipDetails"
$tMessage += "$dividerLine`n" 
$tMessage += "$diskInfo" 

# --------------------------------------------------------------------------------------
# 9. POPUP �LE G�STERME VE KAYDETME
# --------------------------------------------------------------------------------------

# WScript.Shell COM nesnesini olu�tur
$WshShell = New-Object -ComObject WScript.Shell
Write-Host "Bilgiler g�steriliyor..."
$WshShell.Popup($tMessage, 0, "Donan�m Bilgileri |  | by Abdullah ERT�RK", 0 + 64 + 4096)
$userResponse = $WshShell.Popup("Sistem bilgileri Masa�st�ne kaydedilsin mi?", 0, "Onay", 4 + 32 + 4096)

if ($userResponse -eq 6) { # 6 = Evet
    try {
        $formattedDate = (Get-Date).ToString("yyyy_MM_dd")
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $fileName = "${computerName}_Sistem_Bilgileri_${formattedDate}.txt"
        $filePath = Join-Path -Path $desktopPath -ChildPath $fileName

        Set-Content -Path $filePath -Value $tMessage -Encoding UTF8
        
        $WshShell.Popup("Sistem bilgileri $filePath konumuna kaydedildi.", 0, "Bilgi", 64 + 4096)
    
    } catch {
        $WshShell.Popup("Dosya kaydedilemedi: $($_.Exception.Message)", 0, "Hata", 16 + 4096)
    }
}

Write-Host "��lem tamamland�."