<# : hybrid batch + powershell script
@powershell -noprofile -ExecutionPolicy Bypass -Window Normal -c "$param='%*';$ScriptPath='%~f0';iex((Get-Content('%~f0') | Out-String))"&exit/b
#>

Write-Host ""
Write-Host "Sistem bilgileri toplanýyor, lütfen bekleyin..."

# --------------------------------------------------------------------------------------
# 1. TEMEL BÝLGÝLERÝ VE WMI/CIM NESNELERÝNÝ TOPLAMA (Windows 7 uyumlu)
# --------------------------------------------------------------------------------------

$computerName = $env:COMPUTERNAME
$userName = $env:USERNAME
# Get-CimInstance yerine Get-WmiObject kullanýlýyor
$osInfo = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue
$mbInfo = Get-WmiObject -Class Win32_BaseBoard -ErrorAction SilentlyContinue
$cpuInfo = Get-WmiObject -Class Win32_Processor -ErrorAction SilentlyContinue
$diskDrives = Get-WmiObject -Class Win32_DiskDrive -ErrorAction SilentlyContinue
$physMem = Get-WmiObject -Class Win32_PhysicalMemory -ErrorAction SilentlyContinue
$memArray = Get-WmiObject -Class Win32_PhysicalMemoryArray -ErrorAction SilentlyContinue
$videoControllers = Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue
$netAdaptersConfig = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue

# --------------------------------------------------------------------------------------
# 2. RAM BÝLGÝLERÝNÝ ÝÞLEME
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
    $allMemSlots = @($physMem) # Bu @() kullanýmý, tek RAM'de bile koleksiyon olmasýný saðlar
    
    foreach ($mem in $allMemSlots) {
        $memSizeGB = [Math]::Round($mem.Capacity / 1GB)
        $ramSpeed = $mem.ConfiguredClockSpeed # Get-WmiObject'te özellik adý farklý olabilir
        if (-not $ramSpeed) { $ramSpeed = $mem.Speed } # Alternatif özellik

        $ramType = ""
        if ($ramSpeed -ge 1600 -and $ramSpeed -lt 2133) { $ramType = "DDR3" }
        elseif ($ramSpeed -ge 2133 -and $ramSpeed -lt 2933) { $ramType = "DDR4" }
        elseif ($ramSpeed -ge 2933) { $ramType = "DDR5" }
        else { $ramType = "Bilinmiyor" }
        
        $ramDetails += "Slot ${i}: $memSizeGB GB, Hýz: $ramSpeed MHz, Tür: $ramType`n"
        $occupiedSlots++
        $i++
    }
    
    $totalRamBytes = ($allMemSlots | Measure-Object -Property Capacity -Sum).Sum
    $totalRamGB_Display = [Math]::Round($totalRamBytes / 1GB)
}
$emptySlots = $totalSlots - $occupiedSlots

# --------------------------------------------------------------------------------------
# 3. MÝMARÝ BÝLGÝSÝ (DÜZELTÝLMÝÞ BÖLÜM)
# --------------------------------------------------------------------------------------

$cpuArchitecture = ""
if ($cpuInfo) {
    # --- DÜZELTME BURADA ---
    # $cpuInfo'nun tek bir nesne (koleksiyon deðil) olabilme ihtimaline karþý
    # @() ile bir dizi/koleksiyon olmaya zorluyoruz.
    $allCpus = @($cpuInfo) 

    # Artýk $allCpus[0] güvenle kullanýlabilir.
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
# 4. GRAFÝK KARTI (GPU) BÝLGÝSÝ
# --------------------------------------------------------------------------------------

$gpuDetails = ""
if ($videoControllers) {
    # foreach döngüsü tekil nesnelerde veya koleksiyonlarda sorunsuz çalýþýr
    foreach ($gpu in $videoControllers) {
        $gpuDetails += "Modeli    : $($gpu.Description)`n"
    }
}

# --------------------------------------------------------------------------------------
# 5. AÐ BÝLGÝLERÝ (YEREL IP, WAN IP, DNS)
# --------------------------------------------------------------------------------------

$ipDetails = ""
$counter = 1

$validAdapters = $netAdaptersConfig | Where-Object { $_.IPEnabled -eq $true -and $_.Description -notmatch "WAN Miniport" -and $_.Description -notmatch "Microsoft" }

if ($validAdapters) {
    foreach ($adapter in $validAdapters) {
        $ipAddr = $adapter.IPAddress | Select-Object -First 1; if (-not $ipAddr) { $ipAddr = "Bulunamadý" }
        $macAddr = $adapter.MACAddress; if (-not $macAddr) { $macAddr = "Bulunamadý" }
        $dhcpServer = $adapter.DHCPServer; if (-not $dhcpServer) { $dhcpServer = "Bulunamadý" }

        $ipDetails += "Að Kartý ${counter}:`n"
        $ipDetails += "$($adapter.Description)`n"
        $ipDetails += "MAC Adresi`t: $macAddr`n"
        $ipDetails += "IP Adresi`t`t: $ipAddr`n"
        $ipDetails += "DHCP Sunucu`t: $dhcpServer`n`n"
        
        $counter++
    }
}

# --- WAN IP ve Ping Testi (Windows 7 uyumlu) ---
$WANIP = "Bulunamadý"
$pingSuccess = $false

Write-Host "Ýnternet baðlantýsý kontrol ediliyor..."
if (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue) { $pingSuccess = $true }

if ($pingSuccess) {
    Write-Host "WAN IP adresi alýnýyor..."
    try {
        $WebClient = New-Object System.Net.WebClient
        $WANIP = $WebClient.DownloadString("http://api.ipify.org")
        $WebClient.Dispose()
    }
    catch { $WANIP = "Bulunamadý (Servis eriþilemiyor)" }
}

$ipDetails += "WAN IP Adresi`t: $WANIP`n"

# --- DNS Sunucularý ---
$dnsServers = $null
$adapterWithDns = $validAdapters | Where-Object { $_.DNSServerSearchOrder } | Select-Object -First 1

if ($adapterWithDns) {
    $dnsServers = $adapterWithDns.DNSServerSearchOrder -join " / " 
} else {
    $dnsServers = "Bulunamadý"
}

$ipDetails += "DNS Sunucu`t: $dnsServers`n"

# --------------------------------------------------------------------------------------
# 6. ÝÞLETÝM SÝSTEMÝ KURULUM TARÝHÝ (Windows 7 uyumlu)
# --------------------------------------------------------------------------------------

$formattedInstallDate = "Bulunamadý"
if ($osInfo) {
    # Get-WmiObject tarihi WMI formatýnda (string) döndürür, dönüþtürme gerekir.
    $wmiDate = $osInfo.InstallDate
    try {
        $installDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($wmiDate)
        $formattedInstallDate = $installDate.ToString("dd.MM.yyyy HH:mm:ss")
    } catch {
        $formattedInstallDate = "Tarih okunamadý"
    }
}

# --------------------------------------------------------------------------------------
# 7. DÝSK BÝLGÝLERÝ (FONKSÝYON ÝLE)
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

$diskInfo = "Disk Özeti   :`n"
if ($diskDrives) {
    foreach ($drive in $diskDrives) {
        try {
            $diskSizeGB = [Math]::Round($drive.Size / 1GB, 2)
            $driveType = Get-DriveMediaTypeFromVBSLogic -DiskDrive $drive
            $diskInfo += "$($drive.DeviceID) - $driveType - Kapasite: $diskSizeGB GB`n"
        } catch { 
            $diskInfo += "$($drive.DeviceID) - Disk durumu okunamadý`n" 
        }
    }
} else { 
    $diskInfo += "Disk bilgisi bulunamadý.`n" 
}

# --------------------------------------------------------------------------------------
# 8. TÜM BÝLGÝLERÝ BÝRLEÞTÝRME
# --------------------------------------------------------------------------------------

$dividerLine = "-----------------------------------------------------------------------------"

# Win32_BaseBoard (mbInfo) ve Win32_OperatingSystem (osInfo) her zaman tek nesne döner,
# bu yüzden [0] indeksi KULLANILMAMALIDIR. Doðrudan özelliklere eriþiyoruz.
$tMessage = "Ýþletim Sistemi`t`t: $($($osInfo.Caption -replace 'Microsoft ', '').Trim())`n"
$tMessage += "Ýþletim Sistemi Versiyonu`t: $($osInfo.Version)`n"
$tMessage += "Windows Mimari Yapýsý`t: $cpuArchitecture`n" # Bölüm 3'teki düzeltilmiþ deðiþken
$tMessage += "Kullanýcý Adý`t`t: $userName`n"
$tMessage += "Bilgisayar Adý`t`t: $computerName`n"
$tMessage += "Son Format Tarihi`t`t: $formattedInstallDate`n"
$tMessage += "$dividerLine`n" 
$tMessage += "Anakart Üreticisi`t`t: $($mbInfo.Manufacturer)`n" # Düzeltme: [0] indeksi kaldýrýldý
$tMessage += "Anakart Modeli`t`t: $($mbInfo.Product)`n" # Düzeltme: [0] indeksi kaldýrýldý
$tMessage += "Ýþlemci`t`t`t: $($($allCpus = @($cpuInfo); $allCpus[0].Manufacturer))`n" # CPU için de @() garantisi
$tMessage += "Ýþlemci Modeli`t`t: $($($allCpus = @($cpuInfo); $allCpus[0].Name))`n" # CPU için de @() garantisi
$tMessage += "CPU Mimarisi`t`t: $cpuArchitecture`n" # Bölüm 3'teki düzeltilmiþ deðiþken
$tMessage += "Toplam RAM`t`t: $totalRamGB_Display GB`n"
$tMessage += "Desteklenen Toplam RAM`t: $maxSupportedRamGB GB`n"
$tMessage += "Boþ RAM Slotlarý`t`t: $emptySlots`n"
$tMessage += "RAM Yuvalarý`t`t: `n$ramDetails`n" 
$tMessage += "Grafik Kart(lar)ý`t`t: `n$gpuDetails"
$tMessage += "$dividerLine`n" 
$tMessage += "Að Kart(lar)ý ve IP Adres(ler)i :`n`n$ipDetails"
$tMessage += "$dividerLine`n" 
$tMessage += "$diskInfo" 

# --------------------------------------------------------------------------------------
# 9. POPUP ÝLE GÖSTERME VE KAYDETME
# --------------------------------------------------------------------------------------

# WScript.Shell COM nesnesini oluþtur
$WshShell = New-Object -ComObject WScript.Shell
Write-Host "Bilgiler gösteriliyor..."
$WshShell.Popup($tMessage, 0, "Donaným Bilgileri |  | by Abdullah ERTÜRK", 0 + 64 + 4096)
$userResponse = $WshShell.Popup("Sistem bilgileri Masaüstüne kaydedilsin mi?", 0, "Onay", 4 + 32 + 4096)

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

Write-Host "Ýþlem tamamlandý."