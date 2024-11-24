Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2") ' Kullanıcıya daha uyumlu WMI bağlantısı
Set objNetwork = CreateObject("Wscript.Network")
Set wshShell = CreateObject("WScript.Shell")
strComputerName = wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
Set oShell = WScript.CreateObject("WScript.Shell")
proc_arch = oShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
Set oEnv = oShell.Environment("SYSTEM")

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Set colMB = objWMIService.ExecQuery("Select * from Win32_BaseBoard")
Set colProcessors = objWMIService.ExecQuery("Select * from Win32_Processor")
Set colDrives = objWMIService.ExecQuery("Select * from Win32_DiskDrive")

' Toplam RAM hesaplama
Set obj = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
TotalRam = 0
ramDetails = ""
i = 1

For Each obj2 In obj
    memTmp1 = obj2.capacity / 1024 / 1024
    TotalRam = TotalRam + memTmp1
    ramSpeed = obj2.Speed
    ramType = ""

    If ramSpeed >= 1600 And ramSpeed < 2133 Then
        ramType = "DDR3"
    ElseIf ramSpeed >= 2133 And ramSpeed < 2933 Then
        ramType = "DDR4"
    ElseIf ramSpeed >= 2933 Then
        ramType = "DDR5"
    Else
        ramType = "Unknown"
    End If

    ramDetails = ramDetails & "Slot " & i & ": " & Int(memTmp1 / 1024) & " GB, Hız: " & obj2.Speed & " MHz, Tür: " & ramType & vbCrLf
    i = i + 1
Next

' İşlemci mimarisi
Set colItems = objWMIService.ExecQuery("Select Architecture from Win32_Processor")
For Each objItem in colItems
    If objItem.Architecture = 0 Then
       strArchitecture = "x86"
    ElseIf objItem.Architecture = 9 Then
       strArchitecture = "x64"
    End If
Next

' Grafik Kartı bilgisi
On Error Resume Next
Set colItemsx = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
Dim tStr, tStr2
tStr = ""
tStr2 = ""

' Ekran kartı bilgilerini al
For Each objItem in colItemsx
    tStr = tStr & "Modeli    : " & objItem.Description & vbCrLf
    
    ' Bellek miktarını al
    Dim memSize
    memSize = objItem.AdapterRAM / 1024 / 1024 ' MB cinsinden

    ' Bellek miktarını kontrol et
    If InStr(LCase(objItem.Description), "intel") > 0 Then
        ' Dahili ekran kartı ise, bellek miktarını sistem RAM'inden dinamik olarak alabiliriz
        If memSize < 128 Then
            memSize = 128 ' Dahili ekran kartı için varsayılan bellek miktarı 128 MB
        End If
    Else
        ' Harici ekran kartı ise, AdapterRAM değerini kontrol et
        If memSize < 1024 Then
            ' Bellek 1 GB'den küçükse, muhtemelen yanlış bir değer döndürülüyor
            memSize = 4096 ' Harici ekran kartı için varsayılan bellek miktarı 4096 MB (4 GB)
        End If
    End If

    ' Bellek bilgisi hatalı gösteriyor, bu sebeple yorum satırı :) 
    ' tStr = tStr & "Bellek    	: " & memSize & " MB" & vbCrLf
Next
On Error GoTo 0

' Ağ kartı bilgilerini al
Dim myIPAddresses : myIPAddresses = ""
Dim counter : counter = 1
Dim colAdapters : Set colAdapters = objWMIService.ExecQuery("Select IPAddress, Description, MACAddress, DHCPServer from Win32_NetworkAdapterConfiguration")

For Each objAdapter in colAdapters
    description = objAdapter.Description
    macAddr = objAdapter.MACAddress

    If InStr(description, "WAN Miniport") = 0 And InStr(description, "Microsoft") = 0 Then
        If Not IsNull(objAdapter.IPAddress) Then
            ipAddr = objAdapter.IPAddress(0)
        Else
            ipAddr = "Bulunamadı"
        End If

        If Not IsNull(objAdapter.DHCPServer) Then
            dhcpServer = objAdapter.DHCPServer
        Else
            dhcpServer = "Bulunamadı"
        End If

        myIPAddresses = myIPAddresses & "Ağ Kartı " & counter & "" & vbCrLf & _
                        "Açıklama		: " & description & vbCrLf & _
                        "MAC Adresi	: " & macAddr & vbCrLf & _
                        "IP Adresi		: " & ipAddr & vbCrLf & _
                        "DHCP Sunucu	: " & dhcpServer & vbCrLf & vbCrLf

        counter = counter + 1
    End If
Next

' WAN IP almak için dış web servisini kullan
Dim WANIP
On Error Resume Next ' Hata kontrolü
Dim objXMLHttp
Set objXMLHttp = CreateObject("MSXML2.XMLHTTP")
objXMLHttp.Open "GET", "http://api.ipify.org", False
objXMLHttp.Send

' İnternet bağlantısı kontrolü
If Err.Number <> 0 Then
    WANIP = "Bulunamadı" ' İnternet yoksa WAN IP'yi "Bulunamadı" olarak ayarla
Else
    WANIP = objXMLHttp.responseText ' İnternet bağlantısı varsa WAN IP'yi al
End If
On Error GoTo 0 ' Hata kontrolünü sıfırla

' Ping ile internet bağlantısını kontrol et
Dim pingResult
pingResult = PingHost("8.8.8.8")

If pingResult Then
    ' 8.8.8.8'e ping gidiyorsa WAN IP'yi göster
    If WANIP = "Bulunamadı" Then
        WANIP = objXMLHttp.responseText ' Gerçek WAN IP'yi al
    End If
Else
    ' Eğer 8.8.8.8'e ping gitmiyorsa, WAN IP olarak "Bulunamadı" yaz
    WANIP = "Bulunamadı"
End If

myIPAddresses = myIPAddresses & "WAN IP Adresi	: " & WANIP & vbCrLf

' DNS sunucu adreslerini ekle
Set colNicConfigs = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
Dim dnsFound
dnsFound = False ' DNS bulunup bulunmadığını takip etmek için

For Each objNicConfig In colNicConfigs
    If Not IsNull(objNicConfig.DNSServerSearchOrder) Then
        myIPAddresses = myIPAddresses & "DNS Sunucu	: " 
        Dim dnsList
        dnsList = ""
        
        For Each dnsServer In objNicConfig.DNSServerSearchOrder
            If dnsList = "" Then
                dnsList = dnsServer ' İlk DNS sunucusu
            Else
                dnsList = dnsList & " / " & dnsServer ' Sonraki DNS sunucuları arasına "/" ekle
            End If
        Next
        
        myIPAddresses = myIPAddresses & dnsList & vbCrLf
        dnsFound = True ' DNS bulunduğu için işaretle
        Exit For
    End If
Next

If Not dnsFound Then
    myIPAddresses = myIPAddresses & "DNS Sunucu	: Bulunamadı" & vbCrLf
End If

' Ping fonksiyonu
Function PingHost(host)
    Dim objShell, command, result
    Set objShell = CreateObject("WScript.Shell")
    command = "ping -n 1 " & host ' "-n 1" parametresi tek bir ping isteği gönderir
    result = objShell.Run(command, 0, True) ' Komutu çalıştır ve sonucu al
    If result = 0 Then
        PingHost = True ' Ping başarılı
    Else
        PingHost = False ' Ping başarısız
    End If
End Function

' İşletim sistemi bilgisi
Set dtmInstallDate = CreateObject("WbemScripting.SWbemDateTime")
Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
    fthx = getmydat(objOperatingSystem.InstallDate)
    Exit For
Next

Function getmydat(wmitime)
    dtmInstallDate.Value = wmitime
    getmydat = dtmInstallDate.GetVarDate
End Function


' Disk bilgilerini al ve türlerini kontrol et
Dim diskInfo
diskInfo = "Disk Özeti  :" & vbCrLf

' Eğer colDrives boş değilse işleme başla
If colDrives.Count > 0 Then
    For Each objDrive In colDrives
        If Not objDrive Is Nothing Then
            ' Diskin hazır olup olmadığını kontrol et
            On Error Resume Next ' Hataları geçici olarak yoksay
            Dim driveType, diskSize
            ' Disk boyutunu almayı deneyelim
            diskSize = objDrive.Size
            If Err.Number = 0 Then
                ' Boyut verisi mevcutsa
                If objDrive.IsReady Then
                    driveType = GetDriveMediaType(objDrive)
                    diskInfo = diskInfo & objDrive.DeviceID & " - " & driveType & " - Kapasite: " & _
                               FormatNumber(diskSize / 1024 / 1024 / 1024, 2) & " GB" & vbCrLf
                Else
                    diskInfo = diskInfo & objDrive.DeviceID & " - " & "Disk hazır değil" & vbCrLf
                End If
            Else
                diskInfo = diskInfo & objDrive.DeviceID & " - " & "Disk boyutu alınamıyor" & vbCrLf
            End If
            On Error GoTo 0 ' Hata kontrolünü sıfırla
        End If
    Next
Else
    diskInfo = diskInfo & "Disk bilgisi bulunamadı." & vbCrLf
End If

' Disk türünü belirleme fonksiyonu (HDD/SSD/NVMe/USB)
Function GetDriveMediaType(DiskDrive)
    On Error Resume Next
    Dim mediaType
    If Not DiskDrive Is Nothing Then
        ' NVMe diskleri model adı üzerinden tanıyacağız
        If InStr(1, LCase(DiskDrive.Model), "nvme") > 0 Or InStr(1, LCase(DiskDrive.Model), "nvm") > 0 Then
            mediaType = "NVMe"
        ' SSD'yi model adında "ssd" veya "sd" geçiyorsa tanıyacağız
        ElseIf InStr(1, LCase(DiskDrive.Model), "ssd") > 0 Or InStr(1, LCase(DiskDrive.Model), "sd") > 0 Then
            mediaType = "SSD"
        ' USB diskleri model adında "usb" geçiyorsa tanıyacağız
        ElseIf InStr(1, LCase(DiskDrive.InterfaceType), "usb") > 0 Then
            mediaType = "USB"
        ' Diğer diskler için HDD
        Else
            mediaType = "HDD"
        End If
    Else
        mediaType = "Bilinmiyor"
    End If
    On Error GoTo 0
    GetDriveMediaType = mediaType
End Function

' SystemSet koleksiyonunu alalım
Set SystemSet = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")

' Bilgileri derle (Sistem, İşlemci ve Anakart için yalnızca bir döngü)
For Each System in SystemSet
    For Each objProcessor in colProcessors
        For Each bbType In colMB
            MbVendor = bbType.Manufacturer
            MbModel = bbType.Product
            tMessage = "İşletim Sistemi		: " & System.Caption & vbNewLine & _
                       "İşletim Sistemi Versiyonu	: " & System.Version & vbNewLine & _
                       "Windows Mimari Yapısı	: " & strArchitecture & vbNewLine & _
                       "Kullanıcı Adı		: " & objNetwork.UserName & vbNewLine & _
                       "Bilgisayar Adı		: " & strComputerName & vbNewLine & _
                       "Son Format Tarihi		: " & fthx & vbNewLine & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       "Anakart Üreticisi		: " & MbVendor & vbNewLine & _
                       "Anakart Modeli		: " & MbModel & vbNewLine & _
                       "İşlemci			: " & objProcessor.Manufacturer & vbNewLine & _
                       "İşlemci Modeli		: " & objProcessor.Name & vbNewLine & _
                       "CPU Mimarisi		: " & strArchitecture & vbNewLine & _
                       "Toplam RAM		: " & Int(TotalRam / 1024) & " GB" & vbNewLine & _
                       "RAM Yuvaları		: " & vbNewLine & ramDetails & vbNewLine & _
                       "Grafik Kart(lar)ı 		: " & vbNewLine & tStr & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       "Ağ Kart(lar)ı ve IP Adres(ler)i	:" & vbNewLine & vbNewLine & myIPAddresses & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       diskInfo

            Exit For ' Bu döngüyü her bölüm için yalnızca bir kez çalıştırın
        Next
        Exit For
    Next
    Exit For
Next

' WshShell.Popup ile görüntüle
Set WshShell = CreateObject("WScript.Shell")
WshShell.Popup tMessage, 0, "Donanım Bilgileri | by Abdullah ERTÜRK", 4096
