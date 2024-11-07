Set SystemSet = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
strOSArch = GetObject("winmgmts:root\cimv2:Win32_OperatingSystem=@").OSArchitecture
Set objNetwork = CreateObject("Wscript.Network")
Set wshShell = CreateObject("WScript.Shell")
strComputerName = wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
Set oShell = WScript.CreateObject("WScript.Shell")
proc_arch = oShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
Set oEnv = oShell.Environment("SYSTEM")

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colMB = objWMIService.ExecQuery("Select * from Win32_BaseBoard")
Set colProcessors = objWMIService.ExecQuery("Select * from Win32_Processor")
Set colDrives = objWMIService.ExecQuery("Select * from Win32_DiskDrive")

' Toplam RAM hesaplama
Set obj = GetObject("winmgmts:").InstancesOf("Win32_PhysicalMemory")
TotalRam = 0
ramDetails = ""
i = 1

For Each obj2 In obj
    memTmp1 = obj2.capacity / 1024 / 1024
    TotalRam = TotalRam + memTmp1
    ramSpeed = obj2.Speed
    ramType = ""

    ' RAM türünü belirleme (DDR genellikle hızla ilişkilendirilir)
    If ramSpeed >= 1600 And ramSpeed < 2133 Then
        ramType = "DDR3"
    ElseIf ramSpeed >= 2133 And ramSpeed < 2933 Then
        ramType = "DDR4"
    ElseIf ramSpeed >= 2933 Then
        ramType = "DDR5"
    Else
        ramType = "Unknown"
    End If

    ramDetails = ramDetails & "Slot " & i & ": " & FormatNumber(memTmp1 / 1024, 2) & " GB, Hız: " & obj2.Speed & " MHz, Tür: " & ramType & vbCrLf
    i = i + 1
Next

' İşlemci mimarisi belirleme
Set colItems = objWMIService.ExecQuery("Select Architecture from Win32_Processor")
For Each objItem in colItems
    If objItem.Architecture = 0 Then
       strArchitecture = "x86"
    ElseIf objItem.Architecture = 9 Then
       strArchitecture = "x64"
    End If
Next

' Grafik kartı bilgisi
On Error Resume Next
Set colItemsx = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
Dim tStr, tStr2
tStr = ""
tStr2 = ""

For Each objItem in colItemsx
    tStr = tStr & objItem.Description & " " & "Ram " & FormatNumber(objItem.AdapterRAM / 1024 / 1024, 0) & " MB" & vbCrLf
    tStr2 = tStr2 & objItem.Description & " " & "Sürücü versiyonu: " & objItem.DriverVersion & vbCrLf
Next

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

' Disk bilgileri
Dim diskInfo
diskInfo = "Disk Özeti	:" & vbCrLf

' Disk bilgilerini al ve türlerini kontrol et
For Each objDrive In colDrives
    If objDrive.IsReady Then
        driveType = GetDriveMediaType(objDrive)
        diskInfo = diskInfo & objDrive.DeviceID & " - " & driveType & " - Kapasite: " & _
                   FormatNumber(objDrive.Size / 1024 / 1024 / 1024, 2) & " GB" & vbCrLf
    Else
        diskInfo = diskInfo & objDrive.DeviceID & " - " & "Disk hazır değil" & vbCrLf
    End If
Next

' Disk türünü belirleme fonksiyonu (HDD/SSD/USB)
Function GetDriveMediaType(DiskDrive)
    If InStr(1, LCase(DiskDrive.Model), "ssd") > 0 Or InStr(1, LCase(DiskDrive.Model), "sd") > 0 Then
        GetDriveMediaType = "SSD"
    ElseIf InStr(1, LCase(DiskDrive.InterfaceType), "usb") > 0 Then
        GetDriveMediaType = "USB"
    Else
        GetDriveMediaType = "HDD"
    End If
End Function

' Ağ bilgileri
Dim myIPAddresses : myIPAddresses = ""
Dim counter : counter = 1
Dim colAdapters : Set colAdapters = objWMIService.ExecQuery("Select IPAddress, Description, MACAddress from Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objAdapter in colAdapters
  If Not IsNull(objAdapter.IPAddress) Then
    myIPAddresses = myIPAddresses & "Ağ Kartı " & counter & ":" & vbCrLf & _
                    objAdapter.Description & " : " & vbCrLf & _
                    "IP Adresi: " & Trim(objAdapter.IPAddress(0)) & vbCrLf & _
                    "MAC Adresi: " & objAdapter.MACAddress & vbCrLf & _
                    "" & vbCrLf
    counter = counter + 1
  End If
Next

' WAN IP adresini almak için dış bir web servisi kullan
Dim objXMLHttp
Set objXMLHttp = CreateObject("MSXML2.XMLHTTP")
objXMLHttp.Open "GET", "http://api.ipify.org", False
objXMLHttp.Send

Dim WANIP
WANIP = objXMLHttp.responseText

myIPAddresses = myIPAddresses & "WAN IP Adresi: " & WANIP & vbCrLf

' Bilgileri derle (Sistem, İşlemci ve Anakart için yalnızca bir döngü)
For Each System in SystemSet
    For Each objProcessor in colProcessors
        For Each bbType In colMB
            MbVendor = bbType.Manufacturer
            MbModel = bbType.Product
            tMessage = "Ýþletim Sistemi		: " & System.Caption & vbNewLine & _
                       "Ýþletim Sistemi Versiyonu	: " & System.Version & vbNewLine & _
                       "Windows Mimari Yapýsý	: " & strOSArch & vbNewLine & _
                       "Kullanýcý Adý		: " & objNetwork.UserName & vbNewLine & _
                       "Bilgisayar Adý		: " & strComputerName & vbNewLine & _
                       "Son Format Tarihi		: " & fthx & vbNewLine & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       "Anakart Üreticisi		: " & MbVendor & vbNewLine & _
                       "Anakart Modeli		: " & MbModel & vbNewLine & _
                       "Ýþlemci			: " & objProcessor.Manufacturer & vbNewLine & _
                       "Ýþlemci Modeli		: " & objProcessor.Name & vbNewLine & _
                       "CPU Mimarisi		: " & strArchitecture & vbNewLine & _
                       "Toplam RAM		: " & FormatNumber(TotalRam / 1024, 2) & " GB" & vbNewLine & _
                       "RAM Yuvalarý		: " & vbNewLine & ramDetails & vbNewLine & _
                       "Grafik Kart(lar)ý		: " & vbNewLine & tStr & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       "Að Kart(lar)ý ve IP Adres(ler)i	:" & vbNewLine & vbNewLine & myIPAddresses & _
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
WshShell.Popup tMessage, 0, "Donaným Bilgileri | by Abdullah ERTÜRK", 4096
