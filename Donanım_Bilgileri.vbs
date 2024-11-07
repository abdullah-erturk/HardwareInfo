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

    ' RAM t�r�n� belirleme (DDR genellikle h�zla ili�kilendirilir)
    If ramSpeed >= 1600 And ramSpeed < 2133 Then
        ramType = "DDR3"
    ElseIf ramSpeed >= 2133 And ramSpeed < 2933 Then
        ramType = "DDR4"
    ElseIf ramSpeed >= 2933 Then
        ramType = "DDR5"
    Else
        ramType = "Unknown"
    End If

    ramDetails = ramDetails & "Slot " & i & ": " & FormatNumber(memTmp1 / 1024, 2) & " GB, H�z: " & obj2.Speed & " MHz, T�r: " & ramType & vbCrLf
    i = i + 1
Next

' ��lemci mimarisi belirleme
Set colItems = objWMIService.ExecQuery("Select Architecture from Win32_Processor")
For Each objItem in colItems
    If objItem.Architecture = 0 Then
       strArchitecture = "x86"
    ElseIf objItem.Architecture = 9 Then
       strArchitecture = "x64"
    End If
Next

' Grafik kart� bilgisi
On Error Resume Next
Set colItemsx = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
Dim tStr, tStr2
tStr = ""
tStr2 = ""

For Each objItem in colItemsx
    tStr = tStr & objItem.Description & " " & "Ram " & FormatNumber(objItem.AdapterRAM / 1024 / 1024, 0) & " MB" & vbCrLf
    tStr2 = tStr2 & objItem.Description & " " & "S�r�c� versiyonu: " & objItem.DriverVersion & vbCrLf
Next

' ��letim sistemi bilgisi
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
diskInfo = "Disk �zeti	:" & vbCrLf

' Disk bilgilerini al ve t�rlerini kontrol et
For Each objDrive In colDrives
    If objDrive.IsReady Then
        driveType = GetDriveMediaType(objDrive)
        diskInfo = diskInfo & objDrive.DeviceID & " - " & driveType & " - Kapasite: " & _
                   FormatNumber(objDrive.Size / 1024 / 1024 / 1024, 2) & " GB" & vbCrLf
    Else
        diskInfo = diskInfo & objDrive.DeviceID & " - " & "Disk haz�r de�il" & vbCrLf
    End If
Next

' Disk t�r�n� belirleme fonksiyonu (HDD/SSD/USB)
Function GetDriveMediaType(DiskDrive)
    If InStr(1, LCase(DiskDrive.Model), "ssd") > 0 Or InStr(1, LCase(DiskDrive.Model), "sd") > 0 Then
        GetDriveMediaType = "SSD"
    ElseIf InStr(1, LCase(DiskDrive.InterfaceType), "usb") > 0 Then
        GetDriveMediaType = "USB"
    Else
        GetDriveMediaType = "HDD"
    End If
End Function

' A� bilgileri
Dim myIPAddresses : myIPAddresses = ""
Dim counter : counter = 1
Dim colAdapters : Set colAdapters = objWMIService.ExecQuery("Select IPAddress, Description, MACAddress from Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objAdapter in colAdapters
  If Not IsNull(objAdapter.IPAddress) Then
    myIPAddresses = myIPAddresses & "A� Kart� " & counter & ":" & vbCrLf & _
                    objAdapter.Description & " : " & vbCrLf & _
                    "IP Adresi: " & Trim(objAdapter.IPAddress(0)) & vbCrLf & _
                    "MAC Adresi: " & objAdapter.MACAddress & vbCrLf & _
                    "" & vbCrLf
    counter = counter + 1
  End If
Next

' WAN IP adresini almak i�in d�� bir web servisi kullan
Dim objXMLHttp
Set objXMLHttp = CreateObject("MSXML2.XMLHTTP")
objXMLHttp.Open "GET", "http://api.ipify.org", False
objXMLHttp.Send

Dim WANIP
WANIP = objXMLHttp.responseText

myIPAddresses = myIPAddresses & "WAN IP Adresi: " & WANIP & vbCrLf

' Bilgileri derle (Sistem, ��lemci ve Anakart i�in yaln�zca bir d�ng�)
For Each System in SystemSet
    For Each objProcessor in colProcessors
        For Each bbType In colMB
            MbVendor = bbType.Manufacturer
            MbModel = bbType.Product
            tMessage = "��letim Sistemi		: " & System.Caption & vbNewLine & _
                       "��letim Sistemi Versiyonu	: " & System.Version & vbNewLine & _
                       "Windows Mimari Yap�s�	: " & strOSArch & vbNewLine & _
                       "Kullan�c� Ad�		: " & objNetwork.UserName & vbNewLine & _
                       "Bilgisayar Ad�		: " & strComputerName & vbNewLine & _
                       "Son Format Tarihi		: " & fthx & vbNewLine & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       "Anakart �reticisi		: " & MbVendor & vbNewLine & _
                       "Anakart Modeli		: " & MbModel & vbNewLine & _
                       "��lemci			: " & objProcessor.Manufacturer & vbNewLine & _
                       "��lemci Modeli		: " & objProcessor.Name & vbNewLine & _
                       "CPU Mimarisi		: " & strArchitecture & vbNewLine & _
                       "Toplam RAM		: " & FormatNumber(TotalRam / 1024, 2) & " GB" & vbNewLine & _
                       "RAM Yuvalar�		: " & vbNewLine & ramDetails & vbNewLine & _
                       "Grafik Kart(lar)�		: " & vbNewLine & tStr & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       "A� Kart(lar)� ve IP Adres(ler)i	:" & vbNewLine & vbNewLine & myIPAddresses & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       diskInfo
            Exit For ' Bu d�ng�y� her b�l�m i�in yaln�zca bir kez �al��t�r�n
        Next
        Exit For
    Next
    Exit For
Next

' WshShell.Popup ile g�r�nt�le
Set WshShell = CreateObject("WScript.Shell")
WshShell.Popup tMessage, 0, "Donan�m Bilgileri | by Abdullah ERT�RK", 4096