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

Set obj = GetObject("winmgmts:").InstancesOf("Win32_PhysicalMemory")
TotalRam = 0
ramDetails = ""
i = 1

For Each obj2 In obj
    memTmp1 = obj2.capacity / 1024 / 1024
    TotalRam = TotalRam + memTmp1
    ramSpeed = obj2.Speed
    ramType = ""

    ' Determining the type of RAM (DDR is usually associated with speed)
    If ramSpeed >= 1600 And ramSpeed < 2133 Then
        ramType = "DDR3"
    ElseIf ramSpeed >= 2133 And ramSpeed < 2933 Then
        ramType = "DDR4"
    ElseIf ramSpeed >= 2933 Then
        ramType = "DDR5"
    Else
        ramType = "Unknown"
    End If

    ramDetails = ramDetails & "Slot " & i & ": " & FormatNumber(memTmp1 / 1024, 2) & " GB, Speed: " & obj2.Speed & " MHz, Type: " & ramType & vbCrLf
    i = i + 1
Next

' Determining processor architecture
Set colItems = objWMIService.ExecQuery("Select Architecture from Win32_Processor")
For Each objItem in colItems
    If objItem.Architecture = 0 Then
       strArchitecture = "x86"
    ElseIf objItem.Architecture = 9 Then
       strArchitecture = "x64"
    End If
Next

' Graphics card information
On Error Resume Next
Set colItemsx = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
Dim tStr, tStr2
tStr = ""
tStr2 = ""

For Each objItem in colItemsx
    tStr = tStr & objItem.Description & " " & "Ram " & FormatNumber(objItem.AdapterRAM / 1024 / 1024, 0) & " MB" & vbCrLf
    tStr2 = tStr2 & objItem.Description & " " & "Driver version: " & objItem.DriverVersion & vbCrLf
Next

' Operating system information
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

' Disk information
Dim diskInfo
diskInfo = "Disk Summary Information	:" & vbCrLf

' Get disk information and check types
For Each objDrive In colDrives
    If objDrive.IsReady Then
        driveType = GetDriveMediaType(objDrive) ' HDD/SSD/USB detection
        diskInfo = diskInfo & objDrive.DeviceID & " - " & driveType & " - Capacity: " & _
                   FormatNumber(objDrive.Size / 1024 / 1024 / 1024, 2) & " GB" & vbCrLf
    Else
        diskInfo = diskInfo & objDrive.DeviceID & " - " & "Disk not ready" & vbCrLf
    End If
Next

' Function to determine disk type (HDD/SSD/USB)
Function GetDriveMediaType(DiskDrive)
    If InStr(1, LCase(DiskDrive.Model), "ssd") > 0 Or InStr(1, LCase(DiskDrive.Model), "sd") > 0 Then
        GetDriveMediaType = "SSD"
    ElseIf InStr(1, LCase(DiskDrive.InterfaceType), "usb") > 0 Then
        GetDriveMediaType = "USB"
    Else
        GetDriveMediaType = "HDD"
    End If
End Function

' Network information
Dim myIPAddresses : myIPAddresses = ""
Dim counter : counter = 1
Dim colAdapters : Set colAdapters = objWMIService.ExecQuery("Select IPAddress, Description, MACAddress from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
Dim objAdapter

' Get LAN IP addresses and MAC addresses
For Each objAdapter in colAdapters
  If Not IsNull(objAdapter.IPAddress) Then
    myIPAddresses = myIPAddresses & "Network Adapter " & counter & ":" & vbCrLf & _
                    objAdapter.Description & " : " & vbCrLf & _
                    "IP Address: " & Trim(objAdapter.IPAddress(0)) & vbCrLf & _
                    "MAC Address: " & objAdapter.MACAddress & vbCrLf & _
                    "" & vbCrLf
    counter = counter + 1 
  End If
Next

' Use an external web service to obtain WAN IP address
Dim objXMLHttp
Set objXMLHttp = CreateObject("MSXML2.XMLHTTP")
objXMLHttp.Open "GET", "http://api.ipify.org", False
objXMLHttp.Send

Dim WANIP
WANIP = objXMLHttp.responseText

myIPAddresses = myIPAddresses & "WAN IP Address: " & WANIP & vbCrLf

' Compile information (Only one loop for System, Processor, and Motherboard)
For Each System in SystemSet
    For Each objProcessor in colProcessors
        For Each bbType In colMB
            MbVendor = bbType.Manufacturer
            MbModel = bbType.Product
            tMessage = "Operating System		: " & System.Caption & vbNewLine & _
                       "OS Version		: " & System.Version & vbNewLine & _
                       "Windows Architecture	: " & strOSArch & vbNewLine & _
                       "Username		: " & objNetwork.UserName & vbNewLine & _
                       "Computer Name		: " & strComputerName & vbNewLine & _
                       "Last Format Date		: " & fthx & vbNewLine & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       "Motherboard Manufacturer	: " & MbVendor & vbNewLine & _
                       "Motherboard Model	: " & MbModel & vbNewLine & _
                       "Processor		: " & objProcessor.Manufacturer & vbNewLine & _
                       "Processor Model		: " & objProcessor.Name & vbNewLine & _
                       "CPU Architecture		: " & strArchitecture & vbNewLine & _
                       "Total RAM		: " & FormatNumber(TotalRam / 1024, 2) & " GB" & vbNewLine & _
                       "RAM Slots		: " & vbNewLine & ramDetails & vbNewLine & _
                       "Graphics Card(s)		: " & vbNewLine & tStr & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       "Network Adapter(s) and IP Address(es): " & vbNewLine & "" & vbNewLine & myIPAddresses & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       diskInfo
            Exit For ' Only run this loop once for each section
        Next
        Exit For
    Next
    Exit For
Next

' Display with WshShell.Popup
Set WshShell = CreateObject("WScript.Shell")
WshShell.Popup tMessage, 0, "Hardware Information | by Abdullah ERTÜRK", 4096
