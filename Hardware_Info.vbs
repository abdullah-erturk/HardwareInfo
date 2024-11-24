Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2") ' More user-friendly WMI connection
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

' Total RAM calculation
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

    ramDetails = ramDetails & "Slot " & i & ": " & Int(memTmp1 / 1024) & " GB, Hýz: " & obj2.Speed & " MHz, Tür: " & ramType & vbCrLf
    i = i + 1
Next

' Processor architecture
Set colItems = objWMIService.ExecQuery("Select Architecture from Win32_Processor")
For Each objItem in colItems
    If objItem.Architecture = 0 Then
       strArchitecture = "x86"
    ElseIf objItem.Architecture = 9 Then
       strArchitecture = "x64"
    End If
Next

' Graphics Card information
On Error Resume Next
Set colItemsx = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
Dim tStr, tStr2
tStr = ""
tStr2 = ""

' Get display card information
For Each objItem in colItemsx
    tStr = tStr & "Model    : " & objItem.Description & vbCrLf
    
    ' Get memory size
    Dim memSize
    memSize = objItem.AdapterRAM / 1024 / 1024 ' in MB

    ' Check the memory size
    If InStr(LCase(objItem.Description), "intel") > 0 Then
        ' If internal graphics, we can dynamically take the memory from system RAM
        If memSize < 128 Then
            memSize = 128 ' Default memory for integrated graphics is 128 MB
        End If
    Else
        ' For external graphics, check the AdapterRAM value
        If memSize < 1024 Then
            ' If memory is less than 1 GB, probably an incorrect value is returned
            memSize = 4096 ' Default memory for external graphics is 4096 MB (4 GB)
        End If
    End If

    ' Memory information is incorrect, hence commenting out :)
    ' tStr = tStr & "Memory: " & memSize & " MB" & vbCrLf
Next
On Error GoTo 0

' Network card information
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
            ipAddr = "Not found"
        End If

        If Not IsNull(objAdapter.DHCPServer) Then
            dhcpServer = objAdapter.DHCPServer
        Else
            dhcpServer = "Not found"
        End If

        myIPAddresses = myIPAddresses & "Network Adapter " & counter & "" & vbCrLf & _
                        "Description	: " & description & vbCrLf & _
                        "MAC Address	: " & macAddr & vbCrLf & _
                        "IP Address	: " & ipAddr & vbCrLf & _
                        "DHCP Server	: " & dhcpServer & vbCrLf & vbCrLf

        counter = counter + 1
    End If
Next

' Get WAN IP using external web service
Dim WANIP
On Error Resume Next ' Error handling
Dim objXMLHttp
Set objXMLHttp = CreateObject("MSXML2.XMLHTTP")
objXMLHttp.Open "GET", "http://api.ipify.org", False
objXMLHttp.Send

' Internet connection check
If Err.Number <> 0 Then
    WANIP = "Not found" ' Set WAN IP to "Not found" if no internet
Else
    WANIP = objXMLHttp.responseText ' If connected to the internet, get WAN IP
End If
On Error GoTo 0 ' Reset error handling

' Ping to check internet connection
Dim pingResult
pingResult = PingHost("8.8.8.8")

If pingResult Then
    ' If ping to 8.8.8.8 is successful, show WAN IP
    If WANIP = "Not found" Then
        WANIP = objXMLHttp.responseText ' Get actual WAN IP
    End If
Else
    ' If ping to 8.8.8.8 fails, set WAN IP to "Not found"
    WANIP = "Not found"
End If

myIPAddresses = myIPAddresses & "WAN IP Address	: " & WANIP & vbCrLf

' Add DNS server addresses
Set colNicConfigs = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
Dim dnsFound
dnsFound = False ' Flag to track if DNS is found

For Each objNicConfig In colNicConfigs
    If Not IsNull(objNicConfig.DNSServerSearchOrder) Then
        myIPAddresses = myIPAddresses & "DNS Server	: " 
        Dim dnsList
        dnsList = ""
        
        For Each dnsServer In objNicConfig.DNSServerSearchOrder
            If dnsList = "" Then
                dnsList = dnsServer ' First DNS server
            Else
                dnsList = dnsList & " / " & dnsServer ' Append additional DNS servers
            End If
        Next
        
        myIPAddresses = myIPAddresses & dnsList & vbCrLf
        dnsFound = True ' Mark that DNS was found
        Exit For
    End If
Next

If Not dnsFound Then
    myIPAddresses = myIPAddresses & "DNS Server: Not found" & vbCrLf
End If

' Ping function
Function PingHost(host)
    Dim objShell, command, result
    Set objShell = CreateObject("WScript.Shell")
    command = "ping -n 1 " & host ' "-n 1" sends a single ping request
    result = objShell.Run(command, 0, True) ' Run command and get the result
    If result = 0 Then
        PingHost = True ' Successful ping
    Else
        PingHost = False ' Unsuccessful ping
    End If
End Function

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

' Get disk information and check types
Dim diskInfo
diskInfo = "Disk Summary: " & vbCrLf

' If colDrives is not empty, proceed with processing
If colDrives.Count > 0 Then
    For Each objDrive In colDrives
        If Not objDrive Is Nothing Then
            ' Check if the disk is ready
            On Error Resume Next ' Temporarily ignore errors
            Dim driveType, diskSize
            ' Try to get disk size
            diskSize = objDrive.Size
            If Err.Number = 0 Then
                ' If size data is available
                If objDrive.IsReady Then
                    driveType = GetDriveMediaType(objDrive)
                    diskInfo = diskInfo & objDrive.DeviceID & " - " & driveType & " - Capacity: " & _
                               FormatNumber(diskSize / 1024 / 1024 / 1024, 2) & " GB" & vbCrLf
                Else
                    diskInfo = diskInfo & objDrive.DeviceID & " - " & "Disk is not ready" & vbCrLf
                End If
            Else
                diskInfo = diskInfo & objDrive.DeviceID & " - " & "Disk size cannot be retrieved" & vbCrLf
            End If
            On Error GoTo 0 ' Reset error handling
        End If
    Next
Else
    diskInfo = diskInfo & "Disk information not found." & vbCrLf
End If

' Function to determine the disk type (HDD/SSD/NVMe/USB)
Function GetDriveMediaType(DiskDrive)
    On Error Resume Next
    Dim mediaType
    If Not DiskDrive Is Nothing Then
        ' Identify NVMe disks by model name
        If InStr(1, LCase(DiskDrive.Model), "nvme") > 0 Or InStr(1, LCase(DiskDrive.Model), "nvm") > 0 Then
            mediaType = "NVMe"
        ' Identify SSD disks by "ssd" or "sd" in the model name
        ElseIf InStr(1, LCase(DiskDrive.Model), "ssd") > 0 Or InStr(1, LCase(DiskDrive.Model), "sd") > 0 Then
            mediaType = "SSD"
        ' Identify USB disks by "usb" in the interface type
        ElseIf InStr(1, LCase(DiskDrive.InterfaceType), "usb") > 0 Then
            mediaType = "USB"
        ' Default to HDD for other disks
        Else
            mediaType = "HDD"
        End If
    Else
        mediaType = "Unknown"
    End If
    On Error GoTo 0
    GetDriveMediaType = mediaType
End Function

' Get the system collection
Set SystemSet = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")

' Gather information (only one loop for system, processor, and motherboard)
For Each System in SystemSet
    For Each objProcessor in colProcessors
        For Each bbType In colMB
            MbVendor = bbType.Manufacturer
            MbModel = bbType.Product
            tMessage = "Operating System		: " & System.Caption & vbNewLine & _
                       "OS Version		: " & System.Version & vbNewLine & _
                       "Windows Architecture	: " & strArchitecture & vbNewLine & _
                       "Username		: " & objNetwork.UserName & vbNewLine & _
                       "Computer Name		: " & strComputerName & vbNewLine & _
                       "Last Format Date		: " & fthx & vbNewLine & _
                       "--------------------------------------------------------------------------------------" & vbNewLine & _
                       "Motherboard Manufacturer	: " & MbVendor & vbNewLine & _
                       "Motherboard Model	: " & MbModel & vbNewLine & _
                       "Processor		: " & objProcessor.Manufacturer & vbNewLine & _
                       "Processor Model		: " & objProcessor.Name & vbNewLine & _
                       "CPU Architecture		: " & strArchitecture & vbNewLine & _
                       "Total RAM		: " & Int(TotalRam / 1024) & " GB" & vbNewLine & _
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

' Display using WshShell.Popup
Set WshShell = CreateObject("WScript.Shell")
WshShell.Popup tMessage, 0, "Hardware Information | by Abdullah ERTÜRK", 4096
