'====================================================================
'# __ C:\Windows\System32\gatherNetworkInfo.vbs
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 33-GatherNetworkInfo.vbs
'# __ SYNCED BY THIS CODE \VBS 32-COPIER_SYNC.VBS
'# __ VBS 33-GatherNetworkInfo.vbs
'# __ 
'# BY Matthew Lancaster __ Matt.Lan@Btinternet.com
'# BY MICROSOFT ONE
'# __ 
'# START     TIME [ --- 10-Jun-2009 21:36:00 ]
'# __ 
'====================================================================
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
' -------------------------------------------------------------------
' A MICROSOFT CODE FOUND IN LOCATION
' WINDOWS 7 PRO
' WINDOWS\SYSTEM32
' FOUND BY SCHEDULER 
' FOUND NEVER TO BE RUN AND DISABLED ONLY ACCOUNT LOGS NOTHING
' DOES NOT TALK WHAT ANY START SITUATION MAYBE
' A GOOD INFO GATHERING EXAMPLE ABOUT NETWORK CARD INFO
' FILE CREATION 10-Jun-2009 21:36:00
' LOCATION PROPETYS OF AUTORUNS
' C:\windows\sysnative\GatherNetworkInfo.vbs
' & 
' C:\windows\system32\GatherNetworkInfo.vbs
' LOCATION IN SCHEDULER 
' \Microsoft\Windows\NetTrace
' SCHEDULER DESCRIPTION
' Network information collector
' -------------------------------------------------------------------

'#-------------------------------------------------------------------
' -- 001 SESSION __
' -------------------------------------------------------------------
'#-------------------------------------------------------------------
' __ FROM  Sun 07-Oct-2018 06:36:17
' __ TO    Sun 07-Oct-2018 07:38:00 _ MIXED WITH OTHER TIDY HEADER INFO
'#-------------------------------------------------------------------

Dim FSO, shell, xslProcessor

Sub RunCmd(CommandString, OutputFile)
	cmd = "cmd /c " + CommandString + " >> " + OutputFile
	shell.Run cmd, 0, True
End Sub

Sub GetOSInfo(outputFileName)
    On Error Resume Next
    strComputer = "."
    HKEY_LOCAL_MACHINE = &H80000002

    Dim objReg, outputFile
    Dim buildDetailNames, buildDetailRegValNames

    buildDetailNames = Array("Product Name", "Version", "Build Lab", "Type")
    buildDetailRegValNames = Array("ProductName", "CurrentVersion", "BuildLabEx", "CurrentType")

    Set outputFile = FSO.OpenTextFile(outputFileName, 2, True)

    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
                     strComputer & "\root\default:StdRegProv")

    outputFile.WriteLine("[Architecture/Processor Information]")
    outputFile.WriteLine()
    outputFile.Close
    cmd = "cmd /c set processor >> " & outputFileName
    shell.Run cmd, 0, True

    Set outputFile = FSO.OpenTextFile(outputFileName, 8, True)

    outputFile.WriteLine()
    outputFile.WriteLine("[Operating System Information]")
    outputFile.WriteLine()

    strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    for I = 0 to UBound(buildDetailNames)
        objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, buildDetailRegValNames(I), info
        outputFile.WriteLine(buildDetailNames(I) + " = " + info)
    Next

    outputFile.WriteLine()
    strKeyPath = "SYSTEM\SETUP"
    objReg.GetDWordValue HKEY_LOCAL_MACHINE, strKeyPath, "Upgrade", upgradeInfo
    if IsNull(upgradeInfo) Then
        outputFile.WriteLine("This is a clean installed system")
    Else
        outputFile.WriteLine("This is an upgraded system")
    End If

    outputFile.WriteLine(buildDetailNames(I) + " = " + info)

    outputFile.WriteLine()
    outputFile.WriteLine("[File versions]")
    outputFile.WriteLine()

    Set shell = WScript.CreateObject( "WScript.Shell" )
    windir = shell.ExpandEnvironmentStrings("%windir%\system32\")

    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

    Dim FileSet
    FileSet = Array("onex.dll", "l2nacp.dll", "wlanapi.dll", "wlancfg.dll", "wlanconn.dll", "wlandlg.dll", "wlanext.exe", "wlangpui.dll", "wlanhc.dll", "wlanhlp.dll", "wlaninst.dll", "wlanmm.dll", "wlanmmhc.dll", "wlanmsm.dll", "wlanpref.dll", "wlansec.dll", "wlansvc.dll", "wlanui.dll")

    For Each file in FileSet
        filename = windir + file
        strQuery = "Select * from CIM_Datafile Where Name = '" + Replace(filename, "\", "\\") + "'"
        Set fileProp = objWMIService.ExecQuery _
            (strQuery)

        For Each objFile in fileProp
            outputFile.WriteLine(file + "    " + objFile.Version)
        Next
    Next

    Dim Dot3FileSet
    Dot3FileSet = Array("onex.dll", "dot3api.dll", "dot3cfg.dll", "dot3dlg.dll", "dot3gpclnt.dll", "dot3gpui.dll", "dot3msm.dll", "dot3svc.dll", "dot3ui.dll")

    For Each file in Dot3FileSet
        filename = windir + file
        strQuery = "Select * from CIM_Datafile Where Name = '" + Replace(filename, "\", "\\") + "'"
        Set fileProp = objWMIService.ExecQuery _
            (strQuery)

        For Each objFile in fileProp
            outputFile.WriteLine(file + "    " + objFile.Version)
        Next
    Next

    call GetBatteryInfo(outputFile)
    outputFile.Close

    Set outputFile = FSO.OpenTextFile(outputFileName, 8, True)
    outputFile.WriteLine("")
    outputFile.WriteLine("[System Information]")
    outputFile.WriteLine("")
    outputFile.Close

    'Comments: Dumping System Information using "systeminfo" command

    cmd = "cmd /c systeminfo >> " & outputFileName
    shell.Run cmd, 0, True

    Set outputFile = FSO.OpenTextFile(outputFileName, 8, True)
    outputFile.WriteLine("")
    outputFile.WriteLine("[User Information]")
    outputFile.WriteLine("")
    outputFile.Close

    cmd = "cmd /c set u >> " & outputFileName
    shell.Run cmd, 0, True

End Sub

Sub GetBatteryInfo(outputFile)
    On Error Resume Next
    strComputer = "."
    outputFile.WriteLine()
    outputFile.WriteLine("[Power Information]")
    outputFile.WriteLine()
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Battery")
    if colItems.Count = 0 Then
	outputFile.WriteLine("It is a Desktop running on AC")
    Else
    For Each objItem in colItems
        if objItem.Availability = 2 Then
	    outputFile.WriteLine("Machine is running on AC Adapter")
        Else
	    if objitem.Availability = 3 Then
		outputFile.WriteLine("Machine is running on Battery")
	    End If
        End If
    Next
    End If
End Sub



Sub GetWcnInfo(outputFileName)
    On Error Resume Next
    Dim WcnInfoFile

    Set WcnInfoFile= FSO.OpenTextFile(outputFileName, 8, True)
    WcnInfoFile.WriteLine("-------------------------------------")
    WcnInfoFile.WriteLine("---------+ WCN Information +---------")     
    WcnInfoFile.WriteLine("-------------------------------------")    
    WcnInfoFile.WriteLine("")
    WcnInfoFile.WriteLine("")
    WcnInfoFile.WriteLine("-----------------")
    WcnInfoFile.WriteLine("+ Services Status")
    WcnInfoFile.WriteLine("-----------------")
    WcnInfoFile.WriteLine("")
    WcnInfoFile.Close

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    cmd = "cmd /c sc query wcncsvc  >> " & outputFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c sc query wlansvc  >> " & outputFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c sc query eaphost  >> " & outputFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c sc query fdrespub  >> " & outputFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c sc query upnphost   >> " & outputFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c sc query eaphost  >> " & outputFileName
    objShell.Run cmd, 0, True


    Set WcnInfoFile= FSO.OpenTextFile(outputFileName, 8, True)
    WcnInfoFile.WriteLine("")
    WcnInfoFile.WriteLine("")
    WcnInfoFile.WriteLine("-----------------------")
    WcnInfoFile.WriteLine("+ WCN Files Information ")
    WcnInfoFile.WriteLine("-----------------------")
    WcnInfoFile.WriteLine("")

    strComputer = "."

    Set shell = WScript.CreateObject( "WScript.Shell" )
    windir = shell.ExpandEnvironmentStrings("%windir%\system32\")

    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

    Dim FileSet
    FileSet = Array("wcncsvc.dll", "wcnapi.dll", "fdwcn.dll", "wcneapauthproxy.dll", "wcneappeerproxy.dll", "wcnwiz.dll", "wcnnetsh.dll", "wczdlg.dll")

    For Each file in FileSet
        filename = windir + file
        strQuery = "Select * from CIM_Datafile Where Name = '" + Replace(filename, "\", "\\") + "'"
        Set fileProp = objWMIService.ExecQuery _
            (strQuery)

        For Each objFile in fileProp
		WcnInfoFile.WriteLine("")
		WcnInfoFile.WriteLine("---------------------")
		WcnInfoFile.WriteLine(file)
		WcnInfoFile.WriteLine("---------------------")
		WcnInfoFile.WriteLine("	- Version 		:  	" + objFile.Version )
		WcnInfoFile.WriteLine("	- Creation Date		: 	" + objFile.CreationDate  )
		WcnInfoFile.WriteLine("	- Description		: 	" + objFile.Description  )
		WcnInfoFile.WriteLine("	- Installation Date	: 	" +  objFile.InstallDate )
		WcnInfoFile.WriteLine("	- In Use Count		:	" + objFile.InUseCount   )
		WcnInfoFile.WriteLine("	- Last Accessed		: 	" + objFile.LastAccessed  )
		WcnInfoFile.WriteLine("	- Last Modified 	: 	" + objFile.LastModified  )
		WcnInfoFile.WriteLine("	- Status		: 	" + objFile.Status  )
        Next
    Next




    WcnInfoFile.WriteLine("")
    WcnInfoFile.WriteLine("")
    WcnInfoFile.WriteLine("-------------------------------")
    WcnInfoFile.WriteLine("+ Network Adapters Information ")
    WcnInfoFile.WriteLine("-------------------------------")
    WcnInfoFile.WriteLine("")

    strQuery = "Select * from Win32_NetworkAdapter " 
    
    Set AdapterProp = objWMIService.ExecQuery _
            (strQuery)


    For Each objFile in AdapterProp
    	WcnInfoFile.WriteLine("")
	WcnInfoFile.WriteLine("---------------------")
	WcnInfoFile.WriteLine("DeviceID  :  " + objFile.DeviceID   )
	WcnInfoFile.WriteLine("---------------------")
	WcnInfoFile.WriteLine("	- Adapter Type		:  	" + objFile.AdapterType  )
	WcnInfoFile.WriteLine("	- Auto Sense			: 	" + objFile.AutoSense )
	WcnInfoFile.WriteLine("	- Description 		: 	" + objFile.Description   )
	WcnInfoFile.WriteLine("	- NetConnectionID  	: 	" + objFile.NetConnectionID   )
	WcnInfoFile.WriteLine("	- GUID 			: 	" + objFile.GUID )
	WcnInfoFile.WriteLine("	- MACAddress  		: 	" + objFile.MACAddress  )
	WcnInfoFile.WriteLine("	- Manufacturer   	: 	" + objFile.Manufacturer   )
	WcnInfoFile.WriteLine("	- MaxSpeed    	: 	" + objFile.MaxSpeed    )
	WcnInfoFile.WriteLine("	- Speed        		: 	" +  objFile.Speed    )
	WcnInfoFile.WriteLine("	- Name     		: 	" + objFile.Name     )
	
	Select Case objFile.NetConnectionStatus 
		Case 0    strAvail= "Disconnected"		    	
		Case 1    strAvail= "Connecting"
		Case 2 	strAvail= "Connected"
		Case 3 	strAvail= "Disconnecting"
		Case 4 	strAvail= "Hardware not present"
	   	Case 5 	strAvail= "Hardware disabled"
	     	Case 6 	strAvail= "Hardware malfunction"		     	
		Case 7 	strAvail= "Media disconnected"
	     	Case 8 	strAvail= "Authenticating"
	     	Case 9 	strAvail= "Authentication succeeded"
	     	Case 10 	strAvail= "Authentication failed"
	     	Case 11 	strAvail= "Invalid address"		     	
	     	Case 12	strAvail= "Credentials required"
	End Select


	WcnInfoFile.WriteLine("	- NetConnectionStatus	: 	" + strAvail )
	WcnInfoFile.WriteLine("	- NetEnabled  	: 	" +  objFile.NetEnabled  )
	WcnInfoFile.WriteLine("	- NetworkAddresses   	: 	" +  objFile.NetworkAddresses  )
	WcnInfoFile.WriteLine("	- PermanentAddress    	: 	" +  objFile.PermanentAddress   )
	WcnInfoFile.WriteLine("	- PhysicalAdapter    	: 	" +  objFile.PhysicalAdapter   )
	WcnInfoFile.WriteLine("	- PNPDeviceID     	: 	" +  objFile.PNPDeviceID    )
	WcnInfoFile.WriteLine("	- ProductName      	: 	" +  objFile.ProductName     )
	WcnInfoFile.WriteLine("	- ServiceName       	: 	" +  objFile.ServiceName      )

	WcnInfoFile.WriteLine("	- SystemName       	: 	" + objFile.SystemName       )
	WcnInfoFile.WriteLine("	- TimeOfLastReset	: 	" + objFile.TimeOfLastReset )
	WcnInfoFile.WriteLine("	- Status      	: 	" + objFile.Status      )

	Select Case objFile.StatusInfo  
		Case 1    strAvail= "Other"
		Case 2 	strAvail= "Unknown"
		Case 3 	strAvail= "Enabled"
		Case 4 	strAvail= "Disabled"
		Case 5 	strAvail= "Not Applicable"
    	End Select
		
	WcnInfoFile.WriteLine("	- StatusInfo		: 	" + strAvail )
		
       Select Case objFile.Availability 
		Case 1    strAvail= "Other"
		Case 2 	strAvail= "Unknown"
	     	Case 3 	strAvail= "Running or Full Power"
	     	Case 4 	strAvail= "Warning"
		Case 5 	strAvail= "In test"
	     	Case 6 	strAvail= "Not Applicable"
	     	Case 7 	strAvail= "Power Off"
	     	Case 8 	strAvail= "Off Line"
	     	Case 9 	strAvail= "Off Duty"
	     	Case 10 	strAvail= "Degraded"
	     	Case 11 	strAvail= "Not Installed"
	     	Case 12	strAvail= "Install Error"
	     	Case 13 	strAvail= "Power Save - Unknown"
	     	Case 14 	strAvail= "Power Save - Low Power Mode"
	     	Case 15 	strAvail= "Power Save - Standby"
	     	Case 16 	strAvail= "Power Cycle"
	     	Case 17 	strAvail= "Power Save - Warning"
	End Select

	WcnInfoFile.WriteLine("	- Availability		: 	" + strAvail )	
	WcnInfoFile.WriteLine("	- Caption 		: 	" +  objFile.Caption )	

       Select Case objFile.ConfigManagerErrorCode 
	    	Case 0    strAvail= "Device is working properly"
	     	Case 1 	strAvail= "Device is not configured correctly"
	     	Case 2 	strAvail= "Windows cannot load the driver for this device"
	     	Case 3 	strAvail= "Driver for this device might be corrupted, or the system may be low on memory or other resources"    	
	     	Case 4 	strAvail= "Device is not working properly. One of its drivers or the registry might be corrupted."
	     	Case 5 	strAvail= "Driver for the device requires a resource that Windows cannot manage."
	     	Case 6 	strAvail= "Boot configuration for the device conflicts with other devices"
	     	Case 7 	strAvail= "Cannot filter"
	     	Case 8 	strAvail= "Driver loader for the device is missing"
	     	Case 9 	strAvail= "Device is not working properly. The controlling firmware is incorrectly reporting the resources for the device"
	     	Case 10 	strAvail= "Device cannot start"
	     	Case 11  strAvail= "Device failed"
	     	Case 12	strAvail= "Device cannot find enough free resources to use"
	     	Case 13	strAvail= "Windows cannot verify the device's resources"
	     	Case 14	strAvail= "Device cannot work properly until the computer is restarted"
	     	Case 15	strAvail= "Device is not working properly due to a possible re-enumeration problem"
	     	Case 16	strAvail= "Windows cannot identify all of the resources that the device uses"
	     	Case 17	strAvail= "Device is requesting an unknown resource type."
	     	Case 18	strAvail= "Device drivers must be reinstalled"
	     	Case 19	strAvail= "Failure using the VxD loader"
	     	Case 20	strAvail= "Registry might be corrupted."
		Case 21	strAvail= "System failure. If changing the device driver is ineffective, see the hardware documentation. Windows is removing the device"
 		Case 22	strAvail= "Device is disabled"
   		Case 23	strAvail= "System failure. If changing the device driver is ineffective, see the hardware documentation"
   		Case 24	strAvail= "Device is not present, not working properly, or does not have all of its drivers installed."
   		Case 25	strAvail= "Windows is still setting up the device"
   		Case 27 strAvail= "Device does not have valid log configuration."
   		Case 28 strAvail= "Device drivers are not installed."
   		Case 29 strAvail= "Device is disabled. The device firmware did not provide the required resources."
   		Case 30	strAvail= "Device is using an IRQ resource that another device is using."
   		Case 31	strAvail= "Device is not working properly. Windows cannot load the required device drivers." 			
	End Select

	WcnInfoFile.WriteLine("	- ConfigManagerErrorCode:	" + strAvail )
	WcnInfoFile.WriteLine("	- Error Cleared 	: 	" + objFile.ErrorCleared )
	WcnInfoFile.WriteLine("	- Error Description  	: 	" + objFile.ErrorDescription)
	WcnInfoFile.WriteLine("	- LastErrorCode		: 	" + objFile.LastErrorCode)
	WcnInfoFile.WriteLine("	- Index 	 	: 	" + objFile.Index)
	WcnInfoFile.WriteLine("	- Installed  	: 	" + objFile.Installed  )
	WcnInfoFile.WriteLine("	- Install Date   	: 	" + objFile.InstallDate   )				
	WcnInfoFile.WriteLine("	- InterfaceIndex 	: 	" + objFile.InterfaceIndex )	
    Next
    WcnInfoFile.Close





    Set WcnInfoFile = FSO.OpenTextFile(outputFileName, 8, True)
    WcnInfoFile.WriteLine("")
    WcnInfoFile.WriteLine("-----------------------")
    WcnInfoFile.WriteLine("+ ipconfig information")
    WcnInfoFile.WriteLine("-----------------------")
    WcnInfoFile.WriteLine("")
    WcnInfoFile.Close


    cmd = "cmd /c ipconfig /all >> " & outputFileName
    objShell.Run cmd, 0, True



    Set WcnInfoFile = FSO.OpenTextFile(outputFileName, 8, True)
    WcnInfoFile.WriteLine("")    
    WcnInfoFile.WriteLine("----------------------")
    WcnInfoFile.WriteLine("+ Softap Capabilities ")
    WcnInfoFile.WriteLine("----------------------")
    WcnInfoFile.WriteLine("")
    WcnInfoFile.Close

    cmd = "cmd /c netsh wlan show device >> " & outputFileName
    objShell.Run cmd, 0, True

    Set WcnInfoFile = FSO.OpenTextFile(outputFileName, 8, True)
    WcnInfoFile.WriteLine("")    
    WcnInfoFile.WriteLine("----------------------")
    WcnInfoFile.WriteLine("+ Dump wcncsvc RegKey ")
    WcnInfoFile.WriteLine("----------------------")
    WcnInfoFile.WriteLine("")
    WcnInfoFile.Close

    cmd = "cmd /c reg query HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\wcncsvc\Parameters >> " & outputFileName
    objShell.Run cmd, 0, True



'    Set shell = WScript.CreateObject( "WScript.Shell" )
'    windir = shell.ExpandEnvironmentStrings("%windir%\system32\")
'    filename = windir + "wcnwiz.dll"
'    commandname = windir + "rundll32.exe"

'    cmd = "cmd /c "& commandname &" "& filename &" , RunDumpWcnCache >> " & outputFileName
'    objShell.Run cmd, 0, True


    Set WcnInfoFile = FSO.OpenTextFile(outputFileName, 8, True)
    WcnInfoFile.WriteLine("")    
    WcnInfoFile.WriteLine("--------------------------------")
    WcnInfoFile.WriteLine("+ Network Discovery Information.")
    WcnInfoFile.WriteLine("--------------------------------")
    WcnInfoFile.WriteLine("")
    WcnInfoFile.WriteLine("")
    WcnInfoFile.WriteLine("------------------------------")    
    WcnInfoFile.WriteLine("- Current Profile information")
    WcnInfoFile.WriteLine("------------------------------")    
    WcnInfoFile.WriteLine("")

    ' Profile Type
    Const NET_FW_PROFILE2_DOMAIN = 1
    Const NET_FW_PROFILE2_PRIVATE = 2
    Const NET_FW_PROFILE2_PUBLIC = 4

    ' Direction  
    Const NET_FW_RULE_DIR_IN = 1
    Const NET_FW_RULE_DIR_OUT = 2


    ' Create the FwPolicy2 object.
    Dim fwPolicy2    
    Dim ProfileType
    ProfileType = Array("Domain", "Private", "Public")

    Set fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")

    CurrentProfile = fwPolicy2.CurrentProfileTypes
  
    WcnInfoFile.WriteLine ("Current firewall profile is: ")

    '// The returned 'CurrentProfiles' bitmask can have more than 1 bit set if multiple profiles 
    '//   are active or current at the same time

    if ( CurrentProfile AND NET_FW_PROFILE2_DOMAIN ) then
    	WcnInfoFile.WriteLine(ProfileType(0))
    end if

    if ( CurrentProfile AND NET_FW_PROFILE2_PRIVATE ) then
	WcnInfoFile.WriteLine(ProfileType(1))
    end if

    if ( CurrentProfile AND NET_FW_PROFILE2_PUBLIC ) then
	WcnInfoFile.WriteLine(ProfileType(2))
    end if
    WcnInfoFile.Close


    cmd = "cmd /c netsh advfirewall show currentprofile >> " & outputFileName
    objShell.Run cmd, 0, True


    Set WcnInfoFile = FSO.OpenTextFile(outputFileName, 8, True)
    WcnInfoFile.WriteLine("")
    WcnInfoFile.WriteLine("----------------------------------------------")    
    WcnInfoFile.WriteLine("- Network discovery status for current profile")
    WcnInfoFile.WriteLine("----------------------------------------------")    
    WcnInfoFile.WriteLine("")               

    Dim rule
    ' Get the Rules object
    Dim RulesObject
    Set RulesObject = fwPolicy2.Rules

    
    For Each rule In Rulesobject
    	if rule.Grouping = "@FirewallAPI.dll,-32752" then
	        WcnInfoFile.WriteLine("")
	        WcnInfoFile.WriteLine("  Rule Name:          " & rule.Name)
	        WcnInfoFile.WriteLine("   ----------------------------------------------")
	        WcnInfoFile.WriteLine("  Enabled:            " & rule.Enabled)
	        WcnInfoFile.WriteLine("  Description:        " & rule.Description)
	        WcnInfoFile.WriteLine("  Application Name:   " & rule.ApplicationName)
	        WcnInfoFile.WriteLine("  Service Name:       " & rule.ServiceName)

       	Select Case rule.Direction
	            Case NET_FW_RULE_DIR_IN  WcnInfoFile.WriteLine("  Direction:          In")
       	     Case NET_FW_RULE_DIR_OUT WcnInfoFile.WriteLine("  Direction:          Out")
	        End Select
	
	end if
    Next

    WcnInfoFile.Close

    
    
End Sub



Sub GetWirelessAdapterInfo(outputFile)
    On Error Resume Next
    Dim adapters, objReg
    Dim adapterDetailNames, adapterDetailRegValNames

    adapterDetailNames = Array("Driver Description", "Adapter Guid", "Hardware ID", "Driver Date", "Driver Version", "Driver Provider")
    adapterDetailRegValNames = Array("DriverDesc", "NetCfgInstanceId", "MatchingDeviceId", "DriverDate", "DriverVersion", "ProviderName")

    IHVDetailNames = Array("ExtensibilityDLL", "UIExtensibilityCLSID", "GroupName", "DiagnosticsID")
    IHVDetailRegValNames = Array("ExtensibilityDLL", "UIExtensibilityCLSID", "GroupName", "DiagnosticsID")

    HKEY_LOCAL_MACHINE = &H80000002
    strComputer = "."

    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
                     strComputer & "\root\default:StdRegProv")


    strKeyPath = "SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}\"

    objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, adapterSet

    For Each adapter In adapterSet
        If StrComp("Properties", adapter) Then
            fullstrKeyPath = strKeyPath + adapter
            objReg.GetDWORDValue HKEY_LOCAL_MACHINE, fullstrKeyPath, "*IfType", ifType
            If ifType = 71 Then
                for I = 0 to UBound(adapterDetailNames)
                    objReg.GetStringValue HKEY_LOCAL_MACHINE, fullstrKeyPath, adapterDetailRegValNames(I), info
                    outputFile.WriteLine(adapterDetailNames(I) + " = " + info)
                Next

                ihvKeyPath = fullstrKeyPath + "\Ndi\IHVExtensions"
                For J = 0 to UBound(IHVDetailNames)
                    objReg.GetStringValue HKEY_LOCAL_MACHINE, ihvKeyPath, IHVDetailRegValNames(J), ihvInfo
                    outputFile.WriteLine(IHVDetailNames(J) + " = " + ihvInfo)
                Next
                    objReg.GetDWordValue HKEY_LOCAL_MACHINE, ihvKeyPath, "AdapterOUI", ihvInfo
                    outputFile.WriteLine("AdapterOUI = " + CSTR(ihvInfo))
                outputFile.WriteLine()
            End If
        End If
    Next

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    tempFile = "tempfile.txt"
    cmd = "cmd /c tasklist > " & tempFile
    objShell.Run cmd, 0, True

    Set objTextFile = FSO.OpenTextFile(tempFile, 1)
    strIHVOutput = objTextFile.ReadAll()

    Set regEx = New RegExp
    regEx.Pattern = "^wlanext.exe[\s|a-z|A-Z|\d]*"
    regEx.Multiline = True
    regEx.IgnoreCase = True
    regEx.Global = True

    Set Matches = regEx.Execute(strIHVOutput)

    For Each match in Matches
        outputFile.WriteLine(match.Value)
    Next

End Sub

Sub GetWirelessAutoconfigLog(logFileName)
    On Error Resume Next

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    'Export the operational log
    cmd = "cmd /c wevtutil epl ""Microsoft-Windows-WLAN-AutoConfig/Operational"" " & logFileName
    objShell.Run cmd, 0, True	

    'Archive the log so that it can be read on different machines
    cmd = "cmd /c wevtutil al " & logFileName
    objShell.Run cmd, 0, True	
End Sub

Sub GetWiredAdapterInfo(outputFile)
    On Error Resume Next
    Dim adapters, objReg
    Dim adapterDetailNames, adapterDetailRegValNames

    adapterDetailNames = Array("Driver Description", "Adapter Guid", "Hardware ID", "Driver Date", "Driver Version", "Driver Provider")
    adapterDetailRegValNames = Array("DriverDesc", "NetCfgInstanceId", "MatchingDeviceId", "DriverDate", "DriverVersion", "ProviderName")


    HKEY_LOCAL_MACHINE = &H80000002
    strComputer = "."

    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
                     strComputer & "\root\default:StdRegProv")


    strKeyPath = "SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}\"

    objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, adapterSet

    For Each adapter In adapterSet
        If StrComp("Properties", adapter) Then
            fullstrKeyPath = strKeyPath + adapter
            objReg.GetDWORDValue HKEY_LOCAL_MACHINE, fullstrKeyPath, "*IfType", ifType
            If ifType = 6 Then
                for I = 0 to UBound(adapterDetailNames)
                    objReg.GetStringValue HKEY_LOCAL_MACHINE, fullstrKeyPath, adapterDetailRegValNames(I), info
                    outputFile.WriteLine(adapterDetailNames(I) + " = " + info)
                Next
                outputFile.WriteLine()
            End If
        End If
    Next
End Sub


Sub GetEnvironmentInfo(outputFileName)
    On Error Resume Next
    Dim envInfoFile

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    cmd = "cmd /c netsh wlan show all > " & outputFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c netsh lan show interfaces >> " & outputFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c netsh lan show settings >> " & outputFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c netsh lan show profiles >> " & outputFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c ipconfig /all >> " & outputFileName
    objShell.Run cmd, 0, True

    RunCmd "echo.", outputFileName
    RunCmd "echo ROUTE PRINT:", outputFileName
    RunCmd "route print", outputFileName
    
    Set envInfoFile = FSO.OpenTextFile(outputFileName, 8, True)
    envInfoFile.WriteLine("")
    envInfoFile.WriteLine("Machine certificates...")
    envInfoFile.WriteLine("")
    envInfoFile.Close

    cmd = "cmd /c certutil -v -store -silent My >> " & outputFileName
    objShell.Run cmd, 0, True

    Set envInfoFile = FSO.OpenTextFile(outputFileName, 8, True)
    envInfoFile.WriteLine("")
    envInfoFile.WriteLine("User certificates...")
    envInfoFile.WriteLine("")
    envInfoFile.Close

    cmd = "cmd /c certutil -v -store -silent -user My >> " & outputFileName
    objShell.Run cmd, 0, True
End Sub

'Comments: Function to dump a tree under a registry path into a file
Sub DumpRegKey(outputFileName,regpath)
    On Error Resume Next
    Dim cmd

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    cmd = "cmd /c reg export " & regpath & "  " & outputFileName & " /y"
    objShell.Run cmd, 0, True

End Sub

Sub DumpAllKeys
    On Error Resume Next
    Dim NotifRegFile, RegFolder, Key

    RegFolder = "Reg"

    if Not FSO.FolderExists(RegFolder) Then
       FSO.CreateFolder RegFolder
    End If

	' Dump WLAN registry keys
    AllCredRegFile = RegFolder + "\AllCred.reg.txt"
    AllCredFilterFile = RegFolder + "\AllCredFilter.reg.txt"
    CredRegFileA = RegFolder + "\{07AA0886-CC8D-4e19-A410-1C75AF686E62}.reg.txt"
    CredRegFileB = RegFolder + "\{33c86cd6-705f-4ba1-9adb-67070b837775}.reg.txt"
    CredRegFileC = RegFolder + "\{edd749de-2ef1-4a80-98d1-81f20e6df58e}.reg.txt"
    APIPermRegFile = RegFolder + "\APIPerm.reg.txt"
    NotifRegFile = RegFolder + "\Notif.reg.txt"
    GPTRegFile = RegFolder + "\GPT.reg.txt"
    CUWlanSvcRegFile = RegFolder + "\HKCUWlanSvc.reg.txt"
    LMWlanSvcRegFile = RegFolder + "\HKLMWlanSvc.reg.txt"
    NidRegFile = RegFolder + "\NetworkProfiles.reg.txt"

    call DumpRegKey(NotifRegFile ,"""HKLM\SYSTEM\CurrentControlSet\Control\Winlogon\Notifications""")
    call DumpRegKey(AllCredRegFile ,"""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\Credential Providers""")
    call DumpRegKey(AllCredFilterFile,"""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\Credential Provider Filters""")
    call DumpRegKey(CredRegFileA ,"""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\Credential Providers\{07AA0886-CC8D-4e19-A410-1C75AF686E62}""")
    call DumpRegKey(CredRegFileB ,"""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\Credential Providers\{33c86cd6-705f-4ba1-9adb-67070b837775}""")
    call DumpRegKey(CredRegFileC ,"""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\Credential Provider Filters\{edd749de-2ef1-4a80-98d1-81f20e6df58e}""")
    call DumpRegKey(APIPermRegFile ,"""HKLM\SYSTEM\CurrentControlSet\Services\Wlansvc\Parameters\WlanAPIPermissions""")

    call DumpRegKey(GPTRegFile , """HKLM\SOFTWARE\Policies\Microsoft\Windows\Wireless\GPTWirelessPolicy""")
    call DumpRegKey(CUWlanSvcRegFile ,"""HKCU\SOFTWARE\Microsoft\Wlansvc""")
    call DumpRegKey(LMWlanSvcRegFile ,"""HKLM\SOFTWARE\Microsoft\Wlansvc""")

	' Dump Dot3 registry keys
    LMDot3SvcRegFile = RegFolder + "\HKLMDot3Svc.reg.txt"
    CUDot3SvcRegFile = RegFolder + "\HKCUDot3Svc.reg.txt"
    LGPPolicyFile  = RegFolder + "\L2GP.reg.txt"

    call DumpRegKey(LMDot3SvcRegFile ,"""HKLM\SOFTWARE\Microsoft\dot3svc""")
    call DumpRegKey(CUDot3SvcRegFile ,"""HKCU\SOFTWARE\Microsoft\dot3svc""")
    call DumpRegKey(LGPPolicyFile  ,"""HKLM\SOFTWARE\Policies\Microsoft\Windows\WiredL2\GP_Policy""")

    call DumpRegKey(NidRegFile  ,"""HKLM\SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\NetworkList""")
    
End Sub

' Dump Winsock LSP catalog
Sub DumpWinsockCatalog(outputFileName)
    On Error Resume Next
    Dim envInfoFile

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    cmd = "cmd /c netsh winsock show catalog > " & outputFileName
    objShell.Run cmd, 0, True
End Sub

' Dump the Windows Firewall Configuration
Sub GetWindowsFirewallInfo(configFileName, logFileName, effectiveRulesFileName, consecLogFileName, logFileNameVerbose, consecLogFileNameVerbose)
    On Error Resume Next
    Dim envInfoFile

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    cmd = "cmd /c echo Current Profiles: > " & configFileName
    objShell.Run cmd, 0, True
    cmd = "cmd /c echo ------------------------------------------------------------------------ >> " & configFileName
    objShell.Run cmd, 0, True

    'Dump the current profiles	
    cmd = "cmd /c netsh advfirewall monitor show currentprofile >> " & configFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c echo Firewall Configuration: >> " & configFileName
    objShell.Run cmd, 0, True
    cmd = "cmd /c echo ------------------------------------------------------------------------ >> " & configFileName
    objShell.Run cmd, 0, True	

    ' Dump the firewall configuration
    cmd = "cmd /c netsh advfirewall monitor show firewall >> " & configFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c echo Connection Security  Configuration: >> " & configFileName
    objShell.Run cmd, 0, True
    cmd = "cmd /c echo ------------------------------------------------------------------------ >> " & configFileName
    objShell.Run cmd, 0, True		

    'Dump the connection security configuration
    cmd = "cmd /c netsh advfirewall monitor show consec >> " & configFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c echo Firewall Rules : >> " & configFileName
    objShell.Run cmd, 0, True
    cmd = "cmd /c echo ------------------------------------------------------------------------ >> " & configFileName
    objShell.Run cmd, 0, True		

    'Dump the firewall rules
    cmd = "cmd /c netsh advfirewall firewall show rule name=all verbose >> " & configFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c echo Connection Security  Rules : >> " & configFileName
    objShell.Run cmd, 0, True
    cmd = "cmd /c echo ------------------------------------------------------------------------ >> " & configFileName
    objShell.Run cmd, 0, True		
    
    'Dump the connection security rules
    cmd = "cmd /c netsh advfirewall consec show rule name=all verbose >> " & configFileName
    objShell.Run cmd, 0, True	
    
    'Dump the firewall rules from Dynamic Store
    
    cmd = "cmd /c echo Firewall Rules currently enforced : > " & effectiveRulesFileName
    objShell.Run cmd, 0, True
    cmd = "cmd /c echo ------------------------------------------------------------------------ >> " & effectiveRulesFileName
    objShell.Run cmd, 0, True	        
    
    cmd = "cmd /c netsh advfirewall monitor show firewall rule name=all >> " & effectiveRulesFileName
    objShell.Run cmd, 0, True
    
    'Dump the connection security rules from Dynamic Store
    
    cmd = "cmd /c echo Connection Security Rules currently enforced : >> " & effectiveRulesFileName
    objShell.Run cmd, 0, True
    cmd = "cmd /c echo ------------------------------------------------------------------------ >> " & effectiveRulesFileName
    objShell.Run cmd, 0, True		
    
    cmd = "cmd /c netsh advfirewall monitor show consec rule name=all >> " & effectiveRulesFileName
    objShell.Run cmd, 0, True	

    

    'Export the operational log
    cmd = "cmd /c wevtutil epl ""Microsoft-Windows-Windows Firewall With Advanced Security/Firewall"" " & logFileName
    objShell.Run cmd, 0, True	

    'Archive the log so that it could be read on different machines
    cmd = "cmd /c wevtutil al " & logFileName
    objShell.Run cmd, 0, True	
    
      'Export the operational log
    cmd = "cmd /c wevtutil epl ""Microsoft-Windows-Windows Firewall With Advanced Security/ConnectionSecurity"" " & consecLogFileName
    objShell.Run cmd, 0, True	

    'Archive the log so that it could be read on different machines
    cmd = "cmd /c wevtutil al " & consecLogFileName
    objShell.Run cmd, 0, True	

   
    'Export the operational log
    cmd = "cmd /c wevtutil epl ""Microsoft-Windows-Windows Firewall With Advanced Security/FirewallVerbose"" " & logFileNameVerbose
    objShell.Run cmd, 0, True	

    'Archive the log so that it could be read on different machines
    cmd = "cmd /c wevtutil al " & logFileNameVerbose
    objShell.Run cmd, 0, True	
    
      'Export the operational log
    cmd = "cmd /c wevtutil epl ""Microsoft-Windows-Windows Firewall With Advanced Security/ConnectionSecurityVerbose"" " & consecLogFileNameVerbose
    objShell.Run cmd, 0, True	

    'Archive the log so that it could be read on different machines
    cmd = "cmd /c wevtutil al " & consecLogFileNameVerbose
    objShell.Run cmd, 0, True	
    
End Sub

Sub GetWfpInfo(outputFileName, logFileName)
    On Error Resume Next

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    cmd = "cmd /c netsh wfp show filters file=" & outputFileName & " > " & logFileName
    objShell.Run cmd, 0, True

End Sub

' Dump Netio State
Sub GetNetioInfo(outputFileName)
    On Error Resume Next

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    cmd = "cmd /c netsh interface teredo show state > " & outputFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c netsh interface httpstunnel show interface >> " & outputFileName
    objShell.Run cmd, 0, True

    cmd = "cmd /c netsh interface httpstunnel show statistics >> " & outputFileName
    objShell.Run cmd, 0, True

End Sub

Sub GetDnsInfo(logFileName)
    On Error Resume Next

    Set objShell = WScript.CreateObject( "WScript.Shell" )

	RunCmd "echo IPCONFIG /DISPLAYDNS: ", logFileName	
    RunCmd "ipconfig /displaydns", logFileName

    RunCmd "echo. ", logFileName
    RunCmd "echo NETSH NAMESPACE SHOW EFFECTIVE:", logFileName
    RunCmd "netsh namespace show effective", logFileName
    
    RunCmd "echo.", logFileName
    RunCmd "echo NETSH NAMESPACE SHOW POLICY:", logFileName
    RunCmd "netsh namespace show policy", logFileName

End Sub

Sub GetNeighborInfo(logFileName)
    On Error Resume Next

    Set objShell = WScript.CreateObject( "WScript.Shell" )

	RunCmd "echo ARP -A:", logFileName
    RunCmd "arp -a", logFileName

	RunCmd "echo.", logFileName
	RunCmd "echo NETSH INT IPV6 SHOW NEIGHBORS:", logFileName
    RunCmd "netsh int ipv6 show neigh", logFileName

End Sub

Sub GetFileSharingInfo(logFileName)
    On Error Resume Next

    Set objShell = WScript.CreateObject( "WScript.Shell" )

	RunCmd "echo NBTSTAT -N:", logFileName
    RunCmd "nbtstat -n", logFileName

	RunCmd "echo.", logFileName
	RunCmd "echo NBTSTAT -C:", logFileName
    RunCmd "nbtstat -c", logFileName

	RunCmd "echo.", logFileName
	RunCmd "echo NET CONFIG RDR:", logFileName
    RunCmd "net config rdr", logFileName

	RunCmd "echo.", logFileName
	RunCmd "echo NET CONFIG SRV:", logFileName
    RunCmd "net config srv", logFileName

	RunCmd "echo.", logFileName
	RunCmd "echo NET SHARE:", logFileName
    RunCmd "net share", logFileName

End Sub

Sub GetGPResultInfo(logFileName)
    On Error Resume Next

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    cmd = "cmd /c gpresult /scope:computer /v 1> " & logFileName & " 2>&1"
    objShell.Run cmd, 0, True

End Sub

Sub GetNetEventsInfo(outputFileName, logFileName)
    On Error Resume Next

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    cmd = "cmd /c netsh wfp show netevents file=" & outputFileName & " 1> " & logFileName & " 2>&1"
    objShell.Run cmd, 0, True

End Sub

Sub GetShowStateInfo(outputFileName, logFileName)
    On Error Resume Next

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    cmd = "cmd /c netsh wfp show state file=" & outputFileName & " 1> " & logFileName & " 2>&1"
    objShell.Run cmd, 0, True

End Sub

Sub GetSysPortsInfo(outputFileName, logFileName)
    On Error Resume Next

    Set objShell = WScript.CreateObject( "WScript.Shell" )

    cmd = "cmd /c netsh wfp show sysports file=" & outputFileName & " 1> " & logFileName & " 2>&1"
    objShell.Run cmd, 0, True

End Sub


On Error Resume Next

Dim adapterInfoFile, netInfoFile, WcnInfoFile

Set FSO = CreateObject("Scripting.FileSystemObject")
Set shell = WScript.CreateObject( "WScript.Shell" )
sysdrive = shell.ExpandEnvironmentStrings("%SystemDrive%\")

configFolder = "config"
osinfoFileName = configFolder + "\osinfo.txt"
adapterinfoFileName = configFolder + "\adapterinfo.txt"
envinfoFileName = configFolder + "\envinfo.txt"
wirelessAutoconfigLogFileName = configFolder + "\WLANAutoConfigLog.evtx"
wscatFileName = configFolder + "\WinsockCatalog.txt"
wcnFileName = configFolder + "\WcnInfo.txt"
wcncachedumpFile= sysdrive + "\wcncachedump.txt"
windowsFirewallConfigFileName = configFolder + "\WindowsFirewallConfig.txt"
windowsFirewallEffectiveRulesFileName = configFolder + "\WindowsFirewallEffectiveRules.txt"
windowsFirewallLogFileName = configFolder + "\WindowsFirewallLog.evtx"
windowsFirewallConsecLogFileName = configFolder + "\WindowsFirewallConsecLog.evtx"
windowsFirewallVerboseLogFileName = configFolder + "\WindowsFirewallLogVerbose.evtx"
windowsFirewallConsecVerboseLogFileName = configFolder + "\WindowsFirewallConsecLogVerbose.evtx"
wfpfiltersfilename=configFolder + "\wfpfilters.xml"
wfplogfilename=configFolder + "\wfplog.log"
netioStateFilename=configFolder + "\netiostate.txt"
dnsInfoFileName = configFolder + "\Dns.txt"
neighborsFileName = configFolder + "\Neighbors.txt"
filesharingFileName = configFolder + "\FileSharing.txt"
gpresultFileName = configFolder + "\gpresult.txt"
neteventsFileName = configFolder + "\netevents.xml"
neteventsFileLog = configFolder + "\neteventslog.txt"
showstateFileName = configFolder + "\wfpstate.xml"
showstateFileLog = configFolder + "\wfpstatelog.txt"
sysportsFileName = configFolder + "\sysports.xml"
sysportsFileLog = configFolder + "\sysportslog.txt"


if Not FSO.FolderExists(configFolder) Then
    FSO.CreateFolder configFolder
End If

call DumpAllKeys

call GetOSInfo(osinfoFileName)

Set adapterInfoFile = FSO.OpenTextFile(adapterInfoFileName, 2, True)

call GetWirelessAdapterInfo(adapterInfoFile)
call GetWiredAdapterInfo(adapterInfoFile)

adapterInfoFile.Close

call GetWirelessAutoconfigLog(wirelessAutoConfigLogFileName)

call GetEnvironmentInfo(envinfoFileName)

call DumpWinsockCatalog(wscatFileName)

call  GetWindowsFirewallInfo(windowsFirewallConfigFileName, windowsFirewallLogFileName, windowsFirewallEffectiveRulesFileName,windowsFirewallConsecLogFileName, windowsFirewallVerboseLogFileName, windowsFirewallConsecVerboseLogFileName)

call GetWcnInfo(wcnFileName)

call GetWfpInfo(wfpfiltersfilename, wfplogfilename)

call GetNetioInfo(netioStateFilename)

call GetDnsInfo(dnsInfoFileName)

call GetNeighborInfo(neighborsFileName)

call GetFileSharingInfo(filesharingFileName)

call GetGPResultInfo(gpresultFileName)

call GetNetEventsInfo(neteventsFileName, neteventsFileLog)

call GetShowStateInfo(showstateFileName, showstateFileLog)

call GetSysPortsInfo(sysportsFileName, sysportsFileLog)
