' [SETUP]
' SCANIP=True
' IPADDRESS=192.168.4.12
' SUBNETMASK=255.255.255.240
' NUMNODES=2

' # If you assign a number to "numNodes" it will ignore the subnet mask and just scan nodes from 
' # the IPAddress to that many nodes
' #
' # Set numNodes=0 and script will calculate the full IP range from the subnet mask, regardless of the
' # starting IPADDRESS

' # Scan for Active Directory information
' SCANAD=True


' *****************************************************************************************************************
' Script Begins Here
' *****************************************************************************************************************
' Network Discovery Script
' Author: Martin Pugh  (martin@pughspace.com)
' First release: 7/31/09
' =============================================
' Scans through IP addresses and does reverse DNS lookup report, Windows basic configuration report,
' and Active Directory report

Dim objExplorer, txtOutput, fs, ResFile, CSVFile, SrvFile, DTFile, startAddress, numberAddresses, binScanAD
Dim binScanIP
ReDim arrDupName(0)

const crlf="<BR>"


Setup
If binScanIP = True Then
	LoopSearch
End If

If binScanAD = True Then 
	getADInfo
End If



showText(crlf & "Finished!")

'Wscript.Sleep 10000
'objExplorer.Quit




Sub LoopSearch
	'Break down the IP
	IP = Split(startAddress, ".")
	If IP(3) = 0 then IP(3) = 1

	For loopIP = 1 to numberAddresses
		tmpIP = Join(IP, ".")
	  
		ResPing = Ping(tmpIP)

		If ResPing = "Failed" then
			ResParsed = "No Response"
		Else
			If IsNull(ResPing) or ResPing = "" or ResPing = tmpIP then
				ResParsed = "Ping Response, No Name Resolution"
			Else
				ResParsed = ResPing
				discoverDevice ResParsed
			End If
		End If
		writeTxt tmpIP, ResParsed
      
		IP(3) = IP(3) + 1
		For i = 3 to 0 Step - 1
			If IP(i) > 254 and i > 1 then
				IP(i) = 1
				IP(i - 1) = IP(i - 1) + 1
			Else
				If IP(1) > 254 then
					showText("Inputted IP range ran past valid IP range")
					wscript.Quit(0)
				End If
			End If
		Next
	Next
	ResFile.Close
	CSVFile.Close
	SrvFile.Close
	DTFile.Close
End Sub


Function Ping(strHost)
   Dim objPing, objRetStatus

   showText("Pinging " & strHost & "...")
   Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & strHost & "' AND ResolveAddressNames = TRUE")
   For Each objRetStatus in objPing
      If IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode <> 0 then
         Ping = "Failed"
      Else
         Ping = objRetStatus.ProtocolAddressResolved
      End if
   Next
   Set objPing = Nothing
End Function 


Function prepSize(numSize)
   if numSize > 0 then
      numSize = (numSize / 1024) / 1024
      strMem = "MB"
      If numSize > 1000 then
         strMem = "GB"
         numSize = numSize / 1024
      End If
      numSize = round(numSize, 2)
      prepSize = numSize & " " & strMem
   Else 
      PrepSize = ""
   End If
End Function


Sub Setup
   Set objExplorer = WScript.CreateObject("InternetExplorer.Application")
   objExplorer.Navigate "about:blank"   
   objExplorer.ToolBar = 0
   objExplorer.StatusBar = 0
   objExplorer.Width = 400
   objExplorer.Height = 200 
   objExplorer.Left = 100
   objExplorer.Top = 100

   Do While (objExplorer.Busy)
       Wscript.Sleep 200
   Loop

   objExplorer.Visible = 1    
   txtOutput=""

   Set fs = CreateObject ("Scripting.FileSystemObject")
   Set ResFile = fs.CreateTextFile (".\IPDevices.txt")
   writeTxtLine("IP Address                   Node Name")
   writeTxtLine("==============================================================")
   writeTxtLine("")

   Set CSVFile = fs.CreateTextFile (".\IPDevices.csv")
   CSVFile.WriteLine "IP Address,Host"

   Set SrvFile = fs.CreateTextFile (".\WindowsServerList.csv")
   SrvFile.WriteLine "Computer,Make/Model,Service Tag,Serial Number,OS,Processor,RAM,Hard Drives"

   Set DTFile = fs.CreateTextFile (".\WindowsDesktopList.csv")
   DTFile.WriteLine "Computer,Make/Model,Service Tag,Serial Number,OS,Processor,RAM,Hard Drives"
   
   	Set readFile = fs.OpenTextFile (".\nd.ini")
	Do Until readFile.AtEndOfStream
		strLine = readFile.ReadLine
		If InStr(strLine, "=") > 0 Then
			strValue = Trim(Right(strLine, Len(strLine) - InStr(strLine, "=")))
		Else
			strValue = ""
		End If
			
		If UCase(Left(strLine, 9)) = "IPADDRESS" Then
			strIP = strValue
		End If 
		
		If UCase(Left(strLine, 10)) = "SUBNETMASK" Then
			strMask = strValue
		End If
		
		If UCase(Left(strLine, 8)) = "NUMNODES" Then
			numNodes = strValue
		End If
		
		If UCase(Left(strLine, 6)) = "SCANIP" Then
			If UCASE(strValue) = "YES" or UCASE(strValue) = "TRUE" or strValue = "1" Then
				binScanIP = TRUE
			Else
				binScanIP = FALSE
			End If
		End If
			
		If UCase(Left(strLine, 6)) = "SCANAD" Then
			If UCASE(strValue) = "YES" or UCASE(strValue) = "TRUE" or strValue = "1" Then
				binScanAD = TRUE
			Else
				binScanAD = FALSE
			End If
		End If
	Loop
	
	If IsNull(numNodes) or numNodes = "" then numNodes = 0
	If numNodes <= 0 Then
		startAddress = CalcRangeStart(strIP, strMask)
		numberAddresses = MaskLength(strMask)
		numberAddresses = GetNumberOfAvailableHostAddresses(numberAddresses)
	Else
		startAddress = strIP
		numberAddresses = numNodes
	End If
End Sub


Sub ShowText(txtInput)
   txtOutput = "Network Discovery In Progress:" & crlf & "==================================" & crlf
   txtOutput = txtOutput & txtInput
   objExplorer.Document.Body.InnerHTML = txtOutput
End Sub


Sub writeTxt(txtIP, txtRes)
   strPad = "............................"
   CSVFile.WriteLine chr(34) & txtIP & chr(34) & "," & chr(34) & txtRes & chr(34)
   ResFile.WriteLine (txtIP & Left(strPad, 29 - Len(txtIP)) & txtRes)
End Sub


Sub writeTxtLine(txtInput)
   ResFile.WriteLine txtInput
End Sub



Sub discoverDevice(strDeviceName)
   On Error Resume Next
   Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strDeviceName & "\root\cimv2")
   If Err.Number <> 0 Then
      Err.Clear
      On Error Goto 0
   Else
      On Error Goto 0
      Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
      For Each objItems in colItems
         strHost = objItems.CSName
         strOS = Trim(objItems.Caption)
         If (InStr(strOS, "200") or InStr(UCase(strOS), "W")) > 0 then
            bIsNT = False
            strOSSP = Trim(objItems.ServicePackMajorVersion & "." & objItems.ServicePackMinorVersion)
         Else 
            bIsNT = True
            strOSSP = "."
         End If
         If strOSSP <> "." then
            strOS = strOS & " SP: " & strOSSP
         End If
      Next
      'wscript.echo strOS
      Set colItems = Nothing

      If InStr(UCase(strOS), "W") > 0 then
         Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
         numCPUs = 0
         For Each objItems in colItems
            numCPUs = NumCPUs + 1
            strCPUName = Trim(objItems.Name)
         Next
         Set colItems = Nothing

         bFound = False
         For i = 1 to UBound(arrDupName)
            If UCase(arrDupName(i)) = UCase(strHost) then
               bFound = True
               Exit For
            End If
         Next
         If bFound = False then
            showText("Auditing " & strHost & "...")
            i = UBound(arrDupName) + 1
            ReDim Preserve arrDupName(i)
            arrDupName(i) = strHost

            strTemp = ""
            If numCPUs > 1 then
               strTemp = "(" & numCPUs & ") "
            End If
            strCPUName = strTemp & strCPUName


            'Get Computer Info
            If bIsNT = False then
               Set colItems = objWMIService.ExecQuery("Select * from Win32_BaseBoard")
               For Each objItem in colItems
                  strMakeModel = Trim(objItem.Manufacturer) & " " & Trim(objItem.Model)
                  strServiceTag = Trim(objItem.Product)
                  strSN = Trim(objItem.SerialNumber)
               Next
               set colItems = Nothing

               'Get Memory Info
               Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
               numMemCap = 0
               For Each objItem in colItems
                  numMemCap = numMemCap + objItem.Capacity
               Next
               strRAM = prepSize(numMemCap)
               set colItems = Nothing
            End If

            'Get Logical Drive Info
            Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
            strHardDrive = ""
            t = 0
            For Each objItem in colItems
               If objItem.DriveType = 3 then
                  t = t + 1
                  If t > 1 then
                     strHardDrive = strHardDrive & ","
                  End If
                  strHardDrive = strHardDrive & chr(34) & objItem.Name & " " & prepSize(objItem.Size) & " (" & prepSize(objItem.FreeSpace) & " free)" & chr(34)
               End If
            Next
            Set colItems = Nothing



            'Save Line to text file
            strLine = chr(34) & strHost & chr(34) & "," & chr(34) & strMakeModel & chr(34) & "," & chr(34) & strServiceTag & chr(34) & "," & chr(34) & strSN & chr(34) & "," & chr(34) & strOS & chr(34) & "," & chr(34) & strCPUName & chr(34) & "," & chr(34) & strRAM & chr(34) & "," & strHardDrive
            If InStr(UCase(strOS), "SERVER") > 0 then
               SrvFile.WriteLine strLine
            Else
               DTFile.WriteLine strLine
            End If
         Else
            Set colItems = Nothing
         End If
      End If

      Set objWMIService = Nothing
   End If
End Sub


Sub getADInfo
   ShowText("Getting Active Directory Information...")
   Set ADFile = fs.CreateTextFile ("./AD Information.txt")

   On Error Resume Next
   Set objRootDSE = GetObject("LDAP://RootDSE")
   If Err.Number = 0 then
      on Error Goto 0
      strConfig = objRootDSE.Get("configurationNamingContext")

      'Determine AD Name
      Set WSHNetwork = CreateObject("WScript.Network")
      strDomain = WSHNetwork.UserDomain
      Set WSHNetwork = Nothing

      ADFile.WriteLine "Domain Name: " & strDomain
      ADFile.WriteLine
      ADFile.WriteLine


      'Determine AD Sites
      strSitesContainer = "LDAP://cn=Sites," & strConfig
      Set objSitesContainer = GetObject(strSitesContainer)
      objSitesContainer.Filter = Array("site")
      ADFile.WriteLine "AD Sites:"
      For Each objSite In objSitesContainer
         ADFile.WriteLine "  Site Name: " & removeCN(objSite.Name)
      Next
      ADFile.WriteLine
      ADFile.WriteLine


      'Find Domain Controllers
      ' Use ADO to search Active Directory for ObjectClass nTDSDSA.
      Set objCommand = CreateObject("ADODB.Command")
      Set objConnection = CreateObject("ADODB.Connection")
      objConnection.Provider = "ADsDSOObject"
      objConnection.Open "Active Directory Provider"
      objCommand.ActiveConnection = objConnection
   
      strQuery = "<LDAP://" & strConfig & ">;(ObjectClass=nTDSDSA);AdsPath;subtree"
      objCommand.CommandText = strQuery
      objCommand.Properties("Page Size") = 100
      objCommand.Properties("Timeout") = 30
      objCommand.Properties("Cache Results") = False

      Set objRecordSet = objCommand.Execute

      ' The parent object of each object with ObjectClass=nTDSDSA is a Domain
      ' Controller. The parent of each Domain Controller is a "Servers"
      ' container, and the parent of this container is the "Site" container.
      ADFile.WriteLine "Domain Controllers:"
      i = 0
      ReDim arrDC(0)
      Do Until objRecordSet.EOF
         i = i + 1
         ReDim Preserve arrDC(i)
         Set objDC = GetObject(GetObject(objRecordSet.Fields("AdsPath")).Parent)
         Set objSite = GetObject(GetObject(objDC.Parent).Parent)
         arrDC(i) = objDC.cn
         ADFile.WriteLine "    DC: " & removeCN(objDC.cn)
         ADFile.WriteLine "  Site: " & removeCN(objSite.Name)
         ADFile.WriteLine
         objRecordSet.MoveNext
      Loop
      ADFile.WriteLine
      ADFile.WriteLine

      ' Clean up.
      objConnection.Close
      Set objCommand = Nothing
      Set objConnection = Nothing
      Set objRecordSet = Nothing
      Set objDC = Nothing
      Set objSite = Nothing


      ADFile.WriteLine "FSMO Role Holders:"
      'Schema Master
      Set objSchema = GetObject("LDAP://" & objRootDSE.Get("schemaNamingContext"))
      strSchemaMaster = objSchema.Get("fSMORoleOwner")
      Set objNtds = GetObject("LDAP://" & strSchemaMaster)
      Set objComputer = GetObject(objNtds.Parent)
      ADFile.WriteLine "  Forest-wide Schema Master FSMO:        " & removeCN(objComputer.Name)
      Set objNtds = Nothing
      Set objComputer = Nothing
   
      'Domain Naming Master
      Set objPartitions = GetObject("LDAP://CN=Partitions," & objRootDSE.Get("configurationNamingContext"))
      strDomainNamingMaster = objPartitions.Get("fSMORoleOwner")
      Set objNtds = GetObject("LDAP://" & strDomainNamingMaster)
      Set objComputer = GetObject(objNtds.Parent)
      ADFile.WriteLine "  Forest-wide Domain Naming Master FSMO: " & removeCN(objComputer.Name)
      Set objNtds = Nothing
      Set objComputer = Nothing
  
      'PDC Emulator
      Set objDomain = GetObject("LDAP://" & objRootDSE.Get("defaultNamingContext"))
      strPdcEmulator = objDomain.Get("fSMORoleOwner")
      Set objNtds = GetObject("LDAP://" & strPdcEmulator)
      Set objComputer = GetObject(objNtds.Parent)
      ADFile.WriteLine "  Domain's PDC Emulator FSMO:            " & removeCN(objComputer.Name)
      Set objNtds = Nothing
      Set objComputer = Nothing
  
      'RID Master
      Set objRidManager = GetObject("LDAP://CN=RID Manager$,CN=System," & objRootDSE.Get("defaultNamingContext"))
      strRidMaster = objRidManager.Get("fSMORoleOwner")
      Set objNtds = GetObject("LDAP://" & strRidMaster)
      Set objComputer = GetObject(objNtds.Parent)
      ADFile.WriteLine "  Domain's RID Master FSMO:              " & removeCN(objComputer.Name)
      Set objNtds = Nothing
      Set objComputer = Nothing
  
      'Infrastructure Master
      Set objInfrastructure = GetObject("LDAP://CN=Infrastructure," & objRootDSE.Get("defaultNamingContext"))
      strInfrastructureMaster = objInfrastructure.Get("fSMORoleOwner")
      Set objNtds = GetObject("LDAP://" & strInfrastructureMaster)
      Set objComputer = GetObject(objNtds.Parent)
      ADFile.WriteLine "  Domain's Infrastructure Master FSMO:   " & removeCN(objComputer.Name)
      Set objNtds = Nothing
      Set objComputer = Nothing
      ADFile.WriteLine
      ADFile.WriteLine

      Set objRootDSE = Nothing


      'Find GC's
      Const NTDSDSA_OPT_IS_GC = 1
 
      ADFile.WriteLine "Global Catalogs:"
      On Error Resume Next
      For i = 1 to UBound(arrDC)
         Set objRootDSE = GetObject("LDAP://" & arrDC(i) & "/rootDSE")
         strDsServiceDN = objRootDSE.Get("dsServiceName")
         Set objDsRoot  = GetObject("LDAP://" & arrDC(i) & "/" & strDsServiceDN)
         intOptions = objDsRoot.Get("options")
   
         If intOptions And NTDSDSA_OPT_IS_GC Then
            ADFile.WriteLine "  " & arrDC(i)
         End If
      Next

      Set objDsRoot = Nothing
      Set objRootDSE = Nothing
   Else
      ADFile.WriteLine
      ADFile.WriteLine "No Active Directory domain found."
   End If

   on Error Goto 0
   ADFile.Close
End Sub


Sub quickText(strInput)
   txtOutput = txtOutput & strInput
   objExplorer.Document.Body.InnerHTML = txtOutput
End Sub


Function ConvertIPToBinary(strIP)
	' Converts an IP Address into Binary

	Dim arrOctets
	Dim strBinOctet
	Dim intOctet, i, j
 
	arrOctets = Split(strIP, ".")
	For i = 0 to UBound(arrOctets)
		intOctet = CInt(arrOctets(i))
		strBinOctet = ""
		For j = 0 To 7
			If intOctet And (2^(7 - j)) Then
				strBinOctet = strBinOctet & "1"
			Else
				strBinOctet = strBinOctet & "0"
			End If
		Next
		arrOctets(i) = strBinOctet
	Next
	ConvertIPToBinary = Join(arrOctets, ".")
End Function


Function ConvertBinIPToDecimal(strBinIP)
	' Convert binary form of an IP back to decimal

	Dim arrOctets
	Dim intOctet, i, j
 
	arrOctets = Split(strBinIP, ".")
	For i = 0 to UBound(arrOctets)
		intOctet = 0
		For j = 0 to 7
			intBit = CInt(Mid(arrOctets(i), j + 1, 1))
			If intBit = 1 Then
				intOctet = intOctet + 2^(7 - j)
			End If
		Next
		arrOctets(i) = CStr(intOctet)
	Next
 
	ConvertBinIPToDecimal = Join(arrOctets, ".")
End Function


Function CalcRangeStart(strIP, strMask)
	' Generates the Network Address from the IP and Mask
	Dim arrOctets
	Dim strBinIP, strBinMask, strIPBit, strMaskBit, strBinStart
	Dim intOctet, i, j
	' Conversion of IP and Mask to binary
	strBinIP = ConvertIPToBinary(strIP)
	strBinMask = ConvertIPToBinary(strMask)
	' Bitwise AND operation (except for the dot)
	For i = 1 to Len(strBinIP)
		strIPBit = Mid(strBinIP, i, 1)
		strMaskBit = Mid(strBinMask, i, 1)
		If strIPBit = "1" And strMaskBit = "1" Then
			strBinStart = strBinStart & "1"
		ElseIf strIPBit = "." Then
			strBinStart = strBinStart & strIPBit
		Else
			If i = Len(strBinIP) Then
				strBinStart = strBinStart & "1"
			Else
				strBinStart = strBinStart & "0"
			End If
		End If
	Next

	' Conversion of Binary IP to Decimal
	CalcRangeStart = ConvertBinIPToDecimal(strBinStart)
End Function


Function GetNumberOfAvailableHostAddresses(numMaskLength)
	numHosts = -1
	numAvailableBits = 32 - numMaskLength
	'Number of Addresses Available for Hosts in Subnet = 2(32 – Number of Masked Bits) – 2
	numHosts = (2 ^ numAvailableBits) - 2
	If numHosts < 0 then numHosts = 2
	GetNumberOfAvailableHostAddresses = numHosts - 1
End Function


Function MaskLength(strMask)
	' Converts an subnet mask into a mask length in bits
 
	Dim arrOctets
	Dim intOctet, intMaskLength, i, j
 
	arrOctets = Split(strMask, ".")
	For i = 0 to UBound(arrOctets)
		intOctet = CInt(arrOctets(i))
		For j = 0 To 7
			If intOctet And (2^(7 -j)) Then
				intMaskLength = intMaskLength + 1
			End If
		Next
	Next
	MaskLength = intMaskLength
End Function


Function removeCN(strName)
	removeCN = Replace(strName, "CN=", "")
End Function