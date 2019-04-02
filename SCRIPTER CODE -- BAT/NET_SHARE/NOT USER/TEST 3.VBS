' ----
' VBS SCRIPT TO IP ADDRESS SCAN FIND ALL COMPUTER ON NETWORK - Google Search
' https://www.google.co.uk/search?q=VBS+SCRIPT+TO+IP+ADDRESS+SCAN+FIND+ALL+COMPUTER+ON+NETWORK&rlz=1C1CHBF_en-GBGB759GB759&oq=VBS+SCRIPT+TO+IP+ADDRESS+SCAN+FIND+ALL+COMPUTER+ON+NETWORK&aqs=chrome..69i57.16748j0j7&sourceid=chrome&ie=UTF-8
' --------
' List Hostname, IP and MAC address on your network - Script Center - Spiceworks
' https://community.spiceworks.com/scripts/show/200-list-hostname-ip-and-mac-address-on-your-network
' ----

startAddress = "192.168.1.100"
numberAddresses = 155


'******* Do not edit beyond this point!! ********

Dim objExplorer, txtOutput, fs, ResFile, CSVFile

const crlf="<BR>"


Setup
LoopSearch

ResFile.Close
CSVFile.Close
showText(crlf & "Finished!")

Wscript.quit



Sub LoopSearch
	'Break down the IP
	Dim IP(4)
	strIP = startAddress
	For i = 1 to 3
		IP(i) = Left(strIP, Instr(strIP, ".") - 1)
		strIP = Right(strIP, Len(strIP) - Instr(strIP, "."))
	Next
	
	IP(4) = strIP
	If IP(4) = 0 then IP(4) = 1
	For loopIP = 1 to numberAddresses
		tmpIP = ""
		For i = 1 to 3
			tmpIP = tmpIP & IP(i) & "."
		Next
		tmpIP = tmpIP & IP(4)

		ResPing = Ping(tmpIP)
		If ResPing = "Failed" then
			ResParsed = "No Response"
		Else
			If IsNull(ResPing) or ResPing = "" or ResPing = tmpIP then
				ResParsed = "Ping Response, No Name Resolution"
			Else
				ResParsed = ResPing
				discoverDevice ResParsed, tmpIP
			End If
		End If

		IP(4) = IP(4) + 1
		For i = 4 to 1 Step - 1
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
	resFile.WriteLine "IP Address           Node Name                    MAC Address"
	resFile.WriteLine "=============================================================="
	resFile.WriteLine

	Set CSVFile = fs.CreateTextFile (".\IPDevices.csv")
	CSVFile.WriteLine "Host,IP Address,MAC"
End Sub


Sub ShowText(txtInput)
	txtOutput = "Network Discovery In Progress:" & crlf & "==================================" & crlf
	txtOutput = txtOutput & txtInput
	objExplorer.Document.Body.InnerHTML = txtOutput
End Sub


Sub writeTxt(txtIP, txtRes, txtMAC)
	strPad = "                            "
	CSVFile.WriteLine txtRes & "," & txtIP & "," & txtMAC
	ResFile.WriteLine (txtIP & Left(strPad, 21 - Len(txtIP)) & txtRes & Left(strPad,29 - Len(txtRes)) & txtMAC)
End Sub


Sub discoverDevice(strDeviceName, strIP)
	showText(strDeviceName)
	On Error Resume Next
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strDeviceName & "\root\cimv2")
	Set colNICS = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	For Each objItem in colNICS
		writeTxt strIP, strDeviceName, objItem.MACAddress
	Next
End Sub