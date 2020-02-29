Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

' -------------------------------------------------------------------
' WORK TODAY
' -------------------------------------------------------------------
' Wed 18-Sep-2019 21:21:07
' Wed 18-Sep-2019 22:55:17
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' WORK TODAY
' -------------------------------------------------------------------
' Mon 17-Feb-2020 02:10:00
' Mon 17-Feb-2020 02:10:00
' -------------------------------------------------------------------
' COPY FROM 
' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 40-RUN EXE EXPLORER LANCHER.VBS
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' MAIN USER FROM 
' Project_Check_Date
' SUB ROUTINE EXAMPLE FROM -- 
' D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe
' AND ALL CODE THAT UPDATE ITSELF
' PASS 3 PARAM
' 1 -- THE EXE TO RUN
' 2 -- SOURCE OF EXE TO COPY OVER TO DESTINATION BEFORE RUN
' 3 -- TARGET OF EXE -- WHICH MOST LIKLEY SAME AS PARAM
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' HOW TO OPEN A LAUNCHER FOR VBSCRIPT TO PASS THE ARGUMENT THE PROPER METHOD
' THE APP THE WANT USER HERE
' SEND ARGUMENT AND ENCODED IN SAME WAY AS URL WAY
' WITH A %20 FOR SPACE
' THE AUGMENT ARE DIVIDED BY SPACE
' AND UNABLE SPLIT ARRAY THEM IN WITHOUT A LOAD OF SPLIT LINE 
' WHERE EACH SPACE IS IN THE LINE
' I WAS JUST ONTO METHOD LOT PEOPLE USE FOR ADD WILDCARD FOR SPACE
' AND I TAKEN URL FORMAT
' IT SEEM THE CODE HERE IS FROM DEMO ON THE INTERNET
' AND I HAD HARD WORK SEARCHER
' AND NOTHING BETTER THAN THIS AROUND
' JUST A PATHETIC GET THE ARGS 
' AND THEN COMPLAIN FOR HALF AND HOUR
'
' USE STAR METHOD THAN URL METHOD SPACE TO STAR ENCODE
' -------------------------------------------------------------------
' [ Wednesday 21:28:50 Pm_18 September 2019 ]
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' THIS FILE HERE
' -------------------------------------------------------------------
' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 40-RUN EXE EXPLORER LANCHER.VBS
' IS LINKER ON
' D:\VB6\VB-NT\00_Best_VB_01\Shell Explorer Loader\VBS 40-RUN EXE EXPLORER LANCHER.VBS
' -------------------------------------------------------------------
' NEW BEST WAY TO PASS ARGUMENT ON VBSCRIPT FROM ANY APP WANT COMMAND LINE TO ANOTHER
' -------------------------------------------------------------------
' TYPE OF THING SEEN TO GET ARGS INTO CODE FROM OTHER APP AROUND 2003
' AND NEVER SUNK IN TO PRODUCE WHEN COMPILED IN EXE APP EXAMPLE NEVER EXPLAIN
' AS I TALK
' THE NET ONLY SHOW STANDARD WAY TO GET ARGUMENT PARAMETER 
' THE CODE HERE ORIGINAL DESIGN FOR USE BY HERE
' TONS OF REQUEST AND NOBODY COME BACK TO 
' FINAL RESULT ABLE MAKE THING MUCH MORE PRETTY
' NONE MORE FLOAT ABOUT WITH DOG EARS QUOTATION MARK
' AND SOME QUOTATION NEST BIRDY ARE NEAR IMPOSABLE 
' GIVE UP LEVEL LITERALLY ARE IMPOSABLE
' -------------------------------------------------------------------
' D:\VB6\VB-NT\00_Best_VB_01\Shell Explorer Loader\Shell Explorer Loader.exe
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Fri 20-Sep-2019 17:20:29
' Fri 20-Sep-2019 17:53:00 	-- TIME TO PATCH THIS END BIT IN NOTES AND COPY OVER 
' 							-- LIKE SNAIL PACE
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' 0 = ARG FROM OTHER APP WILL GET HERE
' -------------------------------------------------------------------
' 1ST ARG WILL BE APP
' 2ND ARG WILL BE PATH SOMETHING
' -------------------------------------------------------------------

DIM objShell
Set objShell = Wscript.CreateObject("WScript.Shell")

Set objArgs = Wscript.Arguments
' -------------------------------------------------------------------
' IF objArgs.COUNT=0 THEN 
	' SCRIPT_NAME=WScript.ScriptFullName
	' SCRIPT_NAME=REPLACE(SCRIPT_NAME , WScript.ScriptName , VBCRLF + VBCRLF + WScript.ScriptName)
	' IMSGBOX=""
	' IMSGBOX=IMSGBOX+ "NOTHING TO RUN NONE PARAMETER PASS ON TO ME -- MSGBOX FROM"
	' IMSGBOX=IMSGBOX+ VBCRLF + VBCRLF + VBCRLF 
	' IMSGBOX=IMSGBOX+ SCRIPT_NAME
	' MSGBOX IMSGBOX
	' Wscript.QUIT
' END IF
' -------------------------------------------------------------------

Set objArgs = Wscript.Arguments
For Each strArg In objArgs
	PARAM = PARAM + """"+REPLACE(strArg,"*"," ")+"""" +" "
Next
IF objArgs.COUNT>0 THEN 
	PARAM=objArgs(0)
END IF
IF objArgs.COUNT>1    THEN 
	IF objArgs(1)<>"" THEN FILE_SOURCE=objArgs(1)
END IF
IF objArgs.COUNT>1    THEN 
	IF objArgs(2)<>"" THEN FILE_TARGET=objArgs(2)
END IF

PARAM = REPLACE(PARAM,"*"," ")
FILE_SOURCE=TRIM(REPLACE(FILE_SOURCE,"*"," "))
FILE_TARGET=TRIM(REPLACE(FILE_TARGET,"*"," "))

IF FILE_TARGET="" AND FILE_SOURCE<>"" THEN 
	FILE_TARGET=PARAM
END IF

IF FILE_SOURCE="" THEN 
	FILE_SOURCE="D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe"
	FILE_TARGET="D:\VB6-EXE\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe"
	PARAM=FILE_SOURCE
END IF

Set FSO = CreateObject("Scripting.FileSystemObject")

IF FILE_SOURCE<>"" THEN
	IF NOT FSO.FILEEXISTS(FILE_SOURCE) THEN
		MSGBOX "FILE_SOURCE PARAM 2 -- NOT EXIST"+VBCRLF+FILE_SOURCE
	END IF 
END IF 

IF FILE_TARGET<>"" THEN
	IF NOT FSO.FILEEXISTS(FILE_TARGET) THEN
		MSGBOX "FILE_TARGET PARAM 3 -- NOT EXIST"+VBCRLF+FILE_TARGET
	END IF 
END IF 

IF FSO.FileExists(FILE_SOURCE) Then 
	Set F1 = fso.GetFile(FILE_SOURCE)
	FileModDate1 = F1.DateLastModified
END IF
IF FSO.FileExists(FILE_TARGET) Then 
	Set F2 = fso.GetFile(FILE_TARGET)
	FileModDate2 = F2.DateLastModified
END IF

IF FILE_SOURCE<>"" THEN
	Set FILE = FSO.GetFile(FILE_SOURCE)
	FILE_SOURCE_VAR = FILE.OpenAsTextStream().Read(FILE.SIZE)
END IF 

IF FILE_TARGET<>"" THEN
	Set FILE = FSO.GetFile(FILE_TARGET)
	FILE_TARGET_VAR = FILE.OpenAsTextStream().Read(FILE.SIZE)

	FILE_SOURCE_CRC=Hex(CalculateCRC(FILE_SOURCE_VAR))
	FILE_SOURCE_EXE_FILENAME=MID(FILE_SOURCE,INSTRREV(FILE_SOURCE,"\")+1)
	FILE_TARGET_EXE_FILENAME=MID(FILE_TARGET,INSTRREV(FILE_TARGET,"\")+1)
	
	objShell.Run "TASKKILL /IM """+FILE_TARGET_EXE_FILENAME+"""", DONTShowWindow, WaitUntilFinished
	WScript.Sleep 1000
	objShell.Run "TASKKILL /F /IM """+FILE_TARGET_EXE_FILENAME+"""", DONTShowWindow, WaitUntilFinished
	WScript.Sleep 100
	
	SET F1 = FSO.GETFILE(FILE_SOURCE)
	SET F2 = FSO.GETFILE(FILE_TARGET)
	FILE_SOURCE_DATE = F1.DATELASTMODIFIED
	FILE_TARGET_DATE = F2.DATELASTMODIFIED
	
	IF FILE_SOURCE_DATE=FILE_TARGET_DATE THEN 
		' MSGBOX "SAME DATE "& FILE_TARGET_DATE
	END IF
	
	IF FILE_SOURCE_DATE>FILE_TARGET_DATE THEN 
		' MSGBOX FILE_SOURCE+VBCRLF+" TO "+VBCRLF+FILE_TARGET
		FSO.CopyFile FILE_SOURCE, FILE_TARGET
	END IF
	IF FILE_SOURCE_DATE<FILE_TARGET_DATE THEN 
		' MSGBOX FILE_TARGET+VBCRLF+" TO "+VBCRLF+FILE_SOURCE
		ON ERROR RESUME NEXT
		ERR.CLEAR
		FSO.CopyFile FILE_TARGET,FILE_SOURCE
		IF ERR.NUMBER>0 THEN 
		
		END IF
	END IF
END IF 

' MSGBOX PARAM+VBCRLF+FILE_SOURCE+VBCRLF+FILE_TARGET

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
ON ERROR RESUME NEXT
ERR.CLEAR 
WSHSHELL.RUN """"+PARAM+"""" , SHOWWINDOW, DONTWAITUNTILFINISHED
IF ERR.NUMBER>0 THEN
	MSGBOX "ERROR NOT ABLE RUN "+VBCR++VBCR+PARAM+VBCR++VBCR+"ERROR DESCRIPTION "+VBCR+VBCR+ERR.DESCRIPTION
END IF


WScript.Quit(0)




' --------------------------------------------------------------
' IF SENT-TO IS USED SOMETIME ADDITIONAL QUOTE SURROUND ARE USER
' --------------------------------------------------------------
' XF0 = UCASE(Replace(XF0,"""",""))
' --------------------------------------------------------------


				
'**************************************
' Name: CRC Checksum
' Description:Calculates the CRC for a given string.
' By: Leigh Bowers (from psc cd)
'
' Returns:A long representing the CRC checksum.
'
' Assumes:Example Usage:
' Response.Write "$" + Hex(CalculateCRC("http://www.curvesoftware.co.uk"))
'**************************************

Function shr(plValue, pbBits)
' Shift bits to the right by <pbBits>
 shr = plValue \ (2 ^ pbBits)
End Function

Function CalculateCRC(psString)
Dim sValues, alCRCTable, lCRC
Dim l
	sValues = "&h0,&h77073096,&hEE0E612C,&h990951BA,&h076DC419,&h706AF48F,&hE963A535,&h9E6495A3,&h0EDB8832,&h79DCB8A4,&hE0D5E91E,&h97D2D988,&h09B64C2B,&h7EB17CBD,&hE7B82D07,&h90BF1D91,&h1DB71064,&h6AB020F2,&hF3B97148,&h84BE41DE,&h1ADAD47D,&h6DDDE4EB,&hF4D4B551,&h83D385C7,&h136C9856,&h646BA8C0,&hFD62F97A,&h8A65C9EC,&h14015C4F,&h63066CD9,&hFA0F3D63,&h8D080DF5,&h3B6E20C8,&h4C69105E,&hD56041E4,&hA2677172,&h3C03E4D1,&h4B04D447," _
		+ "&hD20D85FD,&hA50AB56B,&h35B5A8FA,&h42B2986C,&hDBBBC9D6,&hACBCF940,&h32D86CE3,&h45DF5C75,&hDCD60DCF,&hABD13D59,&h26D930AC,&h51DE003A,&hC8D75180,&hBFD06116,&h21B4F4B5,&h56B3C423,&hCFBA9599,&hB8BDA50F,&h2802B89E,&h5F058808,&hC60CD9B2,&hB10BE924,&h2F6F7C87,&h58684C11,&hC1611DAB,&hB6662D3D,&h76DC4190,&h01DB7106,&h98D220BC,&hEFD5102A,&h71B18589,&h06B6B51F,&h9FBFE4A5,&hE8B8D433,&h7807C9A2,&h0F00F934,&h9609A88E,&hE10E9818," _
		+ "&h7F6A0DBB,&h086D3D2D,&h91646C97,&hE6635C01,&h6B6B51F4,&h1C6C6162,&h856530D8,&hF262004E,&h6C0695ED,&h1B01A57B,&h8208F4C1,&hF50FC457,&h65B0D9C6,&h12B7E950,&h8BBEB8EA,&hFCB9887C,&h62DD1DDF,&h15DA2D49,&h8CD37CF3,&hFBD44C65,&h4DB26158,&h3AB551CE,&hA3BC0074,&hD4BB30E2,&h4ADFA541,&h3DD895D7,&hA4D1C46D,&hD3D6F4FB,&h4369E96A,&h346ED9FC,&hAD678846,&hDA60B8D0,&h44042D73,&h33031DE5,&hAA0A4C5F,&hDD0D7CC9,&h5005713C,&h270241AA," _
		+ "&hBE0B1010,&hC90C2086,&h5768B525,&h206F85B3,&hB966D409,&hCE61E49F,&h5EDEF90E,&h29D9C998,&hB0D09822,&hC7D7A8B4,&h59B33D17,&h2EB40D81,&hB7BD5C3B,&hC0BA6CAD,&hEDB88320,&h9ABFB3B6,&h03B6E20C,&h74B1D29A,&hEAD54739,&h9DD277AF,&h04DB2615,&h73DC1683,&hE3630B12,&h94643B84,&h0D6D6A3E,&h7A6A5AA8,&hE40ECF0B,&h9309FF9D,&h0A00AE27,&h7D079EB1,&hF00F9344,&h8708A3D2,&h1E01F268,&h6906C2FE,&hF762575D,&h806567CB,&h196C3671,&h6E6B06E7," _
		+ "&hFED41B76,&h89D32BE0,&h10DA7A5A,&h67DD4ACC,&hF9B9DF6F,&h8EBEEFF9,&h17B7BE43,&h60B08ED5,&hD6D6A3E8,&hA1D1937E,&h38D8C2C4,&h4FDFF252,&hD1BB67F1,&hA6BC5767,&h3FB506DD,&h48B2364B,&hD80D2BDA,&hAF0A1B4C,&h36034AF6,&h41047A60,&hDF60EFC3,&hA867DF55,&h316E8EEF,&h4669BE79,&hCB61B38C,&hBC66831A,&h256FD2A0,&h5268E236,&hCC0C7795,&hBB0B4703,&h220216B9,&h5505262F,&hC5BA3BBE,&hB2BD0B28,&h2BB45A92,&h5CB36A04,&hC2D7FFA7,&hB5D0CF31," _
		+ "&h2CD99E8B,&h5BDEAE1D,&h9B64C2B0,&hEC63F226,&h756AA39C,&h026D930A,&h9C0906A9,&hEB0E363F,&h72076785,&h05005713,&h95BF4A82,&hE2B87A14,&h7BB12BAE,&h0CB61B38,&h92D28E9B,&hE5D5BE0D,&h7CDCEFB7,&h0BDBDF21,&h86D3D2D4,&hF1D4E242,&h68DDB3F8,&h1FDA836E,&h81BE16CD,&hF6B9265B,&h6FB077E1,&h18B74777,&h88085AE6,&hFF0F6A70,&h66063BCA,&h11010B5C,&h8F659EFF,&hF862AE69,&h616BFFD3,&h166CCF45,&hA00AE278,&hD70DD2EE,&h4E048354,&h3903B3C2," _
		+ "&hA7672661,&hD06016F7,&h4969474D,&h3E6E77DB,&hAED16A4A,&hD9D65ADC,&h40DF0B66,&h37D83BF0,&hA9BCAE53,&hDEBB9EC5,&h47B2CF7F,&h30B5FFE9,&hBDBDF21C,&hCABAC28A,&h53B39330,&h24B4A3A6,&hBAD03605,&hCDD70693,&h54DE5729,&h23D967BF,&hB3667A2E,&hC4614AB8,&h5D681B02,&h2A6F2B94,&hB40BBE37,&hC30C8EA1,&h5A05DF1B,&h2D02EF8D"
	alCRCTable = Split(sValues, ",")
	lCRC = &hFFFFFFFF
	For l = 1 To Len(psString)
		lCRC = alCRCTable(((lCRC And &hFFFF) Xor Asc(Mid(psString, l, 1))) And &hFF) Xor shr(lCRC, 8)
	Next ' l
	lCRC = lCRC Xor &hFFFFFFFF
	
	CalculateCRC = lCRC
End Function









