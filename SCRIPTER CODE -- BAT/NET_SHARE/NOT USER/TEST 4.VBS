' ----
' VBS SCRIPT TO IP ADDRESS SCAN FIND ALL COMPUTER ON NETWORK - Google Search
' https://www.google.co.uk/search?q=VBS+SCRIPT+TO+IP+ADDRESS+SCAN+FIND+ALL+COMPUTER+ON+NETWORK&rlz=1C1CHBF_en-GBGB759GB759&oq=VBS+SCRIPT+TO+IP+ADDRESS+SCAN+FIND+ALL+COMPUTER+ON+NETWORK&aqs=chrome..69i57.16748j0j7&sourceid=chrome&ie=UTF-8
' ----
' ----
' Subnet Scan - remotely scan a subnet and identify resources - CodeProject
' https://www.codeproject.com/Articles/21363/Subnet-Scan-remotely-scan-a-subnet-and-identify-re
' ----

'*******************************************************************************************************************************
'*** Subnet Scan ©                                                                                                           ***
'***    Written by Frank Lindsey (See credits)                                                                               ***
'***       <a href="mailto:FVLindsey@HotMail.com">FVLindsey@HotMail.com</a> <a href="mailto:FVLindsey@GMail.com">FVLindsey@GMail.com</a> (PLEASE, NO SOLICITATIONS - I already have one of those)                ***
'*******************************************************************************************************************************
'*** Desription:                                                                                                             ***
'***    This script will scan your local subnet, or a user defined range of IP address, and return specific information for  ***
'***    each identifiable resource found.                                                                                    ***
'***                                                                                                                         ***
'*** Assumptions:                                                                                                            ***
'***    - WScript 5.6+                                                                                                       ***
'***    - Network access                                                                                                     ***
'***    - Premssions to access resources                                                                                     ***
'***                                                                                                                         ***
'*** Coding Rules:                                                                                                           ***
'***    - Variable names are in the format of <vartype><DescriptiveName>.                                                    ***
'***       - In <DescriptiveName> caps are used to deliniate words                                                           ***
'***       - <vartype>s are:                                                                                                 ***
'***            int   Integer                                                                                                ***
'***            str   String                                                                                                 ***
'***            obj   Object                                                                                                 ***
'***            col   Collection of objects                                                                                  ***
'***            ary   Array                                                                                                  ***
'***            l     logical                                                                                                ***
'***    - Constant names are all caps with an "_" used as word a seperator                                                   ***
'***    - Global names are preceeded with a "G"                                                                              ***
'***    - Subroutine names are in the format of sub<DescriptiveName>                                                         ***
'***       - In <DescriptiveName> caps are used to deliniate words                                                           ***
'***    - Function names are in the format of fun<DecsriptiveName>                                                           ***
'***       - In <DescriptiveName> caps are used to deliniate words                                                           ***
'***    - Code formated to be viewed at 128 columns. No tab characters, indent level is three spaces                         ***
'***    - Default output file is DEFAULT_PATH & <Month><Day><Year>_<Hour><Minute> & REPORT_TITLE & [i][s]                    ***
'***       - To modify the path, change the constant DEFAULT_PATH                                                            ***
'***       - To modifiy the file name, change the constant REPORT_TITLE                                                      ***
'***       - "i" and/or "s" appended to the file name depending on runtime parameters.                                       ***
'***                                                                                                                         ***
'*** Command Line Parameters:                                                                                                ***
'***    -input, -i, /input, /i   Allow user interaction for application parameters selection                                 ***
'***    -short, -s, /short, /s   Limit the number of fields returned to a predefined subset                                  ***
'***    <None>, <Invalid>        Standard method with default application parameters used (no -s nor -i)                     ***
'***                                                                                                                         ***
'***    If -input (etc.) is used the variable GlAskForInput is set to True                                                   ***
'***    If -short (etc.) is used the variable GlShortFormat is set to True                                                   ***
'***                                                                                                                         ***
'***  Revision History:                                                                                                      ***
'***    Orginal coding   10/ 1/2007   Frank Lindsey                                                                          ***
'***    Update           10/ 3/2007   Frank Lindsey   Added parameters, Local Admins test, and SMS check                     ***
'***    Update           10/ 4/2007   Frank Lindsey   Added MAC Address and merged similar queries                           ***
'***    Update           11/ 9/2007   Frank Lindsey   Reorganized code and added UPS detect                                  ***
'***    Release          11/19/2007   Frank Lindsey   Released into Public Domain                                            ***
'***                                                                                                                         ***
'*** Credit:                                                                                                                 ***
'***   Based on the script AssetScan.vbs - open source. HEAVILY modified.                                                    ***
'***      AssetScan.vbs - Query PC's on your network with WMI and log the responses into an excel spreadsheet. Works with    ***
'***                      Windows NT, 2K, XP. © Sean Kelly - <a href="mailto:skelly@engineer.com">skelly@engineer.com</a>. rev 12 April 2005                          ***
'***                                                                                                                         ***
'***   ShowBar() found on internet. If you created this fuction let me know so I can give credit where credit is due.        ***
'***   SMSStatus based on procedures pulled from get-and-set-sms-sidecode.vbs written by <a href="mailto:Tyson.Flint@Premera.com">Tyson Flint</a>.                        ***
'***   Thanks to <a href="mailto:Scotts.Rossow@Premera.com">Scotts Rossow</a> for beta testing and coding suggestions.                                                      ***
'***                                                                                                                         ***
'*******************************************************************************************************************************
'*** Public Domain                                                                                                           ***
'***    This script is released into the public domain. You may use it freely, however, if you make any modifications and    ***
'***    redistribute, please list your name and describe the changes.                                                        ***
'***                                                                                                                         ***
'***    This script is distributed without any warranty, expressed or implied. If you choose to use this script, in whole or ***
'***    in part, you agree to take sole responsibility for any problems that may occur. Please be aware that this script may ***
'***    cause network slowing depending on the resources available and/or the scan range selected.                           ***
'*******************************************************************************************************************************
'***[ Initializations ]*********************************************************************************************************
Option Explicit
'Verify we are running WScript before we do anything else
If (InStr(LCase(WScript.FullName),"wscript") = 0) Then Call subCheckScriptHost()
'*** Declarations **************************************************************************************************************
'**************************************************************
'** User Definable - reset these values to customize the script
'**************************************************************
   'Constants
   Const DEFAULT_PATH  = "C:\TEMP\" 'Save the file here
   Const REPORT_TITLE  = "Subnet Scan"  'This is the complete title
   Const LINE_HEADER   = "<br />---- " 'Used in the progress bar window
'**************************************************************
'***[ Gobal Variables ] Sorry, I hate them too. I may re-write and use classes to get around this ******************************
'Intergers
Dim GintRow 'Current row in spreadsheet
'Constants
Const FLAG_RETURN_IMMEDIATELY = &h10 : Const FLAG_FORWARD_ONLY = &h20 'Query operations related
'Boolean
Dim GlAskForInput, GlShortFormat 'Runtime parameters
'Arrays
Dim GaryIPRange(2) 'Three values; Subnet, Start, and End
'*********************************************************************
'** Create Global Objects - These objects are used throught the script
'*********************************************************************
   'Objects
   Dim GobjExcel : Set GobjExcel = Nothing
   Dim GobjIE    : Set GobjIE    = Nothing
   On Error Resume Next
      Err.Clear
      Set GobjExcel = WScript.CreateObject("Excel.Application")
      Set GobjIE    = WScript.CreateObject("InternetExplorer.Application")
   On Error Goto 0
   
   'Any Errors?
   If Err.Number <> 0 Then
      Call subCloseApp("Run time Error.", Err.Number, Err.Description, Err.Source)
   ElseIf (GobjExcel Is Nothing) Then
      Call subCloseApp("Fatal Error creating Excel object", Err.Number, Err.Description, Err.Source)
   ElseIf (GobjIE Is Nothing) Then
      Call subCloseApp("Fatal Error creating IE object", Err.Number, Err.Description, Err.Source)
   End If
'************************************************************
'** Determine execution format prior to defining the PC array
'************************************************************
   '*********************
   '** Local Variables **
   '*********************
   'Strings
   Dim strArgument
   'Objects
   Dim objArguments : Set objArguments = Nothing
   On Error Resume Next
      Err.Clear
      Set objArguments = WScript.Arguments
   On Error Goto 0
   'Any Errors?
   If Err.Number <> 0 Then
      Call subCloseApp("Run time Error.", Err.Number, Err.Description, Err.Source)
   ElseIf (GobjExcel Is Nothing) Then
      Call subCloseApp("Fatal Error creating Excel object", Err.Number, Err.Description, Err.Source)
   ElseIf (GobjIE Is Nothing) Then
      Call subCloseApp("Fatal Error creating IE object", Err.Number, Err.Description, Err.Source)
   End If
   '*** Default Global varaibles
   GlAskForInput = False : GlShortFormat = False
   'Check for command line parameters
   For Each strArgument In objArguments
      Select Case LCase(strArgument)
         Case "-input", "/input", "-i", "/i"
            GlAskForInput = True
         Case "-short", "/short", "-s", "/s"
            GlShortFormat = True
         Case Else
            MsgBox "Invalid Parameter: " & strArgument & ". Running in default mode.", vbInformation + vbOKOnly, REPORT_TITLE
      End Select 
   Next
'****************************************************************
'** Create Global report detail array based on runtime parameters
'****************************************************************
   Dim GintPC_IP, GintPC_Name, GintPC_Make, GintPC_Model, GintPC_Serial, GintPC_User
   If GlShortFormat Then
      GintPC_IP = 0 : GintPC_Name = 1 : GintPC_Make = 2 : GintPC_Model = 3 : GintPC_Serial = 4 : GintPC_User = 5
   Else
      Dim GintPC_Role, GintPC_MAC, GintPC_RAM, GintPC_OS, GintPC_BIOS, GintPC_CPU, GintPC_Speed, GintPC_Date, GintPC_Admins,  _
          GintPC_SMS, GintPC_C_Size, GintPC_C_Free, GintPC_D_Size, GintPC_D_Free, GintPC_E_Size, GintPC_E_Free, GintPC_NIC_1, _
          GintPC_NIC_2, GintPC_NIC_3, GintPC_NIC_4, GintPC_NIC_5
      GintPC_IP     =  0 : GintPC_Name   =  1 : GintPC_Role   =  2 : GintPC_Make   =  3 : GintPC_Model  =  4
      GintPC_MAC    =  5 : GintPC_Serial =  6 : GintPC_RAM    =  7 : GintPC_OS     =  8 : GintPC_BIOS   =  9
      GintPC_CPU    = 10 : GintPC_Speed  = 11 : GintPC_User   = 12 : GintPC_Date   = 13 : GintPC_Admins = 14
      GintPC_SMS    = 15 : GintPC_C_Size = 16 : GintPC_C_Free = 17 : GintPC_D_Size = 18 : GintPC_D_Free = 19
      GintPC_E_Size = 20 : GintPC_E_Free = 21 : GintPC_NIC_1  = 22 : GintPC_NIC_2  = 23 : GintPC_NIC_3  = 24
      GintPC_NIC_4  = 25 : GintPC_NIC_5  = 26
   End If
'***[ End of Gobal Variables ] *************************************************************************************************
'***[ MAIN ]********************************************************************************************************************
'*********************
'** Local Variables **
'*********************
'Strings
Dim strDefaultFile, strStart, strIPList
'*******************
'** Start of code **
'*******************
strDefaultFile = funSetThingsUp(strStart) 'Defines the output file name
If GlAskForInput Then If (MsgBox("Run System Auditor?", vbQuestion + vbYesNo, REPORT_TITLE) = vbNo) Then WScript.Quit 'Go/No Go
Call subShowBar() 'Draw the progress box
If GlAskForInput Then GobjIE.Document.ParentWindow.Document.Script.ListOP("Application Started: " & strStart)
strIPList = funIPCreate()  'Determine the scan range
Call subBuildXLS()         'Create the Excel spreadsheet for output
Call subConnect(strIPList) 'Connect to the system and retrieve data
Call subFooter()           'Create the subFooter on the spreadsheet
GobjIE.Quit 'Message Window cleanup
GobjExcel.Visible = True 'Show the output Excel file
'Save all of our work
If GlAskForInput Then strDefaultFile = strDefaultFile & "i" 'Add the i suffix
If GlShortFormat Then strDefaultFile = strDefaultFile & "s" 'Add the s suffix
If funSaveFiles(strDefaultFile) Then MsgBox "Your inventory run is complete!", vbInformation + vbOKOnly, REPORT_TITLE
'***[ End of MAIN ]*************************************************************************************************************
'*************
'** Cleanup **
'*************
'Object cleanup
Set GobjIE = Nothing : Set GobjExcel = Nothing
WScript.Quit 'Really and truely not necessary
'***[ SUBROUTINES ]*************************************************************************************************************
'***************************************************
'** subCheckScriptHost - Are we running WScript? ***
'***************************************************
Sub subCheckScriptHost()
   '*********************
   '** Local Variables **
   '*********************
   'Constants
   Const WINDOW_HIDE = 0 'Run Command Window Style
   'Objects
   Dim objShell : Set objShell = Nothing
   'Strings
   Dim strExec
   '*******************
   '** Start of code **
   '*******************
   'Create Objects
   On Error Resume Next
      Err.Clear
      Set objShell = CreateObject("WScript.Shell")
   On Error Goto 0
   'Any Errors?
   If Err.Number <> 0 Then
      Call subCloseApp("Run time Error.", Err.Number, Err.Description, Err.Source)
   ElseIf (objShell Is Nothing) Then
      Call subCloseApp("Fatal Error creating Shell object", Err.Number, Err.Description, Err.Source)
   End If
   'Restart using WScript
   strExec ="%COMSPEC% /c " & Chr(34) & "wscript.exe //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34) & Chr(34)
   objShell.Run strExec, WINDOW_HIDE, False
   Wscript.Quit
End Sub
'**************************************************************
'** subCloseApp - Called to abnormal application termination **
'**************************************************************
Sub subCloseApp(strError, intError, strDescription, strSource)
   On Error Resume Next 'No way out
      '*********************
      '** Local Variables **
      '*********************
      'Strings
      Dim strMessage 'Error message to be displayed
      '*************
      '** Cleanup **
      '*************
      GobjIE.Quit 'Message Window cleanup
      GobjExcel.Visible = True 'Show the output Excel file
      'Object cleanup
      Set GobjIE = Nothing : Set GobjExcel = Nothing
      strMessage = strError 'Start with the passed messsage
      'Add any error numbers
      If strError > 0 Then strMessage =  strMessage                              & vbCrLf & vbCrLf & _
                                         "*** UNRECOVERABLE ERROR: ABORTING ***" & vbCrLf          & _
                                         "*************************************" & vbCrLf          & _
                                         "  Error:        " & intError           & vbCrLf          & _
                                         "  Description : " & strDescription     & vbCrLf          & _
                                         "  Source:       " & strSource
      MsgBox strMessage, vbInformation + vbOKOnly, REPORT_TITLE
      WScript.Quit 'Abort, Abort, Abort
   On Error Goto 0 'Why? Too keep it clean looking
End Sub
'******************************************
'** subShowBar - Displays a progress bar **
'******************************************
Sub subShowBar()
   '*********************
   '** Local Variables **
   '*********************
   'Integers
   Dim intWindowWidth, intWindowHeight
   '*******************
   '** Start of code **
   '*******************
   On Error Resume Next
      Err.Clear
      GobjIE.Navigate("about:blank")
   On Error Goto 0
   'Any Errors?
   If Err.Number <> 0 Then Call subCloseApp("Error navigating to 'about:blank'.", Err.Number, Err.Description, Err.Source)
   'Do not continue until the page is ready
   Do : WScript.Sleep 50 : Loop Until GobjIE.ReadyState = 4
   With GobjIE.Document.ParentWindow.Screen
      intWindowHeight = .AvailHeight
      intWindowWidth  = .AvailWidth
   End With
   'HTML code
   With GobjIE
      .FullScreen = True
      .Toolbar    = False
      .StatusBar  = False
      .AddressBar = False
      .Resizable  = False
      .Width      = 420
      .Left       = (intWindowWidth  - 420) \ 2
       If GlAskForInput Then
         .Height = 270
         .Top    = (intWindowHeight - 270) \ 2
       Else
         .Height = 100
         .Top    = (intWindowHeight - 100) \ 2
       End If
      With .Document
         .WriteLN("<!DOCTYPE HTML PUBLIC>")
         '         *** Create the page
         .WriteLN("<HTML "                                          & _
                  "    Style         = ""border-style:outset;"      & _
                                        "border-width:4px"" "       & _
                      "OnKeyDown     = ""VBScript:SuppressKeys"" "  & _
                      "onHelp        = ""VBScript:SuppressIeFns"" " & _
                      "onContextMenu = ""VBScript:SuppressIeFns"">")
         .WriteLN(    "<HEAD>")
         .WriteLN(        "<TITLE>"         & _
                               REPORT_TITLE & _
                          "</TITLE>")
         .WriteLN(        "<STYLE "                             & _
                              "Type = ""text/css"">")
         .WriteLN(            "Body {background-color:#ece9d8;" & _
                                    "text-align:center;"        & _
                                    "vertical-align:middle}")
         .WriteLN(        "</STYLE>")
         '                 *** Add the VBScript code
         .WriteLN(        "<SCRIPT " & _
                             "Language = ""VBScript"">")
         '                     ****************************************************************************
         '                     ******** SuppressKeys - Ignore all keys execpt <Ctrl> which exits the window
         '                     ****************************************************************************
         .WriteLN(            "Function SuppressKeys()")
         .WriteLN(               "If NOT CBool(Window.Event.CTRLKey) Then")
         .WriteLN(                  "Exit Function")
         .WriteLN(               "End If")
         .WriteLN(               "Window.Event.KeyCode      = 0")
         .WriteLN(               "Window.Event.CancelBubble = True")
         .WriteLN(               "Window.Event.ReturnValue  = False")
         .WriteLN(            "End Function")
         '                     ********************************************************************
         '                     ******** SuppressIEFns - Stop all calls to Explorer window functions
         '                     ********************************************************************
         .WriteLN(            "Function SuppressIEFns()")
         .WriteLN(               "Window.Event.CancelBubble = True")
         .WriteLN(               "Window.Event.ReturnValue  = False")
         .WriteLN(            "End Function")
         '                     ************************************************
         '                     ******** BarOP - Incress the progress bar length
         '                     ************************************************
         .WriteLN(            "Function BarOP(intPercent)")
         .WriteLN(               "Window.BarArea.Style.Width = intPercent & ""%""")
         .WriteLN(            "End Function")
         If GlAskForInput Then
            '                  *****************************************************
            '                  ******** ListOP - Add new lines to the display window
            '                  *****************************************************
            .WriteLN(         "Function ListOP(strToInsert)")
            .WriteLN(            "Window.DataArea.InsertAdjacentHtml ""beforeBegin"", strToInsert")
            .WriteLN(            "Window.DataArea.ScrollIntoView")
            .WriteLN(         "End Function")
         End If
         .WriteLN(        "</SCRIPT>")
         .WriteLN(    "</HEAD>")
         '             *** Build the body of the window
         .WriteLN(    "<BODY " & _
                          "Scroll = ""No"">")
         .WriteLN(        "<TABLE>")
         .WriteLN(            "<TR>")
         .WriteLN(                "<TD Style = ""text-align:center;"                & _
                                                "font-family:Arial;font-size:16pt;" & _
                                                "font-weight:bold"">")
         .WriteLN(                     "Premera Blue Cross - Auditor")
         .WriteLN(                "</TD>")
         .WriteLN(            "</TR>")
         .WriteLN(            "<TR>")
         .WriteLN(                "<TD ID = ""barcell"" "                                & _
                                             "Style = ""width:400px;"                    & _
                                                       "padding-left:7px;"               & _
                                                       "padding-right:7px;"              & _
                                                       "text-align:left;"                & _
                                                       "border-style:inset;"             & _
                                                       "border-width:thin;"              & _
                                                       "background-color:navajowhite"">")
         .WriteLN(                    "<HR ID = ""BarArea"" "                       & _
                                                 "Style = ""width:0%;height:15px;" & _
                                                           "color:darkblue"" />")
         .WriteLN(                "</TD>")
         .WriteLN(            "</TR>")
         If GlAskForInput Then
            .WriteLN(         "<TR>")
            .WriteLN(             "<TD STYLE = ""padding-top:15px"">")
            .WriteLN(                           "<DIV ID = ""progresslist"" "                                & _
                                                            "Style = ""height:100px;width:380px;"            & _
                                                                      "max-height:100%;max-width:100%;"      & _
                                                                      "padding-left:10px;text-align:left;"   & _
                                                                      "font-family:Arial;font-size:10pt;"    & _
                                                                      "font-weight:bold;border-style:inset;" & _
                                                                      "border-width:thin;overflow:scroll"">")
            .WriteLN(                               "<SPAN "                 & _
                                                        "ID = ""DataArea"">" & _
                                                    "</SPAN>")
            .WriteLN(                           "</DIV>")
            .WriteLN(             "</TD>")
            .WriteLN(         "</TR>")
            .WriteLN(         "<TR>")
            .WriteLN(             "<TD STYLE = ""padding-top:20px;"              & _
                                                "width:400px;font-family:Arial;" & _
                                                "font-size:10pt;"                & _
                                                "font-weight:bold"">")
            .WriteLN(                  "Scanning for systems...")
            .WriteLN(             "</TD>")
            .WriteLN(         "</TR>")
         End If
         .WriteLN(        "</TABLE>")
         .WriteLN(    "</BODY>")
         .WriteLN("</HTML>")
      End With 
   .Visible = True 
   End With 
End Sub
'************************************************
'** subBuildXLS - Builds the actual Excel file **
'************************************************
Sub subBuildXLS()
   '***********************
   '*** Local Variables ***
   '***********************
   'Constants
   Const EXCEL_WHITE = 2 : Const EXCEL_BLUE = 11 : Const EXCEL_SOLID = 1 : Const EXCEL_LEFT = 2 : Const EXCEL_RIGHT = 4
   'Arrays
   Dim aryPCs()
   'Fill the headers for the PC data
   If GlShortFormat Then
      ReDim aryPCs(5)
      aryPCs(GintPC_IP)    = "IP Address" : aryPCs(GintPC_Name)   = "Hostname"      : aryPCs(GintPC_Make) = "Make"
      aryPCs(GintPC_Model) = "Model"      : aryPCs(GintPC_Serial) = "Serial Number" : aryPCs(GintPC_User) = "Logged User"
   Else
      ReDim aryPCs(26)
      aryPCs(GintPC_IP)     = "IP Address"       : aryPCs(GintPC_Name)   = "Hostname"
      aryPCs(GintPC_Role)   = "Role"             : aryPCs(GintPC_Make)   = "Make"
      aryPCs(GintPC_Model)  = "Model"            : aryPCs(GintPC_MAC)    = "MAC Address"
      aryPCs(GintPC_Serial) = "Serial Number"    : aryPCs(GintPC_RAM)    = "RAM"
      aryPCs(GintPC_OS)     = "Operation System" : aryPCs(GintPC_BIOS)   = "BIOS Revision"
      aryPCs(GintPC_CPU)    = "CPU Type"         : aryPCs(GintPC_Speed)  = "CPU Speed"
      aryPCs(GintPC_User)   = "Logged User"      : aryPCs(GintPC_Date)   = "Date Installed"
      aryPCs(GintPC_Admins) = "Local Admins"     : aryPCs(GintPC_SMS)    = "SMS Site"
      aryPCs(GintPC_C_Size) = "C: Size"          : aryPCs(GintPC_C_Free) = "C: Free"
      aryPCs(GintPC_D_Size) = "D: Size"          : aryPCs(GintPC_D_Free) = "D: Free"
      aryPCs(GintPC_E_Size) = "E: Size"          : aryPCs(GintPC_E_Free) = "E: Free"
      aryPCs(GintPC_NIC_1)  = "NIC #1"           : aryPCs(GintPC_NIC_2)  = "NIC #2"
      aryPCs(GintPC_NIC_3)  = "NIC #3"           : aryPCs(GintPC_NIC_4)  = "NIC #4"
      aryPCs(GintPC_NIC_5)  = "NIC #5"
   End If
   '*******************
   '** Start of code **
   '*******************
   GintRow = 1 'Current row in spreadsheet
   GobjExcel.Visible = False
   GobjExcel.WorkBooks.Add
   GobjExcel.Sheets("Sheet1").Select()
   GobjExcel.Sheets("Sheet1").Name = REPORT_TITLE
   GobjExcel.Rows(1).RowHeight = 25 'Set height of Title row
   'Set Cell Format for Column Titles
   If GlShortFormat Then
      GobjExcel.Range("A1:F1").Select
   Else
      GobjExcel.Range("A1:AA1").Select
   End If
   'Global settings on spreadsheet
   GobjExcel.Selection.Font.Size           = 8
   GobjExcel.Selection.Font.ColorIndex     = EXCEL_WHITE
   GobjExcel.Selection.Interior.ColorIndex = EXCEL_BLUE
   GobjExcel.Selection.Interior.Pattern    = EXCEL_SOLID
   GobjExcel.Selection.Font.Bold           = True
   GobjExcel.Selection.WrapText            = True
   If GlShortFormat Then
      GobjExcel.Range("A:F").HorizontalAlignment = EXCEL_LEFT
   Else
      GobjExcel.Range("A:AA").HorizontalAlignment = EXCEL_LEFT
      GobjExcel.Range("H:H" ).HorizontalAlignment = EXCEL_RIGHT
      GobjExcel.Range("L:L" ).HorizontalAlignment = EXCEL_RIGHT
      GobjExcel.Range("N:N" ).HorizontalAlignment = EXCEL_RIGHT
      GobjExcel.Range("Q:V" ).HorizontalAlignment = EXCEL_RIGHT
   End If
   Call subAddLineXLS(aryPCs)
End Sub
'**************************************************
'** subAddLineXLS - Add Lines to the spreadsheet **
'**************************************************
Sub subAddLineXLS(ByRef aryLineDetail)
   '*********************
   '** Local Variables **
   '*********************
   'Integers
   Dim intCounter, intRows 'Basic Counters
   '*******************
   '** Start of code **
   '*******************
   intRows = UBound(aryLineDetail) + 1 'Number of rows sent
   For intCounter = 1 To intRows
      GobjExcel.Cells(GintRow, intCounter).Value = Trim(aryLineDetail(intCounter - 1))
   Next
   GintRow = GintRow + 1 'We are now on the next row in the spreadsheet
   GobjExcel.Cells(1, 1).Select 'Back to the top
End Sub
'***************************************************
'** subFooter - added when speadsheet is complete **
'***************************************************
Sub subFooter()
   '*********************
   '** Local Variables **
   '*********************
   'Constants
   Const EXCEL_BLACK = 1 : Const EXCEL_LEFT = 2
   Const IP_SUBNET = 0 : Const IP_START = 1 : Const IP_END = 2
   'Integers
   Dim intCounter 'Basic Counter
   'Strings
   Dim strParameters
   'Arrays
   Dim aryFooters(2)
   '*******************
   '** Start of code **
   '*******************
   strParameters = "" 'Default it to empty string
   If GlShortFormat Then
      strParameters = "Short Format "
      GobjExcel.Range("A:F").ColumnWidth() = 40
      GobjExcel.Range("A:F").Columns.Autofit
   Else
      GobjExcel.Range("A:AA").ColumnWidth() = 40
      GobjExcel.Range("A:AA").Columns.Autofit
   End If
   If GlAskForInput Then strParameters = strParameters & "Interactive"
   aryFooters(0) = "Premera Blue Cross"
   aryFooters(1) = "Inventory AssetScan " & strParameters
   aryFooters(2) = "IP Range: " & GaryIPRange(IP_SUBNET) & "." & GaryIPRange(IP_START) & _
                   " through "  & GaryIPRange(IP_SUBNET) & "." & GaryIPRange(IP_END)
   GintRow = GintRow + 3 'Give us a little room
   For intCounter = 0 To 2
      GintRow = GintRow + 1
      GobjExcel.Cells(GintRow, 4).Select
      GobjExcel.Selection.Font.ColorIndex     = EXCEL_BLACK
      GobjExcel.Selection.Font.Size           = 8
      GobjExcel.Selection.Font.Bold           = False
      GobjExcel.Selection.HorizontalAlignment = EXCEL_LEFT
      GobjExcel.Cells(GintRow, 4).Value       = aryFooters(intCounter)
   Next
End Sub
'*********************************************
'** subConnect - Get Connect to each system **
'*********************************************
Sub subConnect(strAllIPs)
   '*********************
   '** Local Variables **
   '*********************
   'Constants
   Const ACCESS_DENIED = &H80041003 'Returned from a EWS call
   Const MAX_WAIT      = &H80 'connection timeout 120 seconds
   Const IP_SUBNET = 0 : Const IP_START = 1 : Const IP_END = 2
   'Strings
   Dim strTitle, strURL, strMessage, strPage, strTemp 'Work varialbles
   'Collection of Objects
   Dim colIPAddresses : Set colIPAddresses = Nothing
   Dim colItems       : Set colItems       = Nothing
   'Objects
   Dim objRegularExpression : Set objRegularExpression = New RegExp
   Dim objItem      : Set objItem      = Nothing 'Work object
   Dim objLocator   : Set objLocator   = Nothing
   Dim objSMSClient : Set objSMSClient = Nothing
   Dim objMSXML3    : Set objMSXML3    = Nothing
   'Intergers
   Dim intIPRange, intLoop, intItems
   
   'Strings
   Dim strRunCommand
   'Progress box variables
   Dim intPercentage, intOnLine
   If GlAskForInput Then Dim strResetLine, strResultLine
   'Logicals
   Dim lFoundIt
   'Arrays
   Dim aryPCs()
   '*******************
   '** Start of code **
   '*******************
   'Create Objects
   On Error Resume Next
      Err.Clear
      Set objLocator   = CreateObject("WbemScripting.SWbemLocator")
      Set objSMSClient = CreateObject("Microsoft.SMS.Client")
      Set objMSXML3    = CreateObject("MSXML2.ServerXMLHTTP")
   On Error Goto 0
   'Any Errors?
   If Err.Number <> 0 Then
      Call subCloseApp("Run time Error.", Err.Number, Err.Description, Err.Source)
   ElseIf (objLocator Is Nothing) Then
      Call subCloseApp("Fatal Error creating Locator object", Err.Number, Err.Description, Err.Source)
   ElseIf (objSMSClient Is Nothing) Then
      Call subCloseApp("Fatal Error creating SMS Client object", Err.Number, Err.Description, Err.Source)
   ElseIf (objMSXML3 Is Nothing) Then
      Call subCloseApp("Fatal Error creating XML object", Err.Number, Err.Description, Err.Source)
   End If
   'IP Range
   intIPRange = GaryIPRange(IP_END) - GaryIPRange(IP_START)
   intOnLine  = 0
   'Extract the names of each system
   With objRegularExpression
      .Pattern    = "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$"
      .IgnoreCase = True
      .Multiline  = True
      .Global     = True
   End With
   Set colIPAddresses = objRegularExpression.Execute(strAllIPs)
   intItems = colIPAddresses.Count
   'Loop through each name extracted
   For intLoop = 0 To intItems - 1
      strMessage = "" : strTitle = "" : strTemp = "" : strURL = "" : strPage = ""
      lFoundIt = False
      Err.Clear 'Nothing pending...
      'Get an empty array
      Erase aryPCs
      If GlShortFormat Then ReDim aryPCs(5) Else ReDim aryPCs(26)
      'OK, lets go...
      aryPCs(GintPC_IP) = colIPAddresses.Item(intLoop).Value 'Get the first IP Address
      intOnLine = intOnLine + 1 'Increase Line Counter
      'Update progress window
      intPercentage = CInt((intOnLine / (intIPRange + 1)) * 100)
      GobjIE.Document.ParentWindow.Document.Script.BarOP(intPercentage) 'Progress bar
      If GlAskForInput Then
         strResultLine = "<br />" & aryPCs(GintPC_IP) & LINE_HEADER & "Searching..."
         strResetLine  = funUpdateWindow(intPercentage, strResultLine, strResetLine)
      End If
      'Try a fast ping
      If funConnectable(aryPCs(GintPC_IP), 1, 250) Then
         'Object
         Dim objWMI : Set objWMI = Nothing
         On Error Resume Next
            Err.Clear
            Set objWMI = objLocator.ConnectServer(aryPCs(GintPC_IP), "root\cimv2",,,,, MAX_WAIT)
         On Error Goto 0
         'Not a PC
         If (objWMI Is Nothing) Then
            Err.Clear 'Nothing pending...
            If GlAskForInput Then
               strResultLine = LINE_HEADER & "Checking Network..."
               GobjIE.Document.ParentWindow.Document.Script.ListOP(strResultLine)
            End If
            aryPCs(GintPC_ROLE) = "Network Device"
            'Check for an embedded web server
            strURL = "http://" & aryPCs(GintPC_IP)
            On Error Resume Next 
               Err.Clear
               objMSXML3.open "GET", strURL, False
               objMSXML3.setRequestHeader "User-Agent","My funky browser." 
               objMSXML3.send ""  
               If objMSXML3.readyState <> 4 Then objMSXML3.waitForResponse 5 'Wait for a response
            On Error Goto 0
            'OK, now do we have anything?
            If Err.Number = 0            And _
               objMSXML3.readyState <> 1     Then
               strPage = LCase(objMSXML3.responseText) 'Get source for the entire webpage
               'Did we geat a response and no errors
               If Err.Number = 0 Then
                   lFoundIt = True 'We found something
                  'We found an Embedded Web Server (EWS), now lets try to get more details
                  '*****************************************************
                  '*** Place all the various test for different EWS here
                  '*****************************************************
                  If funCheckILo(strPage, aryPCs) Then
                     strMessage = "ILo Found"
                  ElseIf funCheckHP(strPage, aryPCs) Then
                     strMessage = "HP Device Found"
                  ElseIf funCheckAPC(strPage, aryPCs) Then
                     strMessage = "APC Device Found"
                  Else
                     'I give up, but there is an Embedded Web Server
                     strMessage          = "Unknown EWS!"
                     aryPCs(GintPC_Name) = strMessage
                  End If
               ElseIf Err.Number = 424 Then
                  aryPCs(GintPC_Name)  = "Unknown EWS!"
                  strMessage           = "Access Denied!"
                  aryPCs(GintPC_Model) = strMessage
               End If
            ElseIf Err.Number = ACCESS_DENIED Then
               aryPCs(GintPC_Name)  = "Unknown EWS!"
               strMessage           = "Access Denied!"
               aryPCs(GintPC_Model) = strMessage
            Else
               strMessage = "Nothing found."
            End If
            strResultLine = LINE_HEADER & strMessage
         Else 'No Embedded Web Server, it must be a PC
            If GlAskForInput Then
               strResultLine = LINE_HEADER & "Connected..."
               GobjIE.Document.ParentWindow.Document.Script.ListOP(strResultLine)
            End If
            lFoundIt = True 'We found something
            Call subGetHostname(       aryPCs, objWMI) 'Get Hostname
            Call subGetRoleUser(       aryPCs, objWMI) 'Get Domain Role and User Name
            Call subGetSerialMakeModel(aryPCs, objWMI) 'Get the Serial, Make, and Model
            'Are we doing the long format?
            If Not GlShortFormat Then
               Call subGetRAM(        aryPCs, objWMI) 'Get RAM (Total)
               Call subGetDateOS(     aryPCs, objWMI) 'Get Install Date and OS Version
               Call subGetBIOS(       aryPCs, objWMI) 'Get the BIOS value
               Call subGetCPUSpeed(   aryPCs, objWMI) 'Get the CPU Type and Speed
               Call subGetNICsInfo(   aryPCs, objWMI) 'Get NICs Details
               Call subGetDiskInfo(   aryPCs, objWMI) 'Get complete disk drive details
               Call subGetLocalAdmins(aryPCs)         'Get Local Administrators
               aryPCs(GintPC_SMS) = objSMSClient.AutoDiscoverSite 'Get status of SMS
            End If
            strResultLine = LINE_HEADER & aryPCs(GintPC_Role) & " Processed."
         End If
      Else
         strMessage = "Ping failed."
         strResultLine = LINE_HEADER & strMessage
      End If
      If lFoundIt Then Call subAddLineXLS(aryPCs) 'Is there any data to write
      If GlAskForInput Then GobjIE.Document.ParentWindow.Document.Script.ListOP(strResultLine)
      Set objWMI = Nothing
   Next
   'Clean things up
   Set colItems = Nothing : Set objMSXML3 = Nothing : Set objSMSClient = Nothing : Set objItem = Nothing
End Sub
'*******************************************
'** subGetHostname - Get the PCs Hostname **
'*******************************************
Sub subGetHostname(ByRef aryPCs, ByRef objWMI)
   '*********************
   '** Local Variables **
   '*********************
   'Collections
   Dim colItems : Set colItems = Nothing
   'Objects
   Dim objItem : Set objItem = Nothing
   '*******************
   '** Start of code **
   '*******************
   Set colItems = objWMI.ExecQuery("SELECT DNSHostName, MACAddress"             & _
                                   "   FROM Win32_NetworkAdapterConfiguration"  & _
                                   "   WHERE IPEnabled = TRUE",                   _
                                   "WQL",                                         _
                                   FLAG_RETURN_IMMEDIATELY + FLAG_FORWARD_ONLY)
   For Each objItem In colItems
      aryPCs(GintPC_Name) = objItem.DNSHostName
      aryPCs(GintPC_MAC)  = objItem.MACAddress
   Next
End Sub
'*****************************************************
'** subGetRoleUser - Get the PCs Role and User Name **
'*****************************************************
Sub subGetRoleUser(ByRef aryPCs, ByRef objWMI)
   '*********************
   '** Local Variables **
   '*********************
   'Collection of Objects
   Dim colItems : Set colItems = Nothing
   'Objects
   Dim objItem : Set objItem = Nothing
   '*******************
   '** Start of code **
   '*******************
   Set colItems = objWMI.ExecQuery("SELECT DomainRole, UserName"                & _
                                   "   FROM Win32_ComputerSystem",                _
                                   "WQL",                                         _
                                   FLAG_RETURN_IMMEDIATELY + FLAG_FORWARD_ONLY)
   For Each objItem In colItems
      'How is the Role of the system defined?
      Select Case objItem.DomainRole
         Case 0
            aryPCs(GintPC_Role) = "Standalone Workstation"
         Case 1
            aryPCs(GintPC_Role) = "Workstation"
         Case 2
            aryPCs(GintPC_Role) = "Standalone Server"
         Case 3
            aryPCs(GintPC_Role) = "Server"
         Case 4
            aryPCs(GintPC_Role) = "Backup DC"
         Case 5
            aryPCs(GintPC_Role) = "Primary DC"
         Case Else
            aryPCs(GintPC_Role) = "Unknown System Role"
      End Select
      'Who is logged in currently
      aryPCs(GintPC_User) = objItem.UserName
   Next
End Sub
'************************************************************************
'** subGetSerialMakeModel - Get the PCs Serial Number, Make, and Model **
'************************************************************************
Sub subGetSerialMakeModel(ByRef aryPCs, ByRef objWMI)
   '*********************
   '** Local Variables **
   '*********************
   'Collection of Objects
   Dim colItems : Set colItems = Nothing
   'Objects
   Dim objItem : Set objItem = Nothing
   '*******************
   '** Start of code **
   '*******************
   Set colItems = objWMI.ExecQuery("SELECT IdentifyingNumber, Name, Vendor"     & _
                                   "   FROM Win32_ComputerSystemProduct",         _
                                   "WQL",                                         _
                                   FLAG_RETURN_IMMEDIATELY + FLAG_FORWARD_ONLY)
   For Each objItem In colItems
      aryPCs(GintPC_Serial) = objItem.IdentifyingNumber
      aryPCs(GintPC_Make)   = objItem.Vendor
      aryPCs(GintPC_Model)  = objItem.Name
   Next
End Sub
'***************************************
'** subGetRAM - Get the PCs total RAM **
'***************************************
Sub subGetRAM(ByRef aryPCs, ByRef objWMI)
   '*********************
   '** Local Variables **
   '*********************
   'Collection of Objects
   Dim colItems : Set colItems = Nothing
   'Objects
   Dim objItem : Set objItem = Nothing
   '*******************
   '** Start of code **
   '*******************
   Set colItems = objWMI.ExecQuery("SELECT TotalPhysicalMemory"                 & _
                                   "   FROM Win32_LogicalMemoryConfiguration",    _
                                   "WQL",                                         _
                                   FLAG_RETURN_IMMEDIATELY + FLAG_FORWARD_ONLY)
   For Each objItem In colItems : aryPCs(GintPC_RAM) = funSizeFormat(objItem.TotalPhysicalMemory, "KB", "") : Next
End Sub
'************************************************************
'** subGetDateOS - Get the PCs Install Date and OS Version **
'************************************************************
Sub subGetDateOS(ByRef aryPCs, ByRef objWMI)
   '*********************
   '** Local Variables **
   '*********************
   'Strings
   Dim strTemp
   'Collection of Objects
   Dim colItems : Set colItems = Nothing
   'Objects
   Dim objItem : Set objItem = Nothing
   '*******************
   '** Start of code **
   '*******************
   Set colItems = objWMI.ExecQuery("SELECT Caption, CSDVersion, InstallDate"    & _
                                   "   FROM Win32_OperatingSystem",               _
                                   "WQL",                                         _
                                   FLAG_RETURN_IMMEDIATELY + FLAG_FORWARD_ONLY)
   For Each objItem In colItems
      strTemp = Left(objItem.InstallDate, 8) 'Work string
      'Format Date
      aryPCs(GintPC_Date) = Mid(strTemp, 3, 2) & "/" & Right(strTemp, 2) & "/" & Left(objItem.InstallDate, 4)
      'Shorten Service Pack
      aryPCs(GintPC_OS) = Trim(objItem.Caption) & Replace(objItem.CSDVersion, "Service Pack ", " (SP ") & ")"
      'Clean up unwanted text
      If InStr(aryPCs(GintPC_OS), "Microsoft Windows ") <> 0 Then
         aryPCs(GintPC_OS) = Replace(aryPCs(GintPC_OS), "Microsoft Windows ", "")
      ElseIf InStr(aryPCs(GintPC_OS), "Microsoft(R) Windows(R) ") <> 0 Then
         aryPCs(GintPC_OS) = Replace(aryPCs(GintPC_OS), "Microsoft(R) Windows(R) ", "")
      End If
      'Shorten type description
      If InStr(aryPCs(GintPC_OS), "Professional") <> 0 Then
         aryPCs(GintPC_OS) = Replace(aryPCs(GintPC_OS), "Professional", "PRO")
      ElseIf InStr(aryPCs(GintPC_OS), "Standard Edition") <> 0 Then
         aryPCs(GintPC_OS) = Replace(aryPCs(GintPC_OS), "Standard Edition", "SE")
      End If
   Next
End Sub
'***********************************************
'** subGetBIOS - Get the PCs BIOS information **
'***********************************************
Sub subGetBIOS(ByRef aryPCs, ByRef objWMI)
   '*********************
   '** Local Variables **
   '*********************
   'Collection of Objects
   Dim colItems : Set colItems = Nothing
   'Objects
   Dim objItem : Set objItem = Nothing
   '*******************
   '** Start of code **
   '*******************
   Set colItems = objWMI.ExecQuery("SELECT Version"                            & _
                                   "   FROM Win32_BIOS",                         _
                                   "WQL",                                        _
                                   FLAG_RETURN_IMMEDIATELY + FLAG_FORWARD_ONLY)
   For Each objItem In colItems : aryPCs(GintPC_BIOS) = objItem.Version : Next
End Sub
'*****************************************************
'** subGetCPUSpeed - Get the PCs CPU type and Speed **
'*****************************************************
Sub subGetCPUSpeed(ByRef aryPCs, ByRef objWMI)
   '*********************
   '** Local Variables **
   '*********************
   'Constants
   Const KB = 1024 : Const MB = 1048576 : Const GB = 1073741824 'Metric values
   'Collection of Objects
   Dim colItems : Set colItems = Nothing
   'Objects
   Dim objItem : Set objItem = Nothing
   '*******************
   '** Start of code **
   '*******************
   On Error Resume Next
      Err.Clear
      Set colItems = objWMI.ExecQuery("SELECT Name, MaxClockSpeed, Description"    & _
                                      "   FROM Win32_Processor",                     _
                                      "WQL",                                         _
                                      FLAG_RETURN_IMMEDIATELY + FLAG_FORWARD_ONLY)
   If Err.Number = 0 Then
      For Each objItem In colItems
         aryPCs(GintPC_CPU)  = Left(objItem.Name, InStr(objItem.Name, " CPU ") - 1)
         aryPCs(GintPC_Speed) = FormatNumber(objItem.MaxClockSpeed / KB) & " GHz"
      Next
   End If
End Sub
'***************************************************
'** subGetNICsInfo - Get the PCs NICs inforamtion **
'***************************************************
Sub subGetNICsInfo(ByRef aryPCs, ByRef objWMI)
   '*********************
   '** Local Variables **
   '*********************
   'Integers
   Dim intCounter
   'Collection of Objects
   Dim colItems : Set colItems = Nothing
   'Objects
   Dim objItem : Set objItem = Nothing
   '*******************
   '** Start of code **
   '*******************
   Set colItems = objWMI.ExecQuery("SELECT Name, AdapterType"                   & _
                                   "   FROM Win32_NetworkAdapter",                _
                                   "WQL",                                         _
                                   FLAG_RETURN_IMMEDIATELY + FLAG_FORWARD_ONLY)
   intCounter = 0
   For Each objItem In colItems
      If objItem.AdapterType = "Ethernet 802.3" Then
         Select Case intCounter
            Case 1
               aryPCs(GintPC_NIC_1) = objItem.Name
            Case 2
               aryPCs(GintPC_NIC_2) = objItem.Name
            Case 3
               aryPCs(GintPC_NIC_3) = objItem.Name
            Case 4
               aryPCs(GintPC_NIC_4) = objItem.Name
            Case 5
               aryPCs(GintPC_NIC_5) = objItem.Name
         End Select
         intCounter = intCounter + 1
      End If
   Next
End Sub
'***************************************************
'** subGetDiskInfo - Get the PCs Disk Information **
'***************************************************
Sub subGetDiskInfo(ByRef aryPCs, ByRef objWMI)
   '*********************
   '** Local Variables **
   '*********************
   'Integers
   Dim intCounter
   'Collection of Objects
   Dim colItems : Set colItems = Nothing
   'Objects
   Dim objItem : Set objItem = Nothing
   '*******************
   '** Start of code **
   '*******************
   Set colItems = objWMI.ExecQuery("SELECT DeviceID, Size, FreeSpace"           & _
                                   "   FROM Win32_LogicalDisk"                  & _
                                   "   WHERE DriveType = '3'",                    _
                                   "WQL",                                         _
                                   FLAG_RETURN_IMMEDIATELY + FLAG_FORWARD_ONLY)
   intCounter = 0
   For Each objItem In colItems
      Select Case UCase(objItem.DeviceID)
         Case "C:"
            aryPCs(GintPC_C_Size) = funSizeFormat(objItem.Size     , "BY", "GB")
            aryPCs(GintPC_C_Free) = funSizeFormat(objItem.FreeSpace, "BY", "GB")
         Case "D:"
            aryPCs(GintPC_D_Size) = funSizeFormat(objItem.Size     , "BY", "GB")
            aryPCs(GintPC_D_Free) = funSizeFormat(objItem.FreeSpace, "BY", "GB")
         Case "E:"
            aryPCs(GintPC_E_Size) = funSizeFormat(objItem.Size     , "BY", "GB")
            aryPCs(GintPC_E_Free) = funSizeFormat(objItem.FreeSpace, "BY", "GB")
      End Select
      intCounter = intCounter + 1
      'Only consider the first three hard drives
      If intCounter > 2 Then Exit For
   Next
End Sub
'*******************************************************************
'** subGetLocalAdmins - Identify the local machine Administrators **
'*******************************************************************
Sub subGetLocalAdmins(ByRef aryPCs)
   '*********************
   '** Local Variables **
   '*********************
   'Strings
   Dim strTemp, strToReturn, strLineHeader
   'Collection of Objects
   Dim colItems : Set colItems = Nothing
   'Objects
   Dim objItem  : Set objItem = Nothing
   Dim objWinNT : Set objItem = Nothing
   '*******************
   '** Start of code **
   '*******************
   On Error Resume Next
      Err.Clear
      Set objWinNT = GetObject("WinNT://" & aryPCs(GintPC_IP))
   On Error Goto 0
   'Any Errors?
   If Err.Number <> 0 Then
      Call subCloseApp("Run time Error.", Err.Number, Err.Description, Err.Source)
   ElseIf (objWinNT Is Nothing) Then
      Call subCloseApp("Fatal Error creating Win NT object", Err.Number, Err.Description, Err.Source)
   End If
   'Default variables
   strToReturn = "" : strLineHeader = ""
   'Any Errors?
   If Err.Number = 0            And _
      Not (objWinNT Is Nothing)     Then
      'Read in the local system info
      On Error Resume Next
      Err.Clear
         objWinNT.GetInfo
      On Error Goto 0
      If Err.Number = 0 And _
         objWinNT.PropertyCount > 0 Then
         On Error Resume Next
            Err.Clear
            Set colItems = GetObject("WinNT://" & aryPCs(GintPC_IP) & "/Administrators,group")
         On Error Goto 0
         If Err.Number = 0             And _
            Not (objWinNT Is Nothing)  And _
            colItems.PropertyCount > 0     Then
            For Each objItem In colItems.Members
               strTemp = Right(objItem.adsPath, Len(objItem.adsPath) - 8)
               'Ignore special accounts and we know about oa0ad01
               If InStr(strTemp, "/")       <> 0 And _
                  InStr(strTemp, "$")       =  0 And _
                  InStr(strTemp, " ")       =  0 And _
                  InStr(strTemp, "oa0ad01") =  0     Then
                  strToReturn   = strLineHeader & strTemp
                  strLineHeader = strToReturn & ", " 
               End If
            Next
         End If
      End If
   End If
   aryPCs(GintPC_Admins) = strToReturn 'Return results
End Sub
'***[ End of SUBROUTINES ]******************************************************************************************************
'***[ FUNCTIONS ]***************************************************************************************************************
'***********************************************
'** funSetThingsUp - Initial startup routines **
'***********************************************
Function funSetThingsUp(ByRef strStart)
   '*********************
   '** Local Variables **
   '*********************
   'Date Time
   Dim dteToday, dteNow
   'Strings
   Dim strFile
   '*******************
   '** Start of code **
   '*******************
   'Create the default filename
   dteToday = Date()
   dteNow   = Time()
   strStart = "Inventory run started: " & dteToday & " at " &  dteNow 'Used in Footer
   strFile  = Right("0000" & Year(dteToday) , 2) & _
              Right("00"   & Month(dteToday), 2) & _
              Right("00"   & Day(dteToday)  , 2) & _
              "_"                                & _
              Right("00"   & Hour(dteNow)   , 2) & _
              Right("00"   & Minute(dteNow) , 2) & REPORT_TITLE
   funSetThingsUp = strFile
End Function
'***********************************
'** funIPCreate - Create IP table **
'***********************************
Function funIPCreate()
   '*********************
   '** Local Variables **
   '*********************
   'Constants
   Const IP_BOTTOM = 0 : Const IP_TOP = 255 'Default IP range limits
   Const IP_SUBNET = 0 : Const IP_START = 1 : Const IP_END = 2
   'Integers
   Dim intCounter
   'Strings
   Dim strIPList, strCurrentIP, strLineHeader
   'Collection of Objects
   Dim colTemp : Set colTemp = Nothing
   'Objects
   Dim objRegularExpression : Set objRegularExpression = New RegExp
   '*******************
   '** Start of code **
   '*******************
   'Default variables
   strIPList = ""
   'Get subnet to scan
   strCurrentIP = funGetIP()
   'Break out the subnet
   With objRegularExpression
      .Pattern    = "(\d{1,3}\.\d{1,3}\.\d{1,3})\.\d{1,3}"
      .IgnoreCase = True
      .Multiline  = True
      .Global     = False
   End With
   Set colTemp = objRegularExpression.Execute(strCurrentIP)
   If colTemp.Count > 0 Then
      GaryIPRange(IP_SUBNET) = colTemp.Item(0).Submatches(0)
   End If
   If GlAskForInput Then
      'Verify subnet
      GaryIPRange(IP_SUBNET) = InputBox ("Enter Subnet to Scan - <enter> for Local Subnet", REPORT_TITLE, GaryIPRange(IP_SUBNET))
      'Verify IP range
      GaryIPRange(IP_START) = InputBox ("Start at :", "Scanning Subnet: " & GaryIPRange(IP_SUBNET), IP_BOTTOM)
      GaryIPRange(IP_END)   = InputBox ("  End at :", "Scanning Subnet: " & GaryIPRange(IP_SUBNET), IP_TOP)
   Else
      GaryIPRange(IP_START) = IP_BOTTOM
      GaryIPRange(IP_END)   = IP_TOP
   End If
   'Write IP address to string
   strLineHeader = ""
   For intCounter = GaryIPRange(IP_START) To GaryIPRange(IP_END)
      strIPList = strLineHeader & GaryIPRange(IP_SUBNET) & "." & intCounter 'Append the new address
      strLineHeader = strIPList & vbCrLf 'New header Line
   Next
   funIPCreate = strIPList 'Return the entire IP List
   'Cleanup
   Set colTemp = Nothing : Set objRegularExpression = Nothing
End Function
'******************************************************
'** funGetIP - Get the IP address of the Host system **
'******************************************************
Function funGetIP()
   '*********************
   '** Local Variables **
   '*********************
   'Constants
   Const MAX_WAIT      = &H80 'connection timeout 120 seconds
   'Integers
   Dim intCounter
   'Strings
   Dim strIPAddress : strIPAddress = "0.0.0.0" 'Default it
   'Collection of Objects
   Dim colItems : Set colItems = Nothing
   'Objects
   Dim objItem    : Set objItem    = Nothing 'Work object
   Dim objLocator : Set objLocator = Nothing
   Dim objWMI     : Set objWMI     = Nothing
   '*******************
   '** Start of code **
   '*******************
   On Error Resume Next
      Err.Clear
      Set objLocator = CreateObject("WbemScripting.SWbemLocator")
      Set objWMI = objLocator.ConnectServer(".", "root\cimv2",,,,, MAX_WAIT)
   On Error Goto 0
   'Any Errors?
   If Err.Number <> 0 Then
      Call subCloseApp("Run time Error.", Err.Number, Err.Description, Err.Source)
   ElseIf (objLocator Is Nothing) Then
      Call subCloseApp("Fatal Error creating Locator object", Err.Number, Err.Description, Err.Source)
   ElseIf (objWMI Is Nothing) Then
      Call subCloseApp("Fatal Error creating WMI object", Err.Number, Err.Description, Err.Source)
   End If
   'Default variables
   strIPAddress = ""
   If Err.Number = 0          And _
      Not (objWMI Is Nothing)     Then
      Set colItems = objWMI.ExecQuery("SELECT * "                                   & _
                                      "   FROM  Win32_NetworkAdapterConfiguration " & _
                                      "   WHERE IPEnabled = TRUE")
      'Returns a IP Address for each enabled network card
      For Each objItem in colItems
         If Not IsNull(objItem.IPAddress) Then 
            For intCounter = LBound(objItem.IPAddress) To UBound(objItem.IPAddress)
               strIPAddress = objItem.IPAddress(intCounter) 'We got it!
            Next
         End If
      Next
   End If
   funGetIP = strIPAddress 'Return Results
End Function
'***********************************************************
'** funConnectable - Try to PING a network address / name **
'***********************************************************
Function funConnectable(strHostName, intCount, intTimeOut)
   '*********************
   '** Local Variables **
   '*********************
   'Strings
   Dim strRunCommand
   'Integers
   Dim intReplyTotal, intFailedAttempts, intTestResult
   'Logicals
   Dim lReplyValue
   'Objects
   Dim objShell : Set objShell = Nothing
   '*******************
   '** Start of code **
   '*******************
   On Error Resume Next
      Err.Clear
      Set objShell = CreateObject("WScript.Shell") 
   On Error Goto 0
   'Any Errors?
   If Err.Number <> 0 Then
      Call subCloseApp("Run time Error.", Err.Number, Err.Description, Err.Source)
   ElseIf (objShell Is Nothing) Then
      Call subCloseApp("Fatal Error creating Shell object", Err.Number, Err.Description, Err.Source)
   End If
   'Default the Parameters
   If VarType(strHostName) = vbString Then
      If intCount   = vbEmpty Then intCount   = 2
      If intTimeOut = vbEmpty Then intTimeOut = 750
      'Ping the system.  Will return 0 on success and 1 on failure
      strRunCommand = "%ComSpec% /c "                     & _
                         "%SystemRoot%\system32\PING.EXE" & _
                            " -n " & intCount             & _
                            " -w " & intTimeout & " "     & _
                            strHostName
      Do Until (intReplyTotal = 2) Or (intFailedAttempts = 4)
         lReplyValue = objShell.Run(strRunCommand, 0, True)
         If (lReplyValue = 0) Then intReplyTotal = intReplyTotal + 1 Else intFailedAttempts = intFailedAttempts + 1
      Loop
      funConnectable = Not (intFailedAttempts = 4)
   Else
      funConnectable = False 'Invalid Parameter
   End If
End Function
'*******************************************
'** funCheckILo - Did we find an ILo EWS? **
'*******************************************
Function funCheckILo(strToTest, ByRef aryPCs)
   '*********************
   '** Local Variables **
   '*********************
   'Collection of Objects
   Dim colTemp : Set colTemp = Nothing
   'Objects
   Dim objRegularExpression : Set objRegularExpression = New RegExp
   '*******************
   '** Start of code **
   '*******************
   funCheckILo = False 'Default to failed
   If InStr(strToTest, "integrated lights") > 0 Then
      With objRegularExpression
         .Pattern    = "servername=[\\""]+([^;\\""]*)"
         .IgnoreCase = True
         .Multiline  = True
         .Global     = False
      End With
      Set colTemp = objRegularExpression.Execute(strToTest)
      If colTemp.Count > 0 Then
         aryPCs(GintPC_Name) = colTemp.Item(0).Submatches(0)
         aryPCs(GintPC_MAKE) = "ILo"
      End If
      funCheckILo = True 'We found something
   End If
End Function
'*******************************************
'** funCheckHP - Did we find an HP Device **
'*******************************************
Function funCheckHP(strToTest, ByRef aryPCs)
   '*********************
   '** Local Variables **
   '*********************
   'Collections
   Dim colTemp : Set colTemp = Nothing
   'Objects
   Dim objRegularExpression : Set objRegularExpression = New RegExp
   '*******************
   '** Start of code **
   '*******************
   funCheckHP = False 'Default to failed
   If InStr(strToTest, "hp ") > 0 Then
      'Get the role of the device
      aryPCs(GintPC_MAKE)  = "HP"
      aryPCs(GintPC_MODEL) = "Unknown"
      'Get the model number
      If InStr(strToTest, "sender") > 0 Then
         aryPCs(GintPC_ROLE) = "Digital Sender"
         With objRegularExpression
            .Pattern    = "9[0-9]00[^a-z]*" '9100c, 9200c - known models
            .IgnoreCase = True
            .Multiline  = True
            .Global     = False
         End With
         Set colTemp = objRegularExpression.Execute(strToTest)
         If colTemp.Count > 0 Then aryPCs(GintPC_MODEL) = colTemp.Item(0).Value
      Else
         aryPCs(GintPC_ROLE) = "LaserJet"
         Set objRegularExpression = New RegExp
         With objRegularExpression
            .Pattern    = "laserjet (\w+)"
            .IgnoreCase = True
            .Multiline  = True
            .Global     = False
         End With
         Set colTemp = objRegularExpression.Execute(strToTest)
         If colTemp.Count > 0 Then aryPCs(GintPC_MODEL) = colTemp.Item(0).Submatches(0)
      End If
      funCheckHP = True 'We found something
   End If
End Function
'*********************************************
'** funCheckAPC - Did we find an APC Device **
'*********************************************
Function funCheckAPC(strToTest, ByRef aryPCs)
   '*********************
   '** Local Variables **
   '*********************
   'Collections
   Dim colTemp : Set colTemp = Nothing
   'Objects
   Dim objRegularExpression : Set objRegularExpression = New RegExp
   '*******************
   '** Start of code **
   '*******************
   funCheckAPC = False 'Default to failed
   If InStr(strToTest, " apc ") > 0 Then
      'Get the role of the device
      aryPCs(GintPC_ROLE)  = "APC"
      aryPCs(GintPC_MAKE)  = "UPS"
      aryPCs(GintPC_MODEL) = "Unknown"
      funCheckAPC = True 'We found something
   End If
End Function
'************************************
'** funSaveFiles - Create IP table **
'************************************
Function funSaveFiles(strFileName)
   '*********************
   '** Local Variables **
   '*********************
   'Strings
   Dim strFilter, strTitle, strFullName
   'Objects
   Dim objFSO : Set objFSO = Nothing
   '*******************
   '** Start of code **
   '*******************
   On Error Resume Next
      Err.Clear
      Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
   On Error Goto 0
   If Err.Number <> 0 Then
      Call subCloseApp("Run time Error.", Err.Number, Err.Description, Err.Source)
   ElseIf (objFSO Is Nothing) Then
      Call subCloseApp("Fatal Error creating FSO object", Err.Number, Err.Description, Err.Source)
   End If
   funSaveFiles = False 'Default to fail
   'Configure Save As dialog box
   strFilter   = "Excel File (*.xls), *.xls"
   strTitle    = "Save As"
   strFullName = DEFAULT_PATH & strFileName & ".xls"
   'Start with a clean slate
   Err.Clear
   'Create the folder if it does not exist
   If Not (objFSO.FolderExists(DEFAULT_PATH)) Then
      On Error Resume Next
         Err.Clear
         objFSO.CreateFolder(DEFAULT_PATH)
      On Error Goto 0
   End If
   'Did we fail to create the directory?
   If Err.Number <> 0                         Or _
      Not (objFSO.FolderExists(DEFAULT_PATH))    Then
      MsgBox "Could not create the folder: '" & DEFAULT_PATH & "'"
   Else
      On Error Resume Next
         Err.Clear
         If (GlAskForInput) Then
            strFullName = GobjExcel.GetSaveAsFilename(strFullName, strFilter, 1, strTitle)  'Get the filename from user
         End If
         If Err.Number <> 0 Then
            MsgBox "Could not save Excel File: '" & strFullName & "'"
         Else
            Err.Clear
            GobjExcel.ActiveWorkbook.SaveAs strFullName
            If Err.Number = 0 Then funSaveFiles = True
         End If
      On Error Goto 0
   End If
End Function
'********************************************************
'** funUpdateWindow - Update the message status window **
'********************************************************
Function funUpdateWindow(intPercent, strResult, strReset)
   '*******************
   '** Start of code **
   '*******************
   'Write the error message to the Message Window
   On Error Resume Next
      Err.Clear      
      GobjIE.Document.ParentWindow.Document.Script.BarOP(intPercent) 'Progress bar
      If (GlAskForInput) Then
         GobjIE.Document.ParentWindow.Document.Script.ListOP(strResult) 'Message box
      End If
   On Error Goto 0
   'Problem with the window? Rebuild it
   If Err.Number <> 0 Then
      Err.Clear
      Call subShowBar() 'Rebuild window
      On Error Resume Next
         Err.Clear
         GobjIE.Document.ParentWindow.Document.Script.BarOP(intPercent) 'Progress bar
         If (GlAskForInput) Then
            GobjIE.Document.ParentWindow.Document.Script.ListOP(strReset) 'Message box
         End If
      On Error Goto 0
   End If
   Do While (GobjIE.Busy)
      Sleep 250
   Loop
   funUpdateWindow = strReset & strResult 'Redisplay all the saved lines
End Function
'************************************************************
'** funSizeFormat - Reduce number into lowest metric value **
'************************************************************
Function funSizeFormat(intBaseNumber, strCurrentSize, strReturnSize)
   '*********************
   '** Local Variables **
   '*********************
   'Constants
   Const BYTE_VALUE = 0 : Const KB_VALUE = 1 : Const MB_VALUE = 2 : Const GB_VALUE = 3 : Const TB_VALUE = 4
   Const KB = 1024 : Const MB = 1048576 : Const GB = 1073741824 'Metric values
   'Integers
   Dim intCurrentOffset, intReturnOffset, intTestValue
   'Strings
   Dim strSize
   '*******************
   '** Start of code **
   '*******************
   intCurrentOffset = -1
   intReturnOffset  = -1
   'Only process numbers
   If IsNumeric(intBaseNumber) Then
      'Is it greater then 0?
      If intBaseNumber > 0 Then
         intTestValue = intBaseNumber
         strSize      = strCurrentSize
         'What is the size of the value sent
         Select Case strCurrentSize
            Case "KB"
               intCurrentOffset = KB_VALUE
            Case "MB"
               intCurrentOffset = MB_VALUE
            Case "GB"
               intCurrentOffset = GB_VALUE
            Case "TB"
               intCurrentOffset = TB_VALUE
            Case Else
               intCurrentOffset = BYTE_VALUE
               strSize          = "Bytes"
         End Select
         'What is the size of the value to return
         Select Case strReturnSize
            Case "BY"
               intReturnOffset = BYTE_VALUE
            Case "KB"
               intReturnOffset = KB_VALUE
            Case "MB"
               intReturnOffset = MB_VALUE
            Case "GB"
               intReturnOffset = GB_VALUE
            Case Else
               intReturnOffset = TB_VALUE
         End Select
         'OK, lets make it a the right size
         Do While strReturnSize <> strSize               
            'Return value will be between 0.500 and 512.000
            If intTestValue < (KB / 2) Then
               Exit Do
            End If
            'Do we decress or incress the value?
            If intCurrentOffset < intReturnOffset Then
               intCurrentOffset = intCurrentOffset + 1 'Current is a larger base
               intTestValue     = intTestValue / KB
               Select Case strSize 
                  Case "Bytes"
                     strSize = "KB"
                  Case "KB"
                     strSize = "MB"
                  Case "MB"
                     strSize = "GB"
                  Case "GB"
                     strSize = "TB"
                  Case "TB"
                     strSize = "Error!"
               End Select
            Else
               intCurrentOffset = intCurrentOffset - 1 'Current is a smaller base
               intTestValue     = intTestValue * KB
               Select Case strSize 
                  Case "Bytes"
                     strSize = "Error!"
                  Case "KB"
                     strSize = "Bytes"
                  Case "MB"
                     strSize = "KB"
                  Case "GB"
                     strSize = "MB"
                  Case "TB"
                     strSize = "GB"
               End Select
            End If
         Loop
         funSizeFormat = FormatNumber(Round(intTestValue, 2), 2) & " " & strSize
      Else
         funSizeFormat = "" 'Less then 0? Return and empty string
      End If
   Else
      funSizeFormat = intBaseNumber 'Not a number? Return it unchanged
   End If
End Function
'***[ End of FUNCTIONS ]********************************************************************************************************
'***[ End of SCRIPT ]***********************************************************************************************************