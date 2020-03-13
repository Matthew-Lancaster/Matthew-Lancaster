# ===================================================================
# __ C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1
# __ 
# __ PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1
# 
# __ BY Matthew __ Matt.Lan@Btinternet.com
# ===================================================================
#    FIRST RUN AT HER
# -- Thu 27-Feb-2020 10:06:55
# -- Thu 27-Feb-2020 15:34:00 -- 5 HOUR 27 MINUTE
# ------------------------------------------------------------------
#    ADDITIONAL WITH ONE DAY WORKER
# -- Thu 27-Feb-2020 20:18:14
# -- Thu 27-Feb-2020 22:20:00 -- 02 HOUR 02 MINUTE
# -- Thu 27-Feb-2020 22:20:00 -- 07 HOUR 29 MINUTE -- TOTAL ADDITION OF TWO TIME PART 
# -- Thu 27-Feb-2020 22:20:00 -- 12 HOUR 13 MINUTE -- TOTAL TIME WITH AR ANOTHER THING
# ------------------------------------------------------------------
#    SECOND RUN AT HER -- TWO DAY LATER -- SEEM LIKE ONE ANOTHER ALL DAY THING
# -- Sat 29-Feb-2020 16:15:57
# -- Sat 29-Feb-2020 19:35:00 -- 2 HOUR 48 MINUTE
# ===================================================================

# ------------------------------------------------------------------
# LOCATION SOURCE ONLINE
# ---------------------------------------------------------------------
# Matthew-Lancaster/Autokey -- 94-ALL_LOWER_THAN_NORMAL_PROCCES_PRIORITY_RESTORE.ahk
# ---------------------------------------------------------------------
# https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOHOTKEY/Autokey%20--%2094-ALL_LOW_PROCCES_PRIORITY_TO_NORMAL.ahk 
# ---------------------------------------------------------------------
# HTTPS://TINYURL.COM/SRZ69YL
# ---------------------------------------------------------------------

# ---------------------------------------------------------------------
# -- C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\
# -- \PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.BAT
# -- \PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1
# ---------------------------------------------------------------------
# -- WORK TODAY
# -- OPERATION WILL SCAN ALL NETWORK SHARE PATH
# -- USER TO SUPPLY COMPUTER SYSTEM NAME TO ARRAY
# ---------------------------------------------------------------------
# -- OPERATION USE CODE TO RUN WITH FILELOCATOR
# -- LOOK DOCUMENT FOR FILELOCATOR THEY NOT ANY MORE SUPPLY 
# -- FILELOCATOR NETWORK VERSION
# -- C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe
# -- ONLY ONE APP WHEN PRO - AR 
# -- IT GOT GOOD SCANNER
# -- AND THEN -- NETWORK WE WANT
# -- HERE WILL SCRIPT PROVIDE IN TEXT FILE AND DISPLAY
# -- ALL NETWORK LOCATION 
# -- MAIN SOURCE INITIAL
# -- HERE CREDIT
# ---------------------------------------------------------------------
# -- View all shares on remote machine with PowerShell
# -- https://www.itprotoday.com/powershell/view-all-shares-remote-machine-powershell
# ---------------------------------------------------------------------
# -- AND 40 PAGE URL OPEN MINUS CLICK BACK TRAVELER -- FOR MY FIRST LEARN THIS STYLER
# -- INCLUDE 
# --  1.. ADD STRING
# --  2.. ARRAY 
# --  3.. SET A VARIABLE
# --  4.. WRITE VARIABLE TO SCREEN
# --  5.. CARRIAGE RETURN LINE FED IN A STRING
# --  *.. SEARCH AR STRING
# --  7.. WITH USER __ IF STATEMENT COMMAND CONDITION
# --  8.. CHECK A VARIABLE STRING EQUAL NOTHING
# --  9.. GET THE RUN SCRIPT PATH NAME AND ALSO REMOVE EXTENSION
# -- 10.. AUTOHOTKEYS SCRIPT MAKE THE RUN OF POWER-SHELL EASY -- FROM NOTEPAD++
# -- 11.. REVERSE THE ARRAY
# -- 12.. READ A FILE 
# -- 13.. GET $SCRIPTPATH -- PRO_SOLUTION __ CREDIT ON __ REFERENCE_CREDIT_001 __ ENDER SCRIPTER
# -- 14.. STORE ENTIRE TEXT FILE CONTENTS IN VARIABLE ___ REFERENCE_CREDIT_002
# -- 15.. IF COMPARE NEW NEATER METHOD -EQ ______________ REFERENCE_CREDIT_003
# ---------------------------------------------------------------------
# -- SESSION_001 
# ---------------------------------------------------------------------
# -- Thu 27-Feb-2020 10:06:55
# -- Thu 27-Feb-2020 15:34:00 -- 5 HOUR 27 MINUTE
# -- AND SECOND SESSION LATER SAME SHIFT
# -- Thu 27-Feb-2020 20:18:14
# -- Thu 27-Feb-2020 22:20:00 -- 02 HOUR 02 MINUTE
# -- Thu 27-Feb-2020 22:20:00 -- 07 HOUR 29 MINUTE -- TOTAL ADDITION OF TWO TIME PART 
# -- Thu 27-Feb-2020 22:20:00 -- 12 HOUR 13 MINUTE -- TOTAL TIME WITH AR ANOTHER THING
# ---------------------------------------------------------------------
# -- SESSION_002 WITH ANOTHR TWO DAY LATER
# ---------------------------------------------------------------------
# -- Sat 29-Feb-2020 16:15:57
# -- Sat 29-Feb-2020 19:04:00 -- 2 HOUR 48 MINUTE
# ---------------------------------------------------------------------
# ---------------------------------------------------------------------
# -- HERE IS THE LINE TO USER FOR FILELOCATOR HAS AN EQUAL
# =C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT
# ---------------------------------------------------------------------
# -- HERE IS THE SCRIPT I USER FOR FILELOCATOR -- FOR USE SCRIPT MODE
# -- AND TO FIND ALL GOODSYNC STUFF LEFT OVER SAVE WHAT HAVE YOU
# -- WHILE IT THERE ALSO SCAN THE TEMP FOLDER -- WHICH MINE ARE NAME AFTER COMPUTER NAME SO HELPER
# C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 68-FILE LOCATOR -- SCRIPT - GOODSYNC _GSDATA_.VBS
# -- IT IS ON MY GITHUB -- VBS FOLDER
# ---------------------------------------------------------------------
# AFTER SESSION _ 002
# ---------------------------------------------------------------------
# NOW FILELOCATOR APP USER HAS MORE GO
# ---------------------------------------------------------------------
# C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe
# ---------------------------------------------------------------------
# IT ABLE RUN SCRIPT FOR FIRST USE WITHIN THE APP
# BY FIRST FILE SEARCH IS A DUMMY ONE
# BUTT ON TESTER IT NOT ABLE TO HAVE ANY VARIABLE STORE OR COUNTER
# AND NOT ABLE USE CUSTOM VARIABLE ANY WAY TO STORE BETWEEN CALL SCRIPTOR
# AND NONE READ DEFAULT LABEL ABOUT TELL ANY APART
# SO NOT DESIRE MAY RUN TWICE QUICKER
# AND ANOTHER COMPLICATE
# IT NOT ABLE RUN SCRIPT TO PROVIDE FILE UNTIL PRODUCER OPERATION OVER
# EXAMPLE COMPLAIN FILE SCRIPT INFO IS MISSER
# AND MORE -- REQUIRE TWO QUICK RUN -- NOT TO CREATE SCRIPT FILE 
# TO GIVE IT AN UPDATER 2 PASS METHOD OPERATION
# GET INFO WRITE AND THEN USER
# ---------------------------------------------------------------------
# OR HAVE SYSTEM RUN REGULAR ONCE AN HOUR 
# TWO STOP & RUN TWICE QUICKER
# I CREATE IT NOT STORE FILE IF CONTENT SAME
# AND TRY HIDER WAY
# ---------------------------------------------------------------------
# NIGHT BEFORE I MAKE SO SCRIPT GENERATE THAT NONE 
# NETWORK PATH SHOW FOR LOCAL COMPUTER
# WHEN ALL LOCAL DRIVE PROVIDE -- ANYWAY -- IN ONE OPERATION
# SAVE DUPLICAR
# REM BACK IN -- IF OPERATION WANT AR FULL NETWORK
# ---------------------------------------------------------------------
# POWERSHELL OPERATOR A MSGBOX TO INSTRUCTION RUN ANOTHER OF FILELOCATOR 
# IF CHANGE HAPPEN
# UNTIL NEXT SPERT TIME
# Sat 29-Feb-2020 19:58:00
# ---------------------------------------------------------------------
# ---------------------------------------------------------------------
# ---------------------------------------------------------------------

$COMPUTER = "localhost"   # NOT USER IN THIS SCRIPT

$OFS = "`r`n"

$mdarray1 = @()
$mdarray1 += ,@('\\7-ASUS-GL522VW')
$mdarray1 += ,@('\\4-ASUS-GL522VW')
$mdarray1 += ,@('\\8-MSI-GP62M-7RD')
$mdarray1 += ,@('\\5-ASUS-P2520LA')
$mdarray1 += ,@('\\1-ASUS-X5DIJ')
$mdarray1 += ,@('\\2-ASUS-EEE')
$mdarray1 += ,@('\\3-LINDA-PC')

$SHARE01=""
# Write-host ($array10)

ForEach($array10 in $mdarray1)
{
	$COMPNAME="*"
	$COMPNAME+= $env:computername
	$SET_GO="1"
	IF ($array10 -like $COMPNAME) {$SET_GO="0"}
	# ---------------------------------------------------------------
	# $SET_GO="1" ---- OPTIONAL REM OUT HERE LINE AND OPERATION GIVER ALL COMPUTER
	#             ---- IDEA IS FOR FILELOCATOR APP -- NOT TO GIVE OWN COMPUTER AS AVAILABLE FROM 
	#             ---- $(AllLocalDrives)\ -- AND THEN LOCAL AND WITH AR REMOTE NETWORK
	# ---------------- USER HERE CODER -- I PUT ALL DRIVE IN THE BLUE AGAIN NOT REDDER
	# ---------------- Sat 29-Feb-2020 00:42:00
	# ---------------------------------------------------------------
	IF ($SET_GO -contains '1')
	{
		$SHARE01 += net view ($array10) /all | select -Skip 7 | ?{$_ -match 'disk*'} | %{$_ -match '^(.+?)\s+Disk*'|out-null;$matches[1]}
		
		$SHARE01 += "`r`n"
	}
}

$SHARE02=$SHARE01
$SHARE01=""
ForEach ($line in $($SHARE02 -split "`r`n"))
{
	$CORRECT="0"
	IF ($line -like 'admin*') {$CORRECT="1"}
	IF ($line -like 'print*') {$CORRECT="1"}
	$line=$line.trim()
	IF (!$line)               {$CORRECT="1"}
	
	IF ($CORRECT -contains '0') {
		$SHARE01 += $line
		$SHARE01 += "`r`n"
	}
}


# PUT COMPUTER NAME IN FRONT WHEN MATCHER
# -------------------------------------------------------------------
$SHARE02=$SHARE01
$SHARE01=""

ForEach($array10 in $mdarray1)
{
ForEach($line in $($SHARE02 -split "`r`n"))
{

	$CORRECT=$array10
	$CORRECT=$CORRECT.replace("\\","")
	$CORRECT=$CORRECT.replace("-","_")
	$CORRECT=$CORRECT.trim()
	$CORRECT+="*"
	
	IF ($line -like $CORRECT) {
		$SHARE01 += $array10
		$SHARE01 += "\"
		$SHARE01 += $line
		$SHARE01 += "`r`n"
	}
}
}

# ---------------------------------------------------------------------
# REVERSE THE ARRAY AND THEN QUICKER COMPUTER DO BEFORE
# LOW-END COMPUTER ARE SORT FORWARD MORE
# ---------------------------------------------------------------------
$mdarray2 = @()
ForEach($line in $($SHARE01 -split "`r`n"))
{
	$line=$line.trim()
	IF ($line) {
		$mdarray2 += ,@($line)
	}
}
# ---------------------------------------------------------------------
# REVERSE THE ARRAY AND THEN QUICKER COMPUTER DO BEFORE
# LOW-END COMPUTER ARE SORT FORWARD MORE
# ---------------------------------------------------------------------
# NOT BOTHER DO BY ARRANGE ARRAY AT TOP $MDARRAY1
# [Array]::Reverse($mdarray2)
# ---------------------------------------------------------------------
$SHARE01=""
ForEach($line in $mdarray2)
{
	$line=$line.trim()
	IF ($line) {
		$SHARE01 += $line
		$SHARE01 += "`r`n"
	}

}

# ---------------------------------------------------------------------
# DISPLAY ALL NETWORK DRIVE FINDER ALL COMPUTER SYSTEM
# ---------------------------------------------------------------------
Write-Host ($SHARE01)
# ---------------------------------------------------------------------

# ---------------------------------------------------------------------
# GET VARIABLE TO CONTAIN SCRIPT PATH AND NAME AND 
# STRIP EXTENSION TO MODIFY ADDITION
# ---------------------------------------------------------------------
# $path = Get-Location
# $scriptName = $MyInvocation.MyCommand.Name
# (Get-Item $scriptName ).Name | OUT-NULL
# $COMPNAME = $env:computername
# ---------------------------------------------------------------------
# OF OPERATION SOMETIME DEPEND I USER COMPUTER-NAME
# IT FILTER THE SYNC OPERATION
# AND OTHER OPERATION 
# I USER A CODE LABEL --_ name *#NFS_EX* _-- WITH GOODSYNC GLOBAL 
# IT OPERATE A FILTER THE WORD SO NONE SYNC OVER
# NETWORK FILE SYSTEM EXCLUDER
# -------------------------------------------------------------------
# $COMPNAME = "#NFS_EX"
# $SCRIPT_LOG = "$path\$scriptName"+"__"+"$COMPNAME"+"__"+".TXT"

# $SCRIPT_LOG = "$PSCommandPath\$scriptName"+"__"+"$COMPNAME"+"__"+".TXT"
# -----------------------------------------------------------------
# -- HAVE CODE HASSLE -- POWERSHELL WHEN RUN NOT REPORT LOCATION 
# -- OF SCRIPT RUN FOLDER
# -- SET FROM BAT FOR NOW
# -----------------------------------------------------------------

# -----------------------------------------------------------------
# REFERENCE_CREDIT_001 __ ENDER SCRIPTER
# -----------------------------------------------------------------
# If using ISE
if ($psISE) {
	$SCRIPTPATH = Split-Path -Parent $psISE.CurrentFile.FullPath
# If Using PowerShell 3 or greater
} elseif($PSVersionTable.PSVersion.Major -gt 3) {
	$SCRIPTPATH = $PSScriptRoot
# If using PowerShell 2 or lower
} else {
	$SCRIPTPATH = split-path -parent $MyInvocation.MyCommand.Path
}
# -----------------------------------------------------------------

# Write-Host $SCRIPTPATH
# PAUSE

# -----------------------------------------------------------------
$SCRIPTNAME = $MyInvocation.MyCommand.Name
$COMPNAME = "#NFS_EX"
$SCRIPT_LOG = "$SCRIPTPATH\$SCRIPTNAME"+"__"+"$COMPNAME"+"__"+".TXT"
# -----------------------------------------------------------------


# Write-Host $SCRIPT_LOG
# PAUSE

# WHEN READER CONTENT OPERATION GIVE NOT VBCRLF ON
# SO HARD NOT CORRECT WAY READ FILE FOR OUR LOVELY PEOPLE
# THE INTERNET IS A PAIN
# -------------------------------------------------------------------
# $APPLES = Get-Content $SCRIPT_LOG
# -------------------------------------------------------------------
# $ORANGES = $SHARE01
# $ORANGES = $ORANGES -replace "`r", ""
# $ORANGES = $ORANGES -replace "`n", ""
# -------------------------------------------------------------------
if (Test-Path $SCRIPT_LOG) {
	# FILE EXIST
	$APPLES = [IO.File]::ReadAllText($SCRIPT_LOG)
}
# -------------------------------------------------------------------

# -------------------------------------------------------------------
$ORANGES = $SHARE01
$ORANGES+="`r`n"
# "`r`n" -- ADDITION HERE MAKE SAME LENGTH -- FROM SAVE FILE VERSION
# -------------------------------------------------------------------

# -------------------------------------------------------------------
# Write-Host "01---------------------"
# Write-Host $APPLES 
# Write-Host "02---------------------"
# Write-Host $ORANGES
# -------------------------------------------------------------------
# Write-Host "03---------------------"
# Write-Host $APPLES.LENGTH
# Write-Host $ORANGES.LENGTH
# -------------------------------------------------------------------


# IF BOTH SAME NOT REQUIRE TO SAVE
# -------------------------------------------------------------------
IF ($APPLES -EQ $ORANGES) {
	EXIT
	Write-Host "YEAH03"
}
# -------------------------------------------------------------------

# ---------------------------------------------------------------------
# REMOVE BEFORE WRITE FILE BUTTY NOT REQUIRE -- OVERWRITE IS DO ANYWAY
# ---------------------------------------------------------------------
# Remove-Item "$SCRIPT_LOG"
# ---------------------------------------------------------------------
# WRITE ALL NETWORK DRIVE FINDER ALL COMPUTER SYSTEM
# ---------------------------------------------------------------------
$SHARE01 | Out-File -FilePath "$SCRIPT_LOG"
# ---------------------------------------------------------------------

# ---------------------------------------------------------------------
# PAUSE
# ---------------------------------------------------------------------

# ---------------------------------------------------------------------
EXIT
# ---------------------------------------------------------------------



# ---------------	------------------------------------------------------
# $FileNamePath
# (Get-Item $FileNamePath ).Extension
# (Get-Item $FileNamePath ).Basename
# (Get-Item $FileNamePath ).Name
# (Get-Item $FileNamePath ).DirectoryName
# (Get-Item $FileNamePath ).FullName
# ---------------------------------------------------------------------

# ---------------------------------------------------------------------
# Get-WmiObject Win32_MappedLogicalDisk -ComputerName $computer
# Get-WmiObject Win32_MappedLogicalDisk -ComputerName $computer | fl *
# Get-WmiObject Win32_MappedLogicalDisk -computerName $computer | select Name, ProviderName,FileSystem,@{Name="Total Size";Expression={$_.size / 1GB}}, @{Name="Free Space";Expression={$_.freespace / 1GB}} | ft -AutoSize
# ---------------------------------------------------------------------


# ---------------------------------------------------------------------
# windows - List mapped network drives from the command line to text file - Super User
# https://superuser.com/questions/135752/list-mapped-network-drives-from-the-command-line-to-text-file
# --------
# windows xp - How do I get list of all network drives mapped to my system? - Super User
# https://superuser.com/questions/248783/how-do-i-get-list-of-all-network-drives-mapped-to-my-system?rq=1
# --------
# CMD DIRECTORY NAME WITHOUT EXTENSION - Google Search
# https://www.google.co.uk/search?q=CMD+DIRECTORY+NAME+WITHOUT+EXTENSION
# --------
# PowerShell Get List Of Folders Shared - Super User
# https://superuser.com/questions/769679/powershell-get-list-of-folders-shared
# --------
# powershell - Get a list of computers that are mapped to a server - Super User
# https://superuser.com/questions/1203629/get-a-list-of-computers-that-are-mapped-to-a-server
# --------
# PoshTip #35 – How To List Your Windows Mapped Drives? – www.get-cmd.com
# http://get-cmd.com/?p=3720
# --------
# Get-SmbShareAccess
# https://docs.microsoft.com/en-us/powershell/module/smbshare/get-smbshareaccess?view=win10-ps
# ---------------------------------------------------------------------


# ---------------------------------------------------------------------
# REFERENCE_CREDIT_001 __ ENDER SCRIPTER
# ---------------------------------------------------------------------
# What's the best way to determine the location of the current PowerShell script? - Stack Overflow
# https://stackoverflow.com/questions/5466329/whats-the-best-way-to-determine-the-location-of-the-current-powershell-script/58768926#58768926 
# ---------------------------------------------------------------------

# ---------------------------------------------------------------------
# REFERENCE_CREDIT_002 __ ENDER SCRIPTER
# ---------------------------------------------------------------------
# PowerShell: Store Entire Text File Contents in Variable - Stack Overflow
# https://stackoverflow.com/questions/7976646/powershell-store-entire-text-file-contents-in-variable/7976784#7976784 
# ---------------------------------------------------------------------

# ---------------------------------------------------------------------
# REFERENCE_CREDIT_003 __ ENDER SCRIPTER
# ---------------------------------------------------------------------
# How to compare the contents of two string objects in PowerShell - Stack Overflow
# https://stackoverflow.com/questions/18772063/how-to-compare-the-contents-of-two-string-objects-in-powershell/18773397#18773397 
# ---------------------------------------------------------------------



