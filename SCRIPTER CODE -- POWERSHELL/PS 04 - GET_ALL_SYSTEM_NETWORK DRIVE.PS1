# -------------------------------------------------------------------------------
# -- C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\
# -- \PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.BAT
# -- \PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1
# -------------------------------------------------------------------------------
# -- WORK TODAY
# -- OPERATION WILL SCAN ALL NETWORK SHARE PATH
# -- USER TO SUPPLY COMPUTER SYSTEM NAME TO ARRAY
# -------------------------------------------------------------------------------
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
# -------------------------------------------------------------------------------
# -- View all shares on remote machine with PowerShell
# -- https://www.itprotoday.com/powershell/view-all-shares-remote-machine-powershell
# -------------------------------------------------------------------------------
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
# -------------------------------------------------------------------------------
# -- Thu 27-Feb-2020 10:06:55
# -- Thu 27-Feb-2020 15:34:00 -- 5 HOUR 27 MINUTE
# -- AND SECOND SESSION LATER SAME SHIFT
# -- Thu 27-Feb-2020 20:18:14
# -- Thu 27-Feb-2020 22:20:00 -- 02 HOUR 02 MINUTE
# -- Thu 27-Feb-2020 22:20:00 -- 07 HOUR 29 MINUTE -- TOTAL ADDITION OF TWO TIME PART 
# -- Thu 27-Feb-2020 22:20:00 -- 12 HOUR 13 MINUTE -- TOTAL TIME WITH AR ANOTHER THING
# -------------------------------------------------------------------------------
# -- HERE IS THE LINE TO USER FOR FILELOCATOR HAS AN EQUAL
# =C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT
# -------------------------------------------------------------------------------
# -- HERE IS THE SCRIPT I USER FOR FILELOCATOR -- FOR USE SCRIPT MODE
# -- AND TO FIND ALL GOODSYNC STUFF LEFT OVER SAVE WHAT HAVE YOU
# -- WHILE IT THERE ALSO SCAN THE TEMP FOLDER -- WHICH MINE ARE NAME AFTER COMPUTER NAME SO HELPER
# C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 68-FILE LOCATOR -- SCRIPT - GOODSYNC _GSDATA_.VBS
# -- IT IS ON MY GITHUB -- VBS FOLDER
# -------------------------------------------------------------------------------

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
$path = Get-Location
$scriptName = $MyInvocation.MyCommand.Name
(Get-Item $scriptName ).Name | OUT-NULL
$COMPNAME = $env:computername
# OF OPERATION SOMETIME DEPEND I USER COMPUTER-NAME
# IT FILTER THE SYNC OPERATION
# AND OTHER OPERATION 
# I USER A CODE LABEL --_ name *#NFS_EX* _-- WITH GOODSYNC GLOBAL 
# IT OPERATE A FILTER THE WORD SO NONE SYNC OVER
# NETWORK FILE SYSTEM EXCLUDER
# -------------------------------------------------------------------
$COMPNAME = "#NFS_EX"
$SCRIPT_LOG = "$path\$scriptName"+"__"+"$COMPNAME"+"__"+".TXT"




# ---------------------------------------------------------------------
# WRITE ALL NETWORK DRIVE FINDER ALL COMPUTER SYSTEM
$SHARE01 | Out-File -FilePath "$SCRIPT_LOG"
# ---------------------------------------------------------------------

# ---------------------------------------------------------------------
# PAUSE

# EXIT
# ---------------------------------------------------------------------


# ---------------------------------------------------------------------
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

