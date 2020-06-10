# Version 0.6 of the PinTo10_v4 Powershell script.
#
# Elements taken from various open sources from the internet.  Wrapped in lots of "if" commands to enable argument handling and specific cases.
# In order work for localized versions of Windows modify the "Pin to Start" and "Unpin from Start" strings in the script below.
# This script only handles win32 applications, not modern apps (unlikely to change any time soon).
# Tested on Win 10 1709 & 1809 x64, should work an all version of windows 10 AFAIK.
#
# Version 0.1 (20.02.2019) - First iteration, ready for broad testing.  Basic error handling in-place.
# Version 0.2 (22.02.2019) - After semi-successful testing on 1709 found a bug that I knew from previous version where pinning to start wouldn't always work.
#                            at the first time of asking so introduced multiple retries which attempts 5 times to pin and looks for the "UNPIN" verb to exist.
# Version 0.3 (26.02.2019) - Fixed a couple of spelling mistakes.  Added support for Norwegian and Swedish PIN / UNPIN verbs.  Fixed searching for shortcuts
#                            in the taskbar folder.
# Version 0.4 (28.02.2019) - Fixed an issue that might cause problems with shortcuts with spaces in their names.  Added support for German and Austrian locales.
# Version 0.5 (28.02.2019) - Fixed German and Austrian locale support.  Changed Start Menu variable check to -like from -eq to allow Open Shell Start Menu pinning
#                            to work.  However, Open Shell is really just a case of copying or deleting files from %APPDATA%\OpenShell\Pinned .  If you want to 
#                            unpin from Open Shell using this script you HAVE to run the script on the file in %APPDATA%\OpenShell\Pinned in which case you may as
#                            well re-write the script to simply copy the file into place.
# Version 0.6 (09.03.2019) - Fixed an issue where pinning to the TaskBar would fail if the following key didn't exist in the user's registry hive... HKCU\Software\Classes\*
#
# Stuart Pearson


# Set your localized verb names here (not required for taskbar, use uppercase only)
$SystemLocale = (Get-WinSystemLocale).Name

$PinToStartVerb = ""
$UnPinFromStartVerb = ""

if ($SystemLocale -eq "en-GB" -Or $SystemLocale -eq "en-US"){
	$PinToStartVerb = "PIN TO START"
	$UnPinFromStartVerb = "UNPIN FROM START"
}

if ($SystemLocale -eq "nb-NO"){
	$PinToStartVerb = "FEST TIL START"
	$UnPinFromStartVerb = "LØSNE FRA START"
}

if ($SystemLocale -eq "sv-SE"){
	$PinToStartVerb = "FÄST PÅ START"
	$UnPinFromStartVerb = "TA BORT FRÅN START"
}

if ($SystemLocale -eq "de-DE" -Or $SystemLocale -eq "de-AT"){
	#$PinToStartVerb = "AN START ANHEFTEN"
	#$UnPinFromStartVerb = "VON START LÖSEN"
	$PinToStartVerb = "AN START"
	$UnPinFromStartVerb = "VON START"
}

if ($PinToStartVerb -eq ""){$PinToStartVerb = "PIN TO START"}
if ($UnPinFromStartVerb -eq ""){$UnPinFromStartVerb = "UNPIN FROM START"}

write-host `r`n

if (($args[0] -eq "/?") -Or ($args[0] -eq "-h") -Or ($args[0] -eq "--h") -Or ($args[0] -eq "-help") -Or ($args[0] -eq "--help")){
	write-host "This script needs to be run with three arguments."`r`n
	write-host "1 - Full path to the file you wish to pin (surround in quotes)."
	write-host "2 - Either PIN or UNPIN (case insensitive)."
	write-host "3 - Either START or TASKBAR (case insensitive)."`r`n
	write-host "Example:-"
	write-host 'powershell -noprofile -ExecutionPolicy Bypass -file PinTo10_v4.ps1 "C:\Windows\Notepad.exe" PIN START'`r`n
	Break
}

if ($args.count -eq 3){
	$TargetFile = $args[0]
	$PinUnpin = $args[1].ToUpper()
	$Location = $args[2].ToUpper()
	write-host "TargetFile ="$args[0]`r`n
	write-host "PinUnpin ="$args[1]`r`n
	write-host "Location ="$args[2]`r`n
} else {
	write-host "Incorrect number of arguments.  Exiting..."`r`n
	Break
}

# Check all the variables are correct before starting
if (!(Test-Path "$TargetFile")){
	write-host "File not found.  Exiting..."`r`n
	Break
} else {
	write-host "Target file found.  Continuing..."`r`n
}

if (($PinUnpin -ne "PIN") -And ($PinUnpin -ne "UNPIN")){
	write-host "Second argument not set to PIN or UNPIN.  Exiting..."`r`n
	Break
} else {
	write-host "PINUNPIN variable set to" $PinUnpin`r`n
}

if (($Location -ne "START") -And ($Location -ne "TASKBAR")){
	write-host "Third argument not set to START or TASKBAR.  Exiting..."`r`n
	Break
} else {
	write-host "LOCATION variable set to" $Location`r`n
}

# Split the target path to folder, filename and filename with no extension
$FileNameNoExt = (Get-ChildItem $TargetFile).BaseName
write-host "FullFileName =" $FileNameNoExt`r`n

$FileNameWithExt = (Get-ChildItem $TargetFile).Name
write-host "FileNameWithExt =" $FileNameWithExt`r`n

$Directory = (Get-Childitem $TargetFile).Directory
write-host "Directory =" $Directory`r`n

#####################################################################################################################################################################
### Start of PIN TO START

# Only do this if the target is a .exe
if ((Get-ChildItem $TargetFile).Extension -ne ".lnk" -And ($PinUnpin -eq "PIN") -And ($Location -eq "START")){
	if (Get-Childitem –Path "$env:APPDATA\Microsoft\Windows\Start Menu" -Include "$FileNameNoExt.lnk" -Recurse -ErrorAction SilentlyContinue){
		write-host "Shortcut found in current user's start menu, no need to copy, just pin"`r`n
	} else {
		write-host "Shortcut not found in current user's start menu.  Need to check the all users start menu now"`r`n

		if (Get-Childitem –Path "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu" -Include "$FileNameNoExt.lnk" -Recurse -ErrorAction SilentlyContinue){
			write-host "Shortcut found in all user's start menu, no need to copy, just pin"`r`n
		} else {
			write-host "Shortcut not found in either start menu location.  Need to copy shortcut into current user's start menu"`r`n
			# Start the shortcut creation...
			$WshShell = New-Object -comObject WScript.Shell
			$Shortcut = $WshShell.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\$FileNameNoExt.lnk")
			$Shortcut.TargetPath = $TargetFile
			$Shortcut.Save()
			
			if (test-path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\$FileNameNoExt.lnk"){
				write-host "Shortcut created.  I can now continue on to pin..."`r`n
				# Update the $TargetFile variable so you are actually pinning the shortcut from the start menu, not the original source file (if different)
				$TargetFile = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\$FileNameNoExt.lnk"
			} else {
				write-host "Shortcut didn't get created.  Exiting..."`r`n
				Break
			}
		}
	}
}

# Only do this if the target is a .lnk (and target is START)
if ((Get-ChildItem $TargetFile).Extension -eq ".lnk" -And ($PinUnpin -eq "PIN") -And ($Location -eq "START")){
	if (Get-Childitem -Path "$env:APPDATA\Microsoft\Windows\Start Menu" -Include "$FileNameWithExt" -Recurse -ErrorAction SilentlyContinue){
		write-host "Shortcut found in current user's start menu, no need to copy, just pin"`r`n
	} else {
		write-host "Shortcut not found in current user's start menu.  Need to check the all users start menu now"`r`n
		
		if (Get-Childitem –Path "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu" -Include "$FileNameWithExt" -Recurse -ErrorAction SilentlyContinue){
			write-host "Shortcut found in all user's start menu, no need to copy, just pin"`r`n
		} else {
			write-host "Shortcut not found in either start menu location.  Need to copy shortcut into current user's start menu"`r`n
			# Copy shortcut to current user's start menu
			Copy-Item $TargetFile -Destination "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
			
			if (test-path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\$FileNameWithExt"){
				write-host "Shortcut copied.  I can now continue on to pin..."`r`n
				# Update the $TargetFile variable so you are actually pinning the shortcut from the start menu, not the original source file (if different)
				$TargetFile = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\$FileNameWithExt"
			} else {
				write-host "Shortcut didn't get copied.  Exiting..."`r`n
				Break
			}
		}
	}
}

if (($PinUnpin -eq "PIN") -And ($Location -eq "START")){

	$shell = New-Object -ComObject "Shell.Application"

	$FileNameWithExt = (Get-ChildItem $TargetFile).Name
	write-host "FileNameWithExt = " $FileNameWithExt`r`n

	$Directory = (Get-Childitem $TargetFile).Directory
	write-host "Directory = " $Directory`r`n
	
	$folder = $shell.Namespace("$Directory\")
	$file = $folder.Parsename("$FileNameWithExt")

	#$file.Verbs() | select *

	#$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $PinToStartVerb}
	$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -like "$PinToStartVerb*"}

	if ($verb) {
		write-host "Pin to Start verb found.  Trying to pin."`r`n
		$verb.DoIt()
		
		# Start waiting / looking for the verb to change to "UNPIN".  Try 10 times (should be more than enough).
		Start-Sleep -m 500
		$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $UnPinFromStartVerb}
		if ($verb) {write-host "Verb has now changed to 'UNPIN'.  Exiting" ; Break} else {write-host "Verb hasn't changed from 'PIN'.  Trying again..."}
		
		Start-Sleep -m 500
		$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $UnPinFromStartVerb}
		if ($verb) {write-host "Verb has now changed to 'UNPIN'.  Exiting" ; Break} else {write-host "Verb hasn't changed from 'PIN'.  Trying again..."}
		
		Start-Sleep -m 500
		$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $UnPinFromStartVerb}
		if ($verb) {write-host "Verb has now changed to 'UNPIN'.  Exiting" ; Break} else {write-host "Verb hasn't changed from 'PIN'.  Trying again..."}
		
		Start-Sleep -m 500
		$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $UnPinFromStartVerb}
		if ($verb) {write-host "Verb has now changed to 'UNPIN'.  Exiting" ; Break} else {write-host "Verb hasn't changed from 'PIN'.  Trying again..."}
		
		Start-Sleep -m 500
		$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $UnPinFromStartVerb}
		if ($verb) {write-host "Verb has now changed to 'UNPIN'.  Exiting" ; Break} else {write-host "Verb hasn't changed from 'PIN'.  Trying again..."}
				
		Start-Sleep -m 500
		$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $UnPinFromStartVerb}
		if ($verb) {write-host "Verb has now changed to 'UNPIN'.  Exiting" ; Break} else {write-host "Verb hasn't changed from 'PIN'.  Trying again..."}
				
		Start-Sleep -m 500
		$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $UnPinFromStartVerb}
		if ($verb) {write-host "Verb has now changed to 'UNPIN'.  Exiting" ; Break} else {write-host "Verb hasn't changed from 'PIN'.  Trying again..."}
				
		Start-Sleep -m 500
		$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $UnPinFromStartVerb}
		if ($verb) {write-host "Verb has now changed to 'UNPIN'.  Exiting" ; Break} else {write-host "Verb hasn't changed from 'PIN'.  Trying again..."}
				
		Start-Sleep -m 500
		$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $UnPinFromStartVerb}
		if ($verb) {write-host "Verb has now changed to 'UNPIN'.  Exiting" ; Break} else {write-host "Verb hasn't changed from 'PIN'.  Trying again..."}
		
		Start-Sleep -m 500
		$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $UnPinFromStartVerb}
		if ($verb) {write-host "Verb has now changed to 'UNPIN'.  Exiting" ; Break} else {write-host "Verb still hasn't changed from 'PIN'.  Giving up..."; Break}
		
		
	}
}
### End of PIN TO START
#####################################################################################################################################################################


#####################################################################################################################################################################
### Start of UNPIN FROM START

if (($PinUnpin -eq "UNPIN") -And ($Location -eq "START")){

	$shell = New-Object -ComObject "Shell.Application"
	
	$folder = $shell.Namespace("$Directory\")
	$file = $folder.Parsename("$FileNameWithExt")

	#$file.Verbs() | select *

	#$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -eq $UnPinFromStartVerb}
	$verb = $file.Verbs() | Where-Object {($_.Name.replace('&','')).ToUpper() -like "$UnPinFromStartVerb*"}

	if ($verb) {
		write-host "Unpin from Start verb found.  Trying to unpin."`r`n
		$verb.DoIt()
		Break
	} else {
		write-host "Unpin from Start verb not found.  Exiting..."`r`n
		Break
	}

}
### End of UNPIN FROM START
#####################################################################################################################################################################


#####################################################################################################################################################################
### Start of PIN TO TASKBAR

if (($PinUnpin -eq "PIN") -And ($Location -eq "TASKBAR")){

	$ValueData = (Get-ItemProperty("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\Windows.taskbarpin")).ExplorerCommandHandler

	$classesKey = (Get-Item "HKCU:\SOFTWARE").OpenSubKey("Classes", $true)
	$starKey = $classesKey.CreateSubKey("*", $true)
	
	$classesStarKey = (Get-Item "HKCU:\SOFTWARE\Classes").OpenSubKey("*", $true)
	$shellKey = $classesStarKey.CreateSubKey("shell", $true)
	$specialKey = $shellKey.CreateSubKey("{:}", $true)
	$specialKey.SetValue("ExplorerCommandHandler", $ValueData)
	$specialKey.Close()

	$Shell = New-Object -ComObject "Shell.Application"
	$Folder = $Shell.Namespace((Get-Item $TargetFile).DirectoryName)
	$Item = $Folder.ParseName((Get-Item $TargetFile).Name)

	# Un-comment the line below to show a list of all verbs found on the file
	#$Item.Verbs() | select *
	
	# Need to search location for existing shortcut as this method can't specify pin or unpin
	# This isn't perfect as the shortcut is not always named the same as the target, it can be named after the "Description" in the "General" tab of the file properties
	
	if (!(Get-Childitem -Path "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\*" -Include "$FileNameNoExt.lnk" -ErrorAction SilentlyContinue)){
			write-host "Shortcut not found in the Taskbar shortcuts folder.  Trying to pin."
			$Item.InvokeVerb("{:}")
	} else {
		write-host "Shortcut already found in the Taskbar shortcuts folder.  Exiting..."
	}
	
	$shellKey.DeleteSubKey("{:}")
	Break
}
### End of PIN TO TASKBAR
#####################################################################################################################################################################


#####################################################################################################################################################################
### Start of UNPIN FROM TASKBAR

if (($PinUnpin -eq "UNPIN") -And ($Location -eq "TASKBAR")){

	$ValueData = (Get-ItemProperty("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\Windows.taskbarpin")).ExplorerCommandHandler

	$classesKey = (Get-Item "HKCU:\SOFTWARE").OpenSubKey("Classes", $true)
	$starKey = $classesKey.CreateSubKey("*", $true)
	
	$classesStarKey = (Get-Item "HKCU:\SOFTWARE\Classes").OpenSubKey("*", $true)
	$shellKey = $classesStarKey.CreateSubKey("shell", $true)
	$specialKey = $shellKey.CreateSubKey("{:}", $true)
	$specialKey.SetValue("ExplorerCommandHandler", $ValueData)
	$specialKey.Close()

	$Shell = New-Object -ComObject "Shell.Application"
	$Folder = $Shell.Namespace((Get-Item $TargetFile).DirectoryName)
	$Item = $Folder.ParseName((Get-Item $TargetFile).Name)

	# Un-comment the line below to show a list of all verbs found on the file
	#$Item.Verbs() | select *
	
	# Need to search location for existing shortcut as this method can't specify pin or unpin
	
	if (Get-Childitem -Path "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\*" -Include "$FileNameNoExt.lnk" -ErrorAction SilentlyContinue){
			write-host "Shortcut found in the Taskbar shortcuts folder.  Trying to unpin."
			$Item.InvokeVerb("{:}")
	} else {
		write-host "Shortcut not found in the Taskbar shortcuts folder.  Exiting..."
	}
	
	$shellKey.DeleteSubKey("{:}")
	Break
}

### End of UNPIN FROM TASKBAR
#####################################################################################################################################################################