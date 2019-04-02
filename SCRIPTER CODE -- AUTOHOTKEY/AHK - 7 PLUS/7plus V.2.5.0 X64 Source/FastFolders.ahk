#include *i %A_ScriptDir%\Navigate.ahk
#include *i %A_ScriptDir%\MiscFunctions.ahk
ClearStoredFolder(Slot)
{
	global FastFolders
	Slot+=1
	FastFolders[Slot].Path := ""
	FastFolders[Slot].Title := ""
	if (Settings.Explorer.FastFolders.ShowInFolderBand)
	{
		RemoveAllExplorerButtons("IsFastFolderButton")
		loop 10
		{
			pos:=A_Index-1
			if(FastFolders[A_Index].Path)
				AddButton("",FastFolders[A_Index].Path,,pos ":" FastFolders[A_Index].Title)
		}
	}
}

UpdateStoredFolder(Slot, Folder="")
{
	global FastFolders
	Slot+=1
	if(Folder)
		FastFolders[Slot].Path := Folder
	else
		FastFolders[Slot].Path:=GetCurrentFolder()
	title:=FastFolders[Slot].Path	
	if(InStr(title,"::") = 1 && WinActive("ahk_group ExplorerGroup"))
		WinGetTitle,title,A
	
	SplitPath, title , split
	FastFolders[Slot].Title := split
	if(!FastFolders[Slot].Title)
		FastFolders[Slot].Title:=title
	RefreshFastFolders()	
}

RefreshFastFolders()
{
	if(Settings.Explorer.FastFolders.ShowInFolderBand)
		RemoveAllExplorerButtons("IsFastFolderButton")
	AddAllButtons(Settings.Explorer.FastFolders.ShowInFolderBand, Settings.Explorer.FastFolders.ShowInPlacesBar)
}

AddAllButtons(FolderBand,PlacesBar)
{
	global FastFolders
	loop 10
	{
		pos:=A_Index-1
		if(FastFolders[A_Index].Path)
		{				
			if (FolderBand)		
				AddButton("",FastFolders[A_Index].Path,,pos ":" FastFolders[A_Index].Title)
			if(pos<=4 && PlacesBar)	;Also update placesbar
			{
				value:=FastFolders[A_Index].Path
				RegWrite, REG_SZ,HKCU,Software\Microsoft\Windows\CurrentVersion\Policies\comdlg32\Placesbar, Place%pos%,%value%
			}				
		}
	}
}

;Callback function for determining if a specific registry key was created by 7plus
IsFastFolderButton(Command,Title,Tooltip)
{
	x:=substr(Title,1,1)
	if(IsNumeric(x)&&substr(Title,2,1)=":")
		return true
	return false
}

FastFolderMenu()
{
	global FastFolders
	Menu, FastFolders, add, 1,FastFolderMenuHandler1
	Menu, FastFolders, DeleteAll
	if ((IsWindowUnderCursor("ExploreWClass")||IsWindowUnderCursor("CabinetWClass")||IsWindowUnderCursor("WorkerW")||IsWindowUnderCursor("Progman")) && !IsRenaming())
	{
		win:=WinExist("A")
		y:=GetSelectedFiles()
		loop 10
		{
			i:=A_INDEX-1
			if(FastFolders[A_Index].Path)
			{
				x:=FastFolders[A_Index].Title
				if(x && (!InStr(x,"ftp://") = 1||!y))
				{
					x := "&" i ": " x
					Menu, FastFolders, add, %x%, FastFolderMenuHandler%i%
				}
			} 
		}
		hwnd:=WinExist("A")
		Menu, FastFolders, Show
		return true
	}	
	return false
}

FastFolderMenuHandler0:
FastFolderMenuClicked(0)
return
FastFolderMenuHandler1:
FastFolderMenuClicked(1)
return
FastFolderMenuHandler2:
FastFolderMenuClicked(2)
return
FastFolderMenuHandler3:
FastFolderMenuClicked(3)
return
FastFolderMenuHandler4:
FastFolderMenuClicked(4)
return
FastFolderMenuHandler5:
FastFolderMenuClicked(5)
return
FastFolderMenuHandler6:
FastFolderMenuClicked(6)
return
FastFolderMenuHandler7:
FastFolderMenuClicked(7)
return
FastFolderMenuHandler8:
FastFolderMenuClicked(8)
return
FastFolderMenuHandler9:
FastFolderMenuClicked(9)
return

FastFolderMenuClicked(index)
{
	global FastFolders
	y:=FastFolders[index].Path
	x:=GetSelectedFiles()
	StringReplace, x, x, `n , |, A
	if(x && (GetKeyState("CTRL") || GetKeyState("Shift")))
	{	
		if(GetKeyState("CTRL"))
			ShellFileOperation(0x2, x, y,0,hwnd)   
		else if(GetKeyState("Shift"))
			ShellFileOperation(0x1, x, y,0,hwnd)
	}
	else
	{
		Sleep 100
		SetDirectory(y)
	}
	Menu, FastFolders, DeleteAll
}

;Removes all buttons created with this script. Function can be the name of a function with these arguments: func(command,title,tooltip) and it can be used to tell the script if an entry may be deleted
RemoveAllExplorerButtons(function="")
{
	;go into view folders (clsid)
	Loop, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes, 2, 0
	{			
		regkey:=A_LoopRegName
		;go into number folder
		Loop, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected, 2, 0
		{
			numberfolder:=A_LoopRegName			
			
			;Custom skip function code
			;go into clsid folder
			Loop, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%numberfolder%, 2, 0
			{
				skip:=false
				RegRead, value, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%numberfolder%\%A_LoopRegName%, InfoTip
				RegRead, title, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%numberfolder%\%A_LoopRegName%, Title
				RegRead, cmd, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%numberfolder%\%A_LoopRegName%\shell\InvokeTask\command
				
				if(IsFunc(function))
					if(!%function%(cmd,title,value))
					{
						skip:=true
						break
					}
			}
			if(skip) 
				continue
			;Custom skip function code
			RegRead, ahk, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%numberfolder%, AHK			
			if(ahk)
				RegDelete, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%numberfolder%
		}
		Loop, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected, 2, 0
		{
			numberfolder:=A_LoopRegName
			
			
			;Custom skip function code
			;go into clsid folder
			Loop, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected\%numberfolder%, 2, 0
			{
				skip:=false
				RegRead, value, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected\%numberfolder%\%A_LoopRegName%, InfoTip
				RegRead, title, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected\%numberfolder%\%A_LoopRegName%, Title
				RegRead, cmd, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected\%numberfolder%\%A_LoopRegName%\shell\InvokeTask\command
				
				if(IsFunc(function))
					if(!%function%(cmd,title,value))
					{
						skip:=true
						break
					}
			}
			if(skip) 
				continue
			;Custom skip function code
			
			
			RegRead, ahk, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected\%numberfolder%, AHK
			if(ahk)
				RegDelete, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected\%numberfolder%
		}
	}
}
;Removes a button. Command can either be a real command (with arguments), a path or a function with three arguments (command, key, param) which identifies the proper key
RemoveButton(Command, param="")
{
	if(!IsFunc(Command) && InStr(Command,"\",0,strlen(Command)))
		StringTrimRight, Command, Command,1
	BaseKey := "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes"
	ButtonFound := false
	;go into view folders (clsid)
	Loop, HKLM, %BaseKey%, 2, 0
	{			
		clsid := A_LoopRegName
		;code below needs to be executed for two folders each, [selected item / no selected item]
		Keys := {TasksItemsSelected : "", TasksNoItemsSelected : ""}
		for Key, v in Keys
		{
			;Local variable inside this loop for telling found state of the selected/no selected folders
			maxnumber := -1
			;Loop through all buttons of this view (reg loop goes backwards apparently)
			Loop, HKLM, %BaseKey%\%clsid%\TasksItemsSelected, 2, 0
			{
				ButtonNumber := A_LoopRegName
				maxnumber := max(ButtonNumber, maxnumber)
				
				;Keys created by 7plus have an "AHK" key added to them to make sure that only Keys related to 7plus are modified
				RegRead, ahk, HKLM, %BaseKey%\%clsid%\TasksItemsSelected\%ButtonNumber%, AHK
				if(ahk)
				{
					;go into clsid folder
					Loop, HKLM, %BaseKey%\%clsid%\TasksItemsSelected\%ButtonNumber%, 2, 0
					{
						RegRead, value, HKLM, %BaseKey%\%clsid%\TasksItemsSelected\%ButtonNumber%\%A_LoopRegName%, InfoTip
						;Check if the current key is the correct one (possibly with a caller-defined function)
						if((!IsFunc(Command) && value = Command) || (IsFunc(Command) && %Command%(value, BaseKey "\" clsid "\TasksItemsSelected\" ButtonNumber "\" A_LoopRegName "\shell\InvokeTask\command", param)))
						{
							;If the key is correct, it may be deleted
							RegDelete, HKLM, %BaseKey%\%clsid%\TasksItemsSelected\%ButtonNumber%
							ButtonFound := true
							;after item has been deleted, we need to move the higher ones down by one
							if(maxnumber > ButtonNumber)
							{
								i := ButtonNumber + 1
								while i <= maxnumber
								{
									j := i - 1
									Runwait, reg copy HKLM\%BaseKey%\%clsid%\TasksItemsSelected\%i% HKLM\%BaseKey%\%clsid%\TasksItemsSelected\%j% /s /f, , Hide
									regdelete, HKLM, %BaseKey%\%clsid%\TasksItemsSelected\%i%
									i++
								}
							}
							break 2
						}
					}
				}
			}
		}
	}
	if(!ButtonFound)
		outputdebug % "Explorer button not found: " (param.Extends("CEvent") ? param.Name : Command)
	return ButtonFound
}

;Adds a button. You may specify a command (and possibly an argument) or a path, and a name which should be used.
AddButton(Command,path,Args="",Name="", Tooltip="",AddTo = "Both")
{
	outputdebug addbutton command %command% path %path% args %args% name %name%
	if(A_IsCompiled)
		ahk_path:="""" A_ScriptDir "\7plus.exe"""
	else
		ahk_path := """" A_AhkPath """ """ A_ScriptFullPath """"
	icon=`%SystemRoot`%\System32\shell32.dll`,3 ;Icon is not working, probably not supported by explorer, some ms entries have icons defined but they don't show up either
	if(Command)
	{
		if(!Name)
		{
			SplitPath, Command , Name
			if(Name="")
				Name:=Command
		}
		icon:=Command ",1"
		description:=command
		command .= " " args
	}
	
	if(path)
	{				
		;Remove trailing backslash
		if( InStr(path,"\",0,strlen(path)))
			StringTrimRight, path, path,1
		if(!name)
		{
			SplitPath, path , Name
			if(Name="")
				Name:=path
		}
		Command := ahk_path " """ path """"	
		description:=path	
	}		
	if(!command && !path && args) ;args only, use start 7plus with -id param
	{
		Command := """" (A_IsCompiled ? A_ScriptPath : A_AhkPath """ """ A_ScriptFullPath) """ -id:" args
		description := Tooltip
	}
	SomeCLSID:="{" . uuid(false) . "}"
	;go into view folders (clsid)
	Loop, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes, 2, 0
	{
		if(AddTo = "Both" || AddTo = "Selected")
		{
			;figure out first free key number
			iterations:=0
			regkey:=A_LoopRegName
			Loop, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected, 2, 0
			{
				iterations++
			}
			
			;Marker for easier recognition of ahk-added entries
			RegWrite, REG_SZ, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%iterations%, AHK, 1
			;Write reg keys
			RegWrite, REG_EXPAND_SZ, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%iterations%\%SomeCLSID%, Icon, %icon%
			RegWrite, REG_SZ, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%iterations%\%SomeCLSID%, InfoTip, %description%
			RegWrite, REG_SZ, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%iterations%\%SomeCLSID%, Title, %name%
			RegWrite, REG_SZ, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%iterations%\%SomeCLSID%\shell\InvokeTask\command, , %command%
		}
		if(AddTo = "Both" || AddTo = "NoSelected")
		{
			;Now the same for TasksNoItemsSelected
			iterations:=0
			;figure out first free key number
			Loop, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected, 2, 0
			{
				iterations++
			}
			
			;Marker for easier recognition of ahk-added entries
			RegWrite, REG_SZ, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected\%iterations%, AHK, 1
			;Write reg keys
			RegWrite, REG_EXPAND_SZ, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected\%iterations%\%SomeCLSID%, Icon, %icon%
			RegWrite, REG_SZ, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected\%iterations%\%SomeCLSID%, InfoTip, %description%
			RegWrite, REG_SZ, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected\%iterations%\%SomeCLSID%, Title, %name%
			RegWrite, REG_SZ, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksNoItemsSelected\%iterations%\%SomeCLSID%\shell\InvokeTask\command, , %command%
		}
	}
}
FindButton(function, param)
{
	if(!IsFunc(function))
		return false
	;go into view folders (clsid)
	Loop, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes, 2, 0
	{
		regkey:=A_LoopRegName
		maxnumber:=-1
		;loop through selected item number folders (loop goes backwards)
		Loop, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected, 2, 0
		{
			numberfolder:=A_LoopRegName
			RegRead, ahk, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%numberfolder%, AHK
			if(ahk)
			{
				;go into clsid folder
				Loop, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%numberfolder%, 2, 0
				{
					RegRead, value, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\%regkey%\TasksItemsSelected\%numberfolder%\%A_LoopRegName%\shell\InvokeTask\command
					if(%function%(value, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\" regkey "\TasksItemsSelected\" numberfolder "\" A_LoopRegName "\shell\InvokeTask\command", param))
						return true
				}
			}
		}
	}
	return false
}