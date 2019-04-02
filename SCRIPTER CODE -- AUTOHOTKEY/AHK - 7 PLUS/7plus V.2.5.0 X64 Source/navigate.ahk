;Maybe cleanup later for release as library?

Shell_GoBack(hWnd=0)
{
	if(hWnd||(hWnd:=WinExist("ahk_class CabinetWClass"))||(hWnd:=WinExist("ahk_class ExploreWClass")))
	{
		for Item in ComObjCreate("Shell.Application").Windows
		{
			if (Item.hwnd = hWnd)
			{
				Item.GoBack()
				return
			}
		}		
	}
}
Shell_GoForward(hWnd=0)
{
	if(hWnd||(hWnd:=WinExist("ahk_class CabinetWClass"))||(hWnd:=WinExist("ahk_class ExploreWClass")))
	{
		for Item in ComObjCreate("Shell.Application").Windows
		{
			if (Item.hwnd = hWnd)
			{
				Item.GoForward()
				return
			}
		}		
	}
}

Shell_GoUpward()
{
	path := GetCurrentFolder()
	if (Vista7 && !strEndsWith(path,".search-ms"))
		Send !{Up}
	else
		Send {Backspace}
}
ShellNavigate(Path, hWnd=0)
{
	if(hWnd||(hWnd:=WinExist("ahk_class CabinetWClass"))||(hWnd:=WinExist("ahk_class ExploreWClass")))
		DllCall(A_ScriptDir "\Explorer.dll\SetPath", "Ptr", hwnd, "Str", Path, "Cdecl")
}

RefreshExplorer() 
{ 
	hwnd:=WinExist("A")
	If (WinActive("ahk_group ExplorerGroup"))
	{
		for Item in ComObjCreate("Shell.Application").Windows
		{
			if (Item.hwnd = hWnd)
			{
				Item.Refresh()
				return
			}
		}
	}
	else if(IsDialog())
		Send {F5}
}
	
/*
fileO:
FO_MOVE   := 0x1 
FO_COPY   := 0x2 
FO_DELETE := 0x3 
FO_RENAME := 0x4

flags:
Const FOF_SILENT = 4
Const FOF_RENAMEONCOLLISION = 8
Const FOF_NOCONFIRMATION = 16
Const FOF_NOERRORUI = 1024
http://msdn.microsoft.com/en-us/library/bb759795(VS.85).aspx for more
*/
ShellFileOperation( fileO=0x0, fSource="", fTarget="", flags=0x0, ghwnd=0x0 )     
{
   ;dout_f(A_ThisFunc)
   FO_MOVE   := 0x1
   FO_COPY   := 0x2
   FO_DELETE := 0x3
   FO_RENAME := 0x4
   
   FOF_MULTIDESTFILES :=           0x1            ; Indicates that the to member specifies multiple destination files (one for each source file) rather than one directory where all source files are to be deposited.
   FOF_SILENT :=                0x4            ; Does not display a progress dialog box.
   FOF_RENAMEONCOLLISION :=       0x8            ; Gives the file being operated on a new name (such as "Copy #1 of...") in a move, copy, or rename operation if a file of the target name already exists.
   FOF_NOCONFIRMATION :=          0x10         ; Responds with "yes to all" for any dialog box that is displayed.
   FOF_ALLOWUNDO :=             0x40         ; Preserves undo information, if possible. With del, uses recycle bin.
   FOF_FILESONLY :=             0x80         ; Performs the operation only on files if a wildcard filename (*.*) is specified.
   FOF_SIMPLEPROGRESS :=          0x100         ; Displays a progress dialog box, but does not show the filenames.
   FOF_NOCONFIRMMKDIR :=          0x200         ; Does not confirm the creation of a new directory if the operation requires one to be created.
   FOF_NOERRORUI :=             0x400         ; don't put up error UI
   FOF_NOCOPYSECURITYATTRIBS :=    0x800         ; dont copy file security attributes
   FOF_NORECURSION :=             0x1000         ; Only operate in the specified directory. Don't operate recursively into subdirectories.
   FOF_NO_CONNECTED_ELEMENTS :=    0x2000         ; Do not move connected files as a group (e.g. html file together with images). Only move the specified files.
   FOF_WANTNUKEWARNING :=          0x4000         ; Send a warning if a file is being destroyed during a delete operation rather than recycled. This flag partially overrides FOF_NOCONFIRMATION.

   
   ; no more annoying numbers to deal with (but they should still work, if you really want them to)
   fileO := %fileO% ? %fileO% : fileO
   
   ; the double ternary was too fun to pass up
   _flags := 0
   Loop Parse, flags, |
      _flags |= %A_LoopField%   
   flags := _flags ? _flags : (%flags% ? %flags% : flags)
   
   If ( SubStr(fSource,0) != "|" )
      fSource := fSource . "|"

   If ( SubStr(fTarget,0) != "|" )
      fTarget := fTarget . "|"
   
   char_size := A_IsUnicode ? 2 : 1
   char_type := A_IsUnicode ? "UShort" : "Char"
   
   fsPtr := &fSource
   Loop % StrLen(fSource)
      if NumGet(fSource, (A_Index-1)*char_size, char_type) = 124
         NumPut(0, fSource, (A_Index-1)*char_size, char_type)

   ftPtr := &fTarget
   Loop % StrLen(fTarget)
      if NumGet(fTarget, (A_Index-1)*char_size, char_type) = 124
         NumPut(0, fTarget, (A_Index-1)*char_size, char_type)
   
   VarSetCapacity( SHFILEOPSTRUCT, 60, 0 )                 ; Encoding SHFILEOPSTRUCT
   NextOffset := NumPut( ghwnd, &SHFILEOPSTRUCT )          ; hWnd of calling GUI
   NextOffset := NumPut( fileO, NextOffset+0    )          ; File operation
   NextOffset := NumPut( fsPtr, NextOffset+0    )          ; Source file / pattern
   NextOffset := NumPut( ftPtr, NextOffset+0    )          ; Target file / folder
   NextOffset := NumPut( flags, NextOffset+0, 0, "Short" ) ; options

   code := DllCall( "Shell32\SHFileOperation" . (A_IsUnicode ? "W" : "A"), UInt,&SHFILEOPSTRUCT )
   ErrorLevel := ShellFileOperation_InterpretReturn(code)

   Return NumGet( NextOffset+0 )
}

ShellFileOperation_InterpretReturn(c)
{
   static dict
   if !dict
   {
      dict := Object()
      dict[0x0]      :=    ""
      dict[0x71]      :=   "DE_SAMEFILE - The source and destination files are the same file."
      dict[0x72]      :=   "DE_MANYSRC1DEST - Multiple file paths were specified in the source buffer, but only one destination file path."
      dict[0x73]      :=   "DE_DIFFDIR - Rename operation was specified but the destination path is a different directory. Use the move operation instead."
      dict[0x74]      :=   "DE_ROOTDIR - The source is a root directory, which cannot be moved or renamed."
      dict[0x75]      :=   "DE_OPCANCELLED - The operation was cancelled by the user, or silently cancelled if the appropriate flags were supplied to SHFileOperation."
      dict[0x76]      :=   "DE_DESTSUBTREE - The destination is a subtree of the source."
      dict[0x78]      :=   "DE_ACCESSDENIEDSRC - Security settings denied access to the source."
      dict[0x79]      :=   "DE_PATHTOODEEP - The source or destination path exceeded or would exceed MAX_PATH."
      dict[0x7A]      :=   "DE_MANYDEST - The operation involved multiple destination paths, which can fail in the case of a move operation."
      dict[0x7C]      :=   "DE_INVALIDFILES   - The path in the source or destination or both was invalid."
      dict[0x7D]      :=   "DE_DESTSAMETREE   - The source and destination have the same parent folder."
      dict[0x7E]      :=   "DE_FLDDESTISFILE - The destination path is an existing file."
      dict[0x80]      :=   "DE_FILEDESTISFLD - The destination path is an existing folder."
      dict[0x81]      :=   "DE_FILENAMETOOLONG - The name of the file exceeds MAX_PATH."
      dict[0x82]      :=   "DE_DEST_IS_CDROM - The destination is a read-only CD-ROM, possibly unformatted."
      dict[0x83]      :=   "DE_DEST_IS_DVD - The destination is a read-only DVD, possibly unformatted."
      dict[0x84]      :=   "DE_DEST_IS_CDRECORD - The destination is a writable CD-ROM, possibly unformatted."
      dict[0x85]      :=   "DE_FILE_TOO_LARGE - The file involved in the operation is too large for the destination media or file system."
      dict[0x86]      :=   "DE_SRC_IS_CDROM - The source is a read-only CD-ROM, possibly unformatted."
      dict[0x87]      :=   "DE_SRC_IS_DVD - The source is a read-only DVD, possibly unformatted."
      dict[0x88]      :=   "DE_SRC_IS_CDRECORD - The source is a writable CD-ROM, possibly unformatted."
      dict[0xB7]      :=   "DE_ERROR_MAX - MAX_PATH was exceeded during the operation."
      dict[0x402]      :=    "An unknown error occurred. This is typically due to an invalid path in the source or destination. This error does not occur on Windows Vista and later."
      dict[0x10000]   :=   "RRORONDEST   - An unspecified error occurred on the destination."
      dict[0x10074]   :=   "E_ROOTDIR | ERRORONDEST   - Destination is a root directory and cannot be renamed."
   }
   
   return dict[c] ? dict[c] : "Error code not recognized"
}
SetDirectory(sPath)
{
	if(!sPath)
		return
	sPath:=ExpandEnvVars(sPath)
	if(strEndsWith(sPath,":"))
		sPath .="\"
	If (WinActive("ahk_group ExplorerGroup"))
	{
		;Validity checking is turned off, it is probably better to let explorer handle this
		;Folders || Namespace || FTP || Saved search || network computers
		; if (InStr(FileExist(sPath), "D") || SubStr(sPath,1,3)="::{" || SubStr(sPath,1,6)="ftp://" || strEndsWith(sPath,".search-ms") || (strStartsWith(sPath,"\\") && !InStr(sPath,"\",false,3))) 
		; {
			hWnd:=WinExist("A")
			ShellNavigate(sPath,hwnd)
		; }
		; else
			; MsgBox The path %sPath% cannot be opened!
	}
	else if(WinActive("ahk_group DesktopGroup"))
		Run(A_WinDir "\explorer.exe /n,/e," sPath)
	else if (IsWinRarExtractionDialog())
		SetWinRarDirectory(sPath)
	else if (IsDialog())
		SetDialogDirectory(sPath)
	else
	{
		class:=WinGetClass("A")
		MsgBox Can't navigate: Wrong window %class%
	}
}

SetWinRarDirectory(Path)
{
	ControlSetText , Edit1, %Path%, A 
	ControlClick, Button1, A
}

SetDialogDirectory(Path)
{
	ControlGetFocus, focussed, A
	ControlGetText, w_Edit1Text, Edit1, A
	ControlClick, Edit1, A
	ControlSetText, Edit1, %Path%, A
	hwnd:=WinExist("A")
	ControlSend, Edit1, {Enter}, A
	Sleep, 100	; It needs extra time on some dialogs or in some cases.
	while(hwnd!=WinExist("A")) ;If there is an error dialog, wait until user closes it before continueing
		Sleep, 100
	ControlSetText, Edit1, %w_Edit1Text%, A
	ControlFocus %focussed%,A
}

IsDialog(window=0,ListViewSelected = False)
{
	result:=0
	if(window)
		window:="ahk_id " window
	else
		window:="A"
	if(WinGetClass(window)="#32770")
	{
		;Check for new FileOpen dialog
		ControlGet, hwnd, Hwnd , , DirectUIHWND3, %window%
		if(hwnd)
		{
			ControlGet, hwnd, Hwnd , , SysTreeView321, %window%
			if(hwnd)
			{
				ControlGet, hwnd, Hwnd , , Edit1, %window%
				if(hwnd)
				{
					ControlGet, hwnd, Hwnd , , Button2, %window%
					if(hwnd)
					{
						ControlGet, hwnd, Hwnd , , ComboBox2, %window%
						if(hwnd)
						{
						ControlGet, hwnd, Hwnd , , ToolBarWindow323, %window%
						if(hwnd)
							result:=(!ListViewSelected||IsControlActive("DirectUIHWND2")||IsControlActive("SysTreeView321"))
						}
					}
				}
			}
		}
		;Check for old FileOpen dialog
		if(!result)
		{ 
			ControlGet, hwnd, Hwnd , , ToolbarWindow321, %window%
			if(hwnd)
			{
				ControlGet, hwnd, Hwnd , , SysListView321, %window%
				if(hwnd)
				{
					ControlGet, hwnd, Hwnd , , ComboBox3, %window%
					if(hwnd)
					{
						ControlGet, hwnd, Hwnd , , Button3, %window%
						if(hwnd)
						{
							ControlGet, hwnd, Hwnd , , SysHeader321 , %window%
							if(hwnd)
								result:=(!ListViewSelected||IsControlActive("DirectUIHWND2")||IsControlActive("SysTreeView321")) ? 2 : 0
						}
					}
				}
			}
		}
	}
	return result
}

;Returns selected files separated by `n
GetSelectedFiles(FullName=1, hwnd=0)
{
	global MuteClipboardList
	If (WinActive("ahk_group ExplorerGroup") || (WinExist("ahk_id " hwnd) && InStr("CabinetWClass,ExploreWClass", WinGetClass("ahk_id " hwnd))))
	{
		if(!hwnd)
			hWnd:=WinExist("A")
		if FullName
			return ShellFolder(hwnd,3)
		else
			return ShellFolder(hwnd,4)
	}
	else if(Vista7 && x:=IsDialog())
	{		
		if(x=1)
		{
			ControlGetFocus, focussed ,A
			ControlFocus DirectUIHWND2, A
		}
		MuteClipboardList := true
		clipboardbackup := clipboardall
		outputdebug clearing clipboard
		clipboard := ""
		ClipWait, 0.05, 1
		outputdebug copying files to clipboard
		Send ^c
		ClipWait, 0.05, 1
		result := clipboard
		clipboard := clipboardbackup
		if(x=1)
			ControlFocus %focussed%, A
		OutputDebug, Selected Files: %result%
		MuteClipboardList:=false
		return result
	}
	else if(WinActive("ahk_group DesktopGroup"))
	{	
		; if(A_PtrSize = 8) ;64bit doesn't support listview method below yet
		; {
			; MuteClipboardList := true
			; clipboardbackup := clipboardall
			; outputdebug clearing clipboard
			; clipboard := ""
			; ClipWait, 0.15, 1
			; outputdebug copying files to clipboard
			; Send ^c
			; ClipWait, 0.15, 1
			; result := clipboard
			; clipboard := clipboardbackup
			; OutputDebug, Selected Files: %result%
			; MuteClipboardList:=false
			; return result
		; }
		; else
		; {
			ControlGet, result, List, Selected Col1, SysListView321, A ;This line causes explorer to crash on 64 bit systems when used in a 32 bit AHK build
			if(result)
			{
				Loop, Parse, result, `n  ; Rows are delimited by linefeeds (`n).
					result2 .= "`n" A_Desktop "\" A_LoopField
				return SubStr(result2,2)
			}
			else
				return ""
			}
	; }
}

GetFocussedFile()
{
	If (WinActive("ahk_group ExplorerGroup"))
	{	
		return ShellFolder(WinExist("A"),2)
	}
	else if(IsDialog()=2) ;only old Dialogs supported
	{
		ControlGet, focussed, list,focus, SysListView321, A
		return focussed
	}
}
GetCurrentFolder(hwnd=0, DisplayName=0)
{
	If (WinActive("ahk_group ExplorerGroup")||hwnd)
	{	
		if(!hwnd)
			hWnd:=WinExist("A")
		if(DisplayName)
			return ShellFolder(hwnd,5)
		Else
			return ShellFolder(hwnd,1)
	}
	If (WinActive("ahk_group DesktopGroup"))
		return %A_Desktop%
	else if((IsDialog())=1) ;No Support for old dialogs for now
	{
		ControlGetText, text , ToolBarWindow322, A
		return strTrim(SubStr(text,InStr(text," ")), " ")
	}
	return ""
}
SearchExplorer(Filter,hWnd)
{
	If (hWnd||(hWnd:=WinActive("ahk_class CabinetWClass"))||(hWnd:=WinActive("ahk_class ExploreWClass")))
	{
		for Item in ComObjCreate("Shell.Application").Windows
		{
			if (Item.hwnd = hWnd)
			{
				doc:=Item.Document
				doc.FilterView(Filter)
			}
		}
	}
}
InvertSelection(hWnd)
{
	If (hWnd||(hWnd:=WinActive("ahk_class CabinetWClass"))||(hWnd:=WinActive("ahk_class ExploreWClass")))
	{
		;Calling the menu item is more reliable than doing it through COM because right now only real files are supported
		PostMessage,0x112,0xf100,0,,A
		SendInput {Right}{Down}{Up}{Enter}
		return
		selected:="`n" GetSelectedFiles(0) "`n"
		path:=GetCurrentFolder()
		Loop, %path%\* , 1
			if(!InStr(selected,"`n" A_LoopFileName "`n"))
				NewSelection.= "`n" A_LoopFileName
		NewSelection:=strTrimLeft(NewSelection,"`n")
		SelectFiles(NewSelection,0,0,0,0)
		SelectFiles(selected,0,1,0,0)
		; for Item in ComObjCreate("Shell.Application").Windows
		; {
			; if (Item.hwnd = hWnd)
			; {
				; doc:=Item.Document
				; doc.Select(0x2)
			; }
		; }
	}
}
SelectFiles(Select,Clear=1,Deselect=0,MakeVisible=1,focus=1, hWnd=0)
{
	If (hWnd||(hWnd:=WinActive("ahk_class CabinetWClass"))||(hWnd:=WinActive("ahk_class ExploreWClass")))
	{
		for Item in ComObjCreate("Shell.Application").Windows
		{
			if (Item.hwnd = hWnd)
			{
				doc:=Item.Document
				value:=!(Deselect = 1)
				value1:=!(Deselect = 1)+(focus = 1)*16+(MakeVisible = 1)*8
				count := doc.Folder.Items.Count
				if(Clear = 1)
				{
					if(count > 0)
					{
						item := doc.Folder.Items.Item(0)
						doc.SelectItem(item,4)
						doc.SelectItem(item,0)
						; COM_Invoke(doc,"SelectItem",item,4)
						; COM_Invoke(doc,"SelectItem",item,0)
					}
				}
				if(!IsObject(Select)) ;Convert to array if passed as `n separated list
					Select := ToArray(Select)
				items := Array()
				itemnames := Array()
				Loop % count
				{
					index := A_Index
					while(true)
					{
						item := doc.Folder.Items.Item(index - 1)
						itemname := item.Name
						if(itemname != "")
						{
							; outputdebug itemname %itemname%
							break
						}
						Sleep 10
					}
					items.Insert(item)	
					itemnames.Insert(itemname)
				}
				Loop % Select.MaxIndex()
				{
					index := A_Index
					filter := Select[A_Index]
					If(filter)
					{
						SplitPath, filter, filter ;Make sure only names are used
						If(InStr(filter, "*"))
						{
							filter := "\Q" StringReplace(filter, "*", "\E.*\Q", 1) "\E"
							filter := strTrim(filter,"\Q\E")
							Loop % items.MaxIndex()
							{
								if(RegexMatch(itemnames[A_Index],"i)" filter))
								{
									doc.SelectItem(items[A_Index], index=1 ? value1 : value)
									; COM_Invoke(doc,"SelectItem",items[A_Index],index=1 ? value1 : value) ;http://msdn.microsoft.com/en-us/library/bb774047(VS.85).aspx					
									index++
								}
							}
						}
						else
						{
							Loop % items.MaxIndex()
							{
								if(itemnames[A_Index]=filter)
								{
									doc.SelectItem(items[A_Index], index=1 ? value1 : value)
									index++
									break
								}
							}
							; doc.SelectItem(doc.Folder.ParseName(COMObjParameter(0x0008,filter)),(A_Index=1 ? value1 : value))
							; COM_Invoke(doc,"SelectItem",doc.Folder.ParseName(filter),(A_Index=1 ? value1 : value)) ;http://msdn.microsoft.com/en-us/library/bb774047(VS.85).aspx					
						}
					}
				}
				return
			}
		}
	}
	else if(hWnd:=WinActive("ahk_group DesktopGroup") && A_PtrSize = 4)
	{
		SendInput %Select%
	}
}
/*
InvertSelection()
{
	;Calling the menu item is more reliable than doing it through COM because right now only real files are supported
	PostMessage,0x112,0xf100,0,,A
	SendInput {Right}{Down}{Up}{Enter}
	return
}
*/

ShellFolder(hWnd=0,returntype=0) 
{ 
	if(hWnd||(hWnd:=WinActive("ahk_class CabinetWClass"))||(hWnd:=WinActive("ahk_class ExploreWClass")))
	{
		;Find hwnd window
		for Item in ComObjCreate("Shell.Application").Windows
		{
			if (Item.hwnd = hWnd)
			{
				doc:=Item.Document
				if(!doc)
					return
				sFolder   := doc.Folder.Self.path
				if(!sFolder)
					return
				sDisplay := doc.Folder.Self.name
				if(!sDisplay)
					return
				;Don't get focussed item and selected files unless requested, because it will cause a COM error when called during/shortly after explorer path change sometimes
				if (returntype=2)
				{
					sFocus :=doc.FocusedItem.Path
					SplitPath, sFocus , sFocus
				}
				if(returntype=3 || returntype=4)
				{
					count := doc.SelectedItems.Count
					pos := 1
					while(pos <= count)
					{
						path :=doc.selectedItems.item(pos-1).Path ;= (returntype=3 ? sFolder "\" COM_Invoke(doc.SelectedItems, "Item", A_Index-1).Name "`n" : COM_Invoke(doc.SelectedItems, "Item", A_Index-1).Name "`n")
						; if(path != "") ;This check can block the SelectionChanged event because it fails sometimes here
						; {
							if(returntype=4)
								SplitPath, path , path
							sSelect := sSelect path "`n"
							pos++
						; }
					}
					StringReplace, sSelect, sSelect, \\ , \, 1 
					;The lines below fix network paths
					StringReplace, sSelect, sSelect, `n\, `n\\, 1
					if(InStr(sSelect, "\") = 1)
						sSelect := "\" sSelect
					;Remove last `n
					StringTrimRight, sSelect, sSelect, 1
				}
				if (returntype=1)
					Return   sFolder
				else if (returntype=2)
					Return   sFocus
				else if (returntype=3)
					Return   sSelect
				else if (returntype=4)
					Return 	 sSelect
				else if (returntype=5)
					Return sDisplay
			}
		}
	}
}

IsWinrarExtractionDialog()
{
	static WinRarTitle
	If (WinActive("ahk_class #32770"))
	{		
		if(WinRarTitle="")
		{
			RegRead, path, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe ,Path
			if(path)
			{
				Loop, read, %path%\winrar.lng
				{
					if(strStartsWith(A_LoopReadLine,"8f827d31"))
					{
						WinRarTitle:=strStrip(A_LoopReadLine,"""")
						break
					}
				}
				if(!WinRarTitle)
					WinRarTitle:="WinRar not found"
			}
		}		
		WinGetTitle, wintitle,A
		if(WinRarTitle=wintitle)
			return true
	}
	return false
}
