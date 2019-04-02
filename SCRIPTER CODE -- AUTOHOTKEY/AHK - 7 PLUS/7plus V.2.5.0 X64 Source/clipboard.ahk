;Called when clipboard changes, used for "Paste text/image as file" functionality and for clipboard manager
;To use the clipboard without triggering these features, set MuteClipboardList:=true before writing to clipboard
OnClipboardChange:
if(!ApplicationState.ClipboardListenerRegistered)
	OnClipboardChange()
return

OnClipboardChange()
{
	global MuteClipboardList, ClipboardList
	if(MuteClipboardList)
	{
		FileAppend, %A_Now%: Clipboard changed to %Clipboard% but it's muted`n, %A_Temp%\7plus\Log.log
		return
	}
	if(WinActive("ahk_group ExplorerGroup") || WinActive("ahk_group DesktopGroup")|| IsDialog())
		CreateFileFromClipboard()
	text:=ReadClipboardText()
	FileAppend, %A_Now%: Clipboard changed to %text%`n, %A_Temp%\7plus\Log.log
	if(text && IsObject(ClipboardList))
		ClipboardList.Push(text)
	return
}
;Stack Push function for clipboard manager stack
Stack_Push(stack,item)
{
	x:=stack.IndexOf(item)
	if(x=0)
	{
		stack.Insert(1,item)
		if(stack.MaxIndex()=11)
			stack.Delete(stack.MaxIndex())
	}
	else
	{
		stack.Move(x,1)
	}
}

ClipboardManagerMenu()
{
	global ClipboardList
	Menu, ClipboardMenu, add, 1,ClipboardHandler1
	Menu, ClipboardMenu, DeleteAll
	loop % ClipboardList.MaxIndex()
	{		
		i:=A_Index ;ClipboardList.MaxIndex()-A_Index+1
		
		x:=ClipboardList[i]
		StringReplace,x,x,`r,,All
		StringReplace,x,x,`n,[NEWLINE],All
		y:="`t"
		StringReplace,x,x,%y%,[TAB],All ;Weird syntax bug requires `t to be stored in a variable here
		x:="&" (A_Index-1) ": " Substr(x,1,100)
		if(x)
			Menu, ClipboardMenu, add, %x%, ClipboardHandler%i%
	}
	Menu, ClipboardMenu, Show
}

;Need separate handlers because menu index doesn't have to match array index
ClipboardHandler1:
ClipboardMenuClicked(1)
return
ClipboardHandler2:
ClipboardMenuClicked(2)
return
ClipboardHandler3:
ClipboardMenuClicked(3)
return
ClipboardHandler4:
ClipboardMenuClicked(4)
return
ClipboardHandler5:
ClipboardMenuClicked(5)
return
ClipboardHandler6:
ClipboardMenuClicked(6)
return
ClipboardHandler7:
ClipboardMenuClicked(7)
return
ClipboardHandler8:
ClipboardMenuClicked(8)
return
ClipboardHandler9:
ClipboardMenuClicked(9)
return
ClipboardHandler10:
ClipboardMenuClicked(10)
return

ClipboardMenuClicked(index)
{
	global ClipboardList,MuteClipboardList
	if(ClipboardList[index])
	{
		ClipboardBackup:=ClipboardAll
		MuteClipboardList:=true
		Clipboard:=ClipboardList[index]
		Clipwait,1,1
		if(!Errorlevel)
		{
			Sleep 100 ;Some extra sleep to increase reliability
			if(WinActive("ahk_class ConsoleWindowClass"))
			{
				CoordMode, Mouse, Screen
				MouseGetPos, mx,my
				CoordMode, Mouse, Relative
				Click Right 40, 40
				CoordMode, Mouse, Screen
				MouseMove, %mx%, %my%
				Send {Down 3}{Enter}
			}
			else	
				Send ^v
			Sleep 20
		}
		else
			Notify("Error pasting text", "Error pasting text", "5", "GC=555555 TC=White MC=White",NotifyIcons.Error)
		Clipboard:=ClipboardBackup
		Clipwait,1,1
		MuteClipboardList:=false
	}
	Menu, ClipboardMenu, add, 1,ClipboardHandler1
	Menu, ClipboardMenu, DeleteAll
}

;Creates a file for pasting text/image in explorer
CreateFileFromClipboard()
{
	global MuteClipboardList
	static CF_HDROP := 0xF
	;outputdebug CreateFile
	if(!MuteClipboardList)
	{
		MuteClipboardList:=true
		if(!DllCall("IsClipboardFormatAvailable", "Uint", CF_HDROP))
		{
			If(DllCall("IsClipboardFormatAvailable", "Uint", 2) && Settings.Explorer.PasteImageAsFileName !="" )
			{
				outputdebug image in clipboard			
				PasteImageAsFilePath := A_Temp "\" Settings.Explorer.PasteImageAsFileName
				success := WriteClipboardImageToFile(PasteImageAsFilePath, Settings.Misc.ImageQuality)
				if(success)
					CopyToClipboard(PasteImageAsFilePath, false)
			}
			else if (DllCall("IsClipboardFormatAvailable", "Uint", 1) && Settings.Explorer.PasteTextAsFileName !="" )
			{
				outputdebug text in clipboard
				PasteTextAsFilePath := A_Temp "\" Settings.Explorer.PasteTextAsFileName
				success := WriteClipboardTextToFile(PasteTextAsFilePath)
				if(success)
					CopyToClipboard(PasteTextAsFilePath, false)
			}
		}
		else
			outputdebug a file is already in the clipboard
		MuteClipboardList:=false
	}
}
;Read real text (=not filenames, when CF_HDROP is in clipboard) from clipboard
ReadClipboardText()
{
	if((!A_IsUnicode && DllCall("IsClipboardFormatAvailable", "Uint", 1)) || (A_IsUnicode && DllCall("IsClipboardFormatAvailable", "Uint", 13))) ;CF_TEXT = 1 ;CF_UNICODETEXT = 13
	{
		DllCall("OpenClipboard", "Ptr", 0)	
		htext:=DllCall("GetClipboardData", "Uint", A_IsUnicode ? 13 : 1, "Ptr")
		ptext := DllCall("GlobalLock", "Ptr", htext)
		text := StrGet(pText, A_IsUnicode ? "UTF-16" : "cp0")
		DllCall("GlobalUnlock", "Ptr", htext)
		DllCall("CloseClipboard")
	}
	return text
}

;Reads an image from clipboard, saves it to a file, and puts CF_HDROP structure in clipboard for file pasting
WriteClipboardImageToFile(path,Quality="")
{
	if(!Quality)
		Quality:=Settings.Misc.ImageQuality
	pBitmap:=Gdip_CreateBitmapFromClipboard()
	if(pBitmap > 0)
	{
		Gdip_SaveBitmapToFile(pBitmap, path, Settings.Misc.ImageQuality)
		return 1
	}
	return -1
}

WriteClipboardTextToFile(path)
{
	text:=ReadClipboardText()
	if(!text)
		return -1
	if(FileExist(path))
		FileDelete, %path%
	FileAppend , %text%, %path%, % A_IsUnicode ? "UTF-8" : ""
	return 1 
}

;Copies a list of files (separated by new line) to the clipboard so they can be pasted in explorer
/* Example clipboard data:
00000000   14 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00    ................
00000010   01 00 00 00 43 00 3A 00 5C 00 62 00 6F 00 6F 00    ....C.:.\.b.o.o.				<-- I believe the 01 byte at the start of this line could indicate unicode?
00000020   74 00 6D 00 67 00 72 00 00 00 43 00 3A 00 5C 00    t.m.g.r...C.:.\.
00000030   62 00 6C 00 61 00 2E 00 6C 00 6F 00 67 00 00 00    b.l.a...l.o.g...
00000040   00 00                                             										..     

typedef struct _DROPFILES {
  DWORD pFiles;
  POINT pt;
  BOOL  fNC;
  BOOL  fWide;
} DROPFILES, *LPDROPFILES;

_DROPFILES struct: 20 characters at the start
null-terminated filename list, and double-null termination at the end
*/
CopyToClipboard(files, clear, cut=0){
	static CF_HDROP := 0xF
	FileCount:=0
	PathLength:=0
	;Count files and total string length
	Loop, Parse, files, `n,`r  ; Rows are delimited by linefeeds (`r`n).
	{
		FileCount++
		PathLength+=StrLen(A_LoopField)
	}
	pid:=DllCall("GetCurrentProcessId","Uint")
	hwnd:=WinExist("ahk_pid " . pid)
	DllCall("OpenClipboard", "Ptr", hwnd)
	hPath := DllCall("GlobalAlloc", "uint", 0x42, "uint", 20 + (PathLength + FileCount + 1) * 2, "Ptr")      ; 0x42 = GMEM_MOVEABLE(0x2) | GMEM_ZEROINIT(0x40)
	pPath := DllCall("GlobalLock", "Ptr", hPath)                   ; Lock the moveable memory, retrieving a pointer to it.
	NumPut(20, pPath+0), pPath += 16 ; DROPFILES.pFiles = offset of file list
	NumPut(1, pPath+0), pPath += 4 ;fWide = 0 -->ANSI, fWide = 1 -->Unicode
	Offset:=0
	Loop, Parse, files, `n,`r  ; Rows are delimited by linefeeds (`r`n).
		offset += StrPut(A_LoopField, pPath+offset,StrLen(A_LoopField)+1,"UTF-16") * 2
	DllCall("GlobalUnlock", "Ptr", hPath)  
	;hPath must not be freed! ->http://msdn.microsoft.com/en-us/library/ms649051(VS.85).aspx
	if clear
	{
		DllCall("EmptyClipboard")                                                              ; Empty the clipboard, otherwise SetClipboardData may fail.
		Clipwait, 1, 1
	}
	result:=DllCall("SetClipboardData", "uint", CF_HDROP, "Ptr", hPath) ; Place the data on the clipboard. CF_HDROP=0xF
	Clipwait, 1, 1
	
	;Write Preferred DropEffect structure to clipboard to switch between copy/cut operations
 	mem := DllCall("GlobalAlloc","UInt",0x42,"UInt",4, "Ptr")  ; 0x42 = GMEM_MOVEABLE(0x2) | GMEM_ZEROINIT(0x40)
	str := DllCall("GlobalLock","Ptr",mem) 
	if(!cut)
		DllCall("RtlFillMemory","UInt",str,"UInt",1,"UChar",0x05) 
	else 
		DllCall("RtlFillMemory","UInt",str,"UInt",1,"UChar",0x02) 
	DllCall("GlobalUnlock","Ptr",mem)
	cfFormat := DllCall("RegisterClipboardFormat","Str","Preferred DropEffect") 
	result:=DllCall("SetClipboardData","UInt",cfFormat,"Ptr",mem)
	Clipwait, 1, 1
	DllCall("CloseClipboard")
	;mem must not be freed! ->http://msdn.microsoft.com/en-us/library/ms649051(VS.85).aspx
	return
}

;Appends files to CF_HDROP structure in clipboard
AppendToClipboard( files, cut=0) { 
	DllCall("OpenClipboard", "Ptr", 0)
	if(DllCall("IsClipboardFormatAvailable", "Uint", 1)) ;If text is stored in clipboard, clear it and consider it empty (even though the clipboard may contain CF_HDROP due to text being copied to a temp file for pasting)
		DllCall("EmptyClipboard")
	DllCall("CloseClipboard")
	txt:=clipboard (clipboard = "" ? "" : "`n") files
	Sort, txt , U ;Remove duplicates
	CopyToClipboard(txt, true, cut)
	return
}

;Writes image data from file to clipboard
Gdip_ImageToClipboard(Filename) 
{
  pBitmap := Gdip_CreateBitmapFromFile(Filename) 
  if !pBitmap 
      return 
  hbm := Gdip_CreateHBITMAPFromBitmap(pBitmap) 
  Gdip_DisposeImage(pBitmap) 
  if !hbm 
      return 
  if hdc := DllCall("CreateCompatibleDC","Ptr",0) 
  { 
      ; Get BITMAPINFO. 
      VarSetCapacity(bmi,40,0), NumPut(40,bmi) 
      DllCall("GetDIBits","Ptr",hdc,"Ptr",hbm,"uint",0 
           ,"uint",0,"uint",0,"uint",&bmi,"uint",0) 
      ; GetDIBits seems to screw up and give the image the BI_BITFIELDS 
      ; (i.e. colour-indexed) compression type when it is in fact BI_RGB. 
      NumPut(0,bmi,16) 
      ; Get bitmap bits. 
      if size := NumGet(bmi,20) 
      { 
          VarSetCapacity(bits,size) 
          DllCall("GetDIBits","Ptr",hdc,"Ptr",hbm,"uint",0 
              ,"uint",NumGet(bmi,8),"uint",&bits,"uint",&bmi,"uint",0) 
          ; 0x42 = GMEM_MOVEABLE(0x2) | GMEM_ZEROINIT(0x40) 
          hMem := DllCall("GlobalAlloc","uint",0x42,"uint",40+size, "Ptr") 
          pMem := DllCall("GlobalLock","Ptr",hMem) 
          DllCall("RtlMoveMemory","uint",pMem,"uint",&bmi,"uint",40) 
          DllCall("RtlMoveMemory","uint",pMem+40,"uint",&bits,"uint",size) 
          DllCall("GlobalUnlock","Ptr",hMem) 
      } 
      DllCall("DeleteDC","Ptr",hdc) 
  } 
  if hMem 
  { 
  		DllCall("OpenClipboard", "Ptr", 0)
      ; Place the data on the clipboard. CF_DIB=0x8 
      if ! DllCall("SetClipboardData","uint",0x8,"Ptr",hMem) 
          DllCall("GlobalFree","Ptr",hMem) 
      DllCall("CloseClipboard")
  } 
} 
