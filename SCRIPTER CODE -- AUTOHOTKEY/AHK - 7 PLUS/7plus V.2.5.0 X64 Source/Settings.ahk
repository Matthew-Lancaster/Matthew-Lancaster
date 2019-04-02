#include %A_ScriptDir%\WindowsSettings.ahk
;---------------------------------------------------------------------------------------------------------------
; The following functions create the GUI and are only called once at startup
;---------------------------------------------------------------------------------------------------------------
Settings_CreateIntroduction(ByRef TabCount) {
	global
	local yIt,x1,x2,x,y,hAdmin, Languages
	x1:=xBase
	yIt:=yBase
	x2:=x1+180
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vIntroductionTab
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	Gui, 1:Add, Text, x%x1% y%yIt%, 
	(
Welcome to 7plus! If you are new to this program, here are some tips:

 - Be sure to check out the events settings page (or more specifically, the subpages for specific categories).
   The event system allows to create all kinds of functions (hotkeys, timers, context menu entries...).
   If you look for a specific feature, use the search field on that page. To edit an event, just double-click it.
   Use the help buttons in the "Edit Event" and "Edit Subevent" windows for help on specific triggers/conditions/actions.
   
 - You should also check out the Accessor settings. The Accessor is a launcher program that can be used 
   to launch programs with the keyboard (and much more!).
   
 - For explorer features, check out the Explorer, Fast Folders and Explorer Tabs pages,
   in addition to the explorer-related events.
 
Finally, here are some settings that you're likely to change at the beginning:
	)	
	yIt+=checkboxstep * 10
	if(!IsPortable)
	{
		Gui, 1:Add, Checkbox, x%x1% y%yIt% vAutorun, Autorun 7plus on windows startup
		yIt+=checkboxstep
	}
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vHideTrayIcon, Hide Tray Icon (press WIN + H (default settings) to show settings!)
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vAutoUpdate, Automatically look for updates on startup
	yIt+=checkboxstep
	y:=yIt+TextBoxCheckBoxOffset
	Gui, 1:Add, Text, x%x1% y%y%, Run as admin:
	Gui, 1:Add, DropDownList, x%x2% y%yIt% w%wTBMedium% vRunAsAdmin hwndhAdmin, Always/Ask|Never
	AddToolTip(hAdmin, "Required for explorer buttons, Autoupdate and for accessing programs which are running as admin. Also make sure that 7plus has write access to its config files when not running as admin.")
	yIt+=textboxstep
	y:=yIt+TextBoxCheckBoxOffset
	Gui, 1:Add, Text, x%x1% y%y%, Documentation language:	
	Loop % Languages.Languages.MaxIndex()
		Languages .= (A_Index = 1 ? "" : "|") Languages.Languages[A_Index].FullName
	Gui, 1:Add, DropDownList, x%x2% y%yIt% w%wTBMedium% hwndhLanguage, % Languages
}
Settings_CreateEvents(ByRef TabCount) {
	global
	local yIt,x1,x2,x,y
	x1:=xBase+10
	x2 := x1 + 530
	yIt:=yBase
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vEventsTab, 
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	Gui, 1:Add, Text, x%x1% y%yIt% w600, You can add events here that are triggered under certain conditions. When triggered, the event can launch a series of actions. This is a very powerful tool to add all kinds of features, and many features from 7plus are now implemented with this system.
	yIt+=41
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vShowAdvancedEvents gShowAllEvents, Show advanced events
	Gui, 1:Add, Text, x+150 y%yIt%, Event search:
	yIt-=4
	Gui, 1:Add, Edit, x+10 y%yIt% w160 hwndEventFilter gEventFilterChange
	yIt += textboxstep
	Gui, 1:Add, ListView, x%x1% y%yIt% w520 h263 vGUI_EventsList gGUI_EventsList_SelectionChange Grid -LV0x10 AltSubmit Checked, Enabled|ID|Trigger|Name
	OnMessage(0x0111, "WM_COMMAND")
	Gui, 1:Add, Edit, x%x1% y+10 w520 h50 vGUI_EventsDescription +ReadOnly
	Gui, 1:Add, Button, x%x2% y%yIt% w80 vGUI_EventsList_Add gGUI_EventsList_Add, Add Event
	yIt += textboxstep
	Gui, 1:Add, Button, x%x2% y%yIt% w80 vGUI_EventsList_Remove gGUI_EventsList_Remove, Delete Event
	yIt += textboxstep
	Gui, 1:Add, Button, x%x2% y%yIt% w80 vGUI_EventsList_Edit gGUI_EventsList_Edit, Edit Event
	yIt += textboxstep
	Gui, 1:Add, Button, x%x2% y%yIt% w80 vGUI_EventsList_Copy gGUI_EventsList_Copy, Copy Event
	yIt += textboxstep
	Gui, 1:Add, Button, x%x2% y%yIt% w80 vGUI_EventsList_Paste gGUI_EventsList_Paste Disabled, Paste Event
	yIt += textboxstep * 2
	Gui, 1:Add, Button, x%x2% y%yIt% w80 vGUI_EventsList_Import gGUI_EventsList_Import, Import
	yIt += textboxstep
	Gui, 1:Add, Button, x%x2% y%yIt% w80 vGUI_EventsList_Export gGUI_EventsList_Export, Export
	yIt += textboxstep
	Gui, 1:Add, Button, x%x2% y%yIt% w80 gGUI_EventsList_Help, Help
	yIt += textboxstep + 4
	y := yIt + TextBoxTextOffset
}
Settings_CreateAccessorKeywords(ByRef TabCount) {
	global
	local yIt,x1,x2,x,y, hKeywords, hCommands
	x1:=xBase+10
	x2 := x1 + 530
	yIt:=yBase
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vAccessorKeywordsTab, 
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	Gui, 1:Add, ListView, x%x1% y%yIt% w520 h232 vGUI_AccessorKeywordsList gGUI_AccessorKeywordsList_Events hwndhwndKeywords Grid -Multi -LV0x10 AltSubmit, ID|Key|Command
	WinSet, ExStyle, +0x10, ahk_id %hwndKeywords%
	LV_ModifyCol(1, 0)
    LV_ModifyCol(2, 100)
	LV_ModifyCol(3, "AutoHdr")
	Gui, 1:Add, Button, x%x2% y%yIt% w80 vGUI_AccessorKeywords_Add gGUI_AccessorKeywords_Add, Add keyword
	yIt += textboxstep
	Gui, 1:Add, Button, x%x2% y%yIt% w80 vGUI_AccessorKeywords_Delete gGUI_AccessorKeywords_Delete, Delete keyword
	yIt += textboxstep + 4
	Gui, 1:Add, Text, x%x1% y+190, Keyword:
	Gui, 1:Add, Edit, x+27 y+-17 w450 vGUI_AccessorKeywordsKey gGUI_AccessorKeywordsTextChange hwndhKeywords
	AddToolTip(hKeywords, "The keyword which is typed into accessor at the start of the query, i.e. ""Google""")
	Gui, 1:Add, Text, x%x1% y+10, Command:
	Gui, 1:Add, Edit, x+21 y+-17 w450 vGUI_AccessorKeywordsCommand gGUI_AccessorKeywordsTextChange hwndhCommands
	AddToolTip(hCommands, "You can use parameters here which are inserted into the command at specific places. This is currently only supported by the URL plugin. Example: Keyword: ""google"" Command: ""www.google.com/search?q=${1}"" Entered Text: ""google 7plus"" result: ""www.google.com/search?q=7plus""")
}
Settings_CreateAccessor(ByRef TabCount) {
	global
	local yIt,x1,x2,x,y
	x1:=xBase+10
	x2 := x1 + 530
	yIt:=yBase
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vAccessorPluginsTab, 
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	Gui, 1:Add, ListView, x%x1% y%yIt% w520 h232 vGUI_AccessorPluginsList gGUI_AccessorPluginsList_Events Grid -LV0x10 AltSubmit Checked, Enabled|ID|Name
	Gui, 1:Add, Button, x%x2% y%yIt% w80 vGUI_AccessorSettings gGUI_AccessorSettings, Plugin Settings
	yIt += textboxstep
	Gui, 1:Add, Button, x%x2% y%yIt% w80 gGUI_Accessor_Help, Help
	Gui, 1:Add, Text, x%x1% y+190, Accessor is a versatile tool that is used to perform many commands through the keyboard, `nlike launching programs, switching windows, open URLs, browsing the filesystem,...`nPress the assigned hotkey (Default: ALT+Space) and start typing!
}
Settings_CreateExplorer(ByRef TabCount) {
	global
	local yIt,x1,x2,x,hScrollUnderMouse
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vExplorerTab, 
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	yIt:=yBase
	
	x1:=xHelp+10
	x2:=xBase+280
	x:=xBase+247
	y:=yIt+TextBoxCheckBoxOffset
	Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghNavigation vURL_Navigation1, ?
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vHKMouseGestures, Hold right mouse and click left: Go back, Hold left mouse and click right: Go forward
	yIt+=checkboxstep
	
	Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghSelectFirstFile vURL_SelectFirstFile, ?
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vHKSelectFirstFile, Explorer automatically selects the first file when you enter a directory
	yIt+=checkboxstep	
	Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghSelectFirstFile vURL_SelectFirstFile1, ?
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vHKImproveEnter, Files which are only focussed but not selected can be executed by pressing enter
	yIt+=checkboxstep		
	Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghScrollUnderMouse vURL_ScrollUnderMouse, ?
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vScrollUnderMouse hwndhScrollUnderMouse, Scroll explorer scrollbars with mouse over them
	AddToolTip(hScrollUnderMouse, "This makes it possible to scroll the file tree or the file list when another part of the explorer window is focused.")
	yIt+=checkboxstep
	Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghShowSpaceAndSize vURL_ShowSpaceAndSize, ?
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vHKShowSpaceAndSize, Show free space and size of selected files in status bar like in XP (7 only)
	yIt+=checkboxstep	
	Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghApplyOperation vURL_ApplyOperation, ?
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vHKAutoCheck, Automatically check "Apply to all further operations" checkboxes in file operations (Vista/7 only)
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vRecallExplorerPath, Win+E: Open explorer in last active directory
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vAlignExplorer, Win+E + explorer window active: Open new explorer and align them left and right
	yIt+=checkboxstep*2
	
	Gui, 1:Add, Text, x%x1% y%yIt%, Text and images from clipboard can be pasted as file in explorer with these settings
	yIt+=checkboxstep
	
	y:=yIt+TextBoxCheckBoxOffset
	Gui, 1:Add, CheckBox, x%x1% y%y% gtxt, Paste text as file
	Gui, 1:Add, Text, y%y% x%xhelp% cBlue ghPasteAsFile vURL_PasteAsFile, ?
	
	x:=xBase+232
	Gui, 1:Add, Text, x%x% y%y%, Filename:
	y:=yIt
	Gui, 1:Add, Edit, x%x2% y%y% w%wTBMedium% vTxtName R1
	yIt+=textboxstep
	
	y:=yIt+TextBoxCheckBoxOffset
	Gui, 1:Add, CheckBox, x%x1% y%y% gimg, Paste image as file
	Gui, 1:Add, Text, y%y% x%xhelp% cBlue ghPasteAsFile vURL_PasteAsFile1, ?
	y:=yIt+TextBoxTextOffset
	Gui, 1:Add, Text, x%x% y%y%, Filename:
	y:=yIt
	Gui, 1:Add, Edit, x%x2% y%y% w%wTBMedium% vImgName R1	
	yIt+=textboxstep	
}
Settings_CreateFastFolders(ByRef TabCount) {
	global
	local yIt,x1,x,y, hCleanFolderBand, hRemoveAllExplorerButtons
	yIt:=yBase
	xHelp:=xBase
	x1:=xHelp+10
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vFastFoldersTab
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	
	x:=x1 ;+xCheckboxTextOffset
	y:=yIt ;+yCheckboxTextOffset
	Gui, 1:Add, Text, x%x% y%y% R2, In explorer and file dialogs you can store a path in one of ten slots by pressing CTRL`nand a numpad number key (default settings), and restore it by pressing the numpad number key again
	yIt+=checkboxstep*1.5
	if(!IsPortable && A_IsAdmin)
	{
		Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghFastFolders1 vURL_FastFolders11, ?
		Gui, 1:Add, Checkbox, x%x% y%yIt% vHKFolderBand, Integrate Fast Folders into explorer folder band bar (Vista/7 only)	
		yIt+=checkboxstep
		Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghFastFolders2 vURL_FastFolders2, ?
		Gui, 1:Add, Checkbox, x%x% y%yIt% vHKCleanFolderBand hwndhCleanFolderBand, Remove windows folder band buttons (Vista/7 only)
		yIt+=checkboxstep
		AddToolTip(hCleanFolderBand, "If you use the folder band as a favorites bar like in browsers, it is recommended that you get rid of the buttons predefined by windows whereever possible (such as Slideshow, Add to Library,...)")
		Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghFastFolders2 vURL_FastFolders21, ?
		Gui, 1:Add, Checkbox, x%x% y%yIt% vHKPlacesBar, Integrate Fast Folders into open/save dialog places bar (First 5 Entries)
		yIt+=checkboxstep
		Gui, 1:Add, Button, x%x% y%yIt% gRemoveAllExplorerButtons hwndhRemoveAllExplorerButtons, Remove custom Explorer buttons
		AddToolTip(hRemoveAllExplorerButtons, "By doing this all custom buttons in the explorer folder band bar will be removed. This is useful if an error occurred and some buttons get duplicated. Once you press OK or Apply in this dialog, the buttons created with an ExplorerButton trigger will reappear. To make the FastFolder buttons reappear, save a directory to a FastFolder slot by pressing CTRL+Numpad[0-9] (Default keys)")
	}
	else
		Gui, 1:Add, Text, x%x% y%yIt%, Explorer bar functions don't work in portable mode and require admin priviledges!
}
Settings_CreateTabs(ByRef TabCount) {
	global
	local yIt,x1,x,y,x2
	yIt:=yBase
	xHelp:=xBase
	x1:=xHelp+10
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vExplorerTabsTab
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	
	Gui, 1:Add, Text, x%x1% y%yIt% R3, 7plus makes it possible to use tabs in explorer. New tabs are opened with the middle mouse button`nand with CTRL+T, Tabs are cycled by clicking the Tabs or pressing CTRL+(SHIFT)+TAB,`nand closed by middle clicking a tab and with CTRL+W
	yIt+=CheckboxStep*2.25
	Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghTabs vURL_Tabs, ?		
	Gui, 1:Add, Checkbox, x%x1% y%yIt% gUseTabs vUseTabs,Use Tabs in Explorer
	yIt+=checkboxstep
	x:=x1+xCheckboxTextOffset
	xhelp+=xCheckboxTextOffset
	x2:=x+240
	y:=yIt+TextBoxCheckBoxOffset
	Gui, 1:Add, Text, x%x% y%y% vTabLabel1, Create tabs:
	Gui, 1:Add, DropDownList, x%x2% y%yIt% w%wTBMedium% vNewTabPosition AltSubmit,next to current tab|at the end
	yIt+=textboxstep
	y:=yIt+TextBoxCheckBoxOffset
	Gui, 1:Add, Text, x%x% y%y% vTabLabel2, Tab startup path (empty for current dir):
	Gui, 1:Add, Edit, x%x2% y%yIt% vTabStartupPath w%wTBMedium% R1,%TabStartupPath%
	x2+=wTBMedium+10
	y:=yIt+TextBoxButtonOffset
	Gui, 1:Add, Button, x%x2% y%yIt% w%wButton% vTabStartupPathBrowse gTabStartupPathBrowse,...
	x2:=x+240
	yIt+=textboxstep
	Gui, 1:Add, Checkbox, x%x% y%yIt% vActivateTab,Activate tab on tab creation
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x% y%yIt% vTabWindowClose,Close all tabs when window is closed
	yIt+=checkboxstep	
	y:=yIt+TextBoxCheckBoxOffset
	Gui, 1:Add, Text, x%x% y%y% vTabLabel3, On tab close:
	Gui, 1:Add, DropDownList, x%x2% y%yIt% w%wTBMedium% vOnTabClose AltSubmit,activate left tab|activate right tab
	yIt+=textboxstep
}
Settings_CreateFTPProfiles(ByRef TabCount) {
	global
	local yIt,x1,x,y,x2
	yIt:=yBase
	xHelp:=xBase
	x1:=xHelp+10
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vFTPProfilesTab
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	
	Gui, 1:Add, Text, x%x1% y%yIt% R3, You can define FTP profiles for use with the upload action here.
	yIt+=CheckboxStep
	Gui, Add, DropDownList, x%x1% y%yIt% w200 vFTPProfilesDropDownList gFTPProfilesDropDownList,
	Gui, Add, Button, x+10 w80 gFTPProfiles_Add, Add profile
	Gui, Add, Button, x+10 w80 vFTPProfiles_Delete gFTPProfiles_Delete, Delete profile
	yIt += CheckboxStep * 2
	y:=yIt+yCheckboxTextOffset
	x2 := x1 + 100
	Gui, Add, Text, x%x1% y%yIt%, Hostname:
	Gui, Add, Edit, x%x2% y%y% w200 vFTPProfiles_Hostname gFTPProfiles_Hostname,
	yIt += TextboxStep
	y:=yIt+yCheckboxTextOffset
	Gui, Add, Text, x%x1% y%yIt%, Port:
	Gui, Add, Edit, x%x2% y%y% w40 Number vFTPProfiles_Port,
	yIt += TextboxStep
	y:=yIt+yCheckboxTextOffset
	Gui, Add, Text, x%x1% y%yIt%, User:
	Gui, Add, Edit, x%x2% y%y% w200 vFTPProfiles_User,
	yIt += TextboxStep
	y:=yIt+yCheckboxTextOffset
	Gui, Add, Text, x%x1% y%yIt%, Password:
	Gui, Add, Edit, x%x2% y%y% w200 Password vFTPProfiles_Password,
	yIt += TextboxStep
	y:=yIt+yCheckboxTextOffset
	Gui, Add, Text, x%x1% y%yIt%, URL:
	Gui, Add, Edit, x%x2% y%y% w200 vFTPProfiles_URL,
	yIt += TextboxStep
	Gui, 1:Add, Text, x%x1% y%yIt% R3, Target folder and filename are set separately for each event that uses the FTP upload function on the Events page.
}
Settings_CreateHotstrings(ByRef TabCount) {
	global
	local yIt,x1,x2,x,y
	xHelp:=xBase
	x1:=xHelp+10
	x2 := x1 + 520
	yIt:=yBase
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vHotstringsTab, 
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	Gui, 1:Add, ListView, x%x1% y%yIt% w510 h216 vGUI_HotstringsList gGUI_HotstringsList_Events Grid -Multi -LV0x10 AltSubmit, ID|Key|Command
	LV_ModifyCol(1, 0)
    LV_ModifyCol(2, 100)
	LV_ModifyCol(3, "AutoHdr")
	Gui, 1:Add, Button, x%x2% y%yIt% w100 vGUI_Hotstrings_Add gGUI_Hotstrings_Add, Add HotString
	yIt += textboxstep
	Gui, 1:Add, Button, x%x2% y%yIt% w100 vGUI_Hotstrings_Delete gGUI_Hotstrings_Delete, Remove HotString
	yIt += textboxstep
	Gui, 1:Add, Button, x%x2% y%yIt% w100 vGUI_Hotstrings_RegexHelp gGUI_Hotstrings_RegexHelp, RegEx Help
	Gui, 1:Add, Text, x%x1% y+144, HotString:
	Gui, 1:Add, Edit, x+27 y+-17 w440 vGUI_HotstringsKey gGUI_HotstringsTextChange
	Gui, 1:Add, Text, x%x1% y+10, Value:
	Gui, 1:Add, Edit, x+41 y+-17 w440 vGUI_HotstringsValue gGUI_HotstringsTextChange
	Gui, 1:Add, Text, x%x1% y+10, HotStrings are used to expand abbreviations and acronyms, such as "btw" -> "by the way". They support regular`nexpressions in PCRE format. If you want a HotString to trigger only when typed as a seperate word, prepend \b`nand append \s.  For case-insensitive HotStrings, put i) at the start. You can also use keys like {Enter}.
}
Settings_CreateWindows(ByRef TabCount) {
	global
	local yIt, x1, x, y, hSlideWindows, hUpdate
	xHelp:=xBase
	x1:=xHelp+10
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vWindowsTab, 
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	yIt:=yBase
	Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghSlideWindow vURL_SlideWindow, ?
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vHKSlideWindows gSlideWindow hwndhSlideWindows, WIN + SHIFT + Arrow keys: Slide Window function
	AddToolTip(hSlideWindows, "A Slide Window is moved off screen and will not be shown until you activate it through task bar / ALT + TAB or move the mouse to the border where it was hidden. It will then slide into the screen, and slide out again when the mouse leaves the window or when another window gets activated. Deactivate this mode by moving the window or pressing WIN+SHIFT+Arrow key in another direction.")
	yIt+=checkboxstep
	x:=x1+xCheckboxTextOffset
	Gui, 1:Add, Checkbox, x%x% y%yIt% vSlideWinHide, Hide Slide Windows in taskbar and from ALT + TAB
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x% y%yIt% vSlideWindowSideLimit, Allow only one Slide Window per screen side
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x% y%yIt% vSlideWindowRequireMouseUp hwndhRequireMouseUp, Require left mouse button to be up to activate slide window at screen border
	AddToolTip(hRequireMouseUp, "This feature is used to prevent accidently activating a slide window while dragging with the mouse. It's still possible to drag something to the slide window by holding the modifier key which is set below.")
	yIt+=checkboxstep
	Gui, 1:Add, Text, y%yIt% x%x%, Slide Windows modifier key:
	y:=yIt+yCheckboxTextOffset
	x:=x1+200
	Gui, 1:Add, DropDownList, x%x% y%y% vSlideWindowsModifier hwndhSlideModifier, Control|ALT|Shift|WIN
	AddToolTip(hSlideModifier, "If this key is pressed, the mouse may be moved out of the currently active slide window without sliding it out. This is useful if the slide window has child windows that don't overlap with the main window. If the option above is enabled, it may also be used to drag something into a hidden slide window by moving the mouse to the screen border and holding this key.")
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vShowResizeTooltip, Show window size as tooltip while resizing
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vAutoCloseWindowsUpdate hwndhUpdate, Automatically close Windows Update reboot notification dialog
	AddToolTip(hUpdate, "If you enable this setting you will not be able to open this dialog anymore. You can simply reboot windows though...")
}
Settings_CreateWindowsSettings(ByRef TabCount) {
	global
	local yIt,x1,x,y
	xHelp:=xBase
	x1:=xHelp
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vWindowsSettingsTab, 
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	yIt:=yBase
	Gui, 1:Add, Text, x%x1% y%yIt%, Explorer:
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vRemoveUserDir, Remove user directory from directory tree
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vRemoveWMP, Remove Windows Media Player context menu entries (Play, Add to playlist, Buy music)
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vRemoveOpenWith, Remove "Open With Webservice or choose program" dialogs for unknown file extensions
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vShowExtensions, Always show file extensions
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vShowHiddenFiles, Show hidden files
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vShowSystemFiles, Show system files
	yIt+=checkboxstep
	if(A_OSVersion = "WIN_XP")
	{
		Gui, 1:Add, Checkbox, x%x1% y%yIt% vClassicView, Use classic explorer view
		yIt+=checkboxstep	
	}
	if(A_OSVersion != "WIN_XP" && A_OSVersion != "WIN_VISTA")
	{		
		Gui, 1:Add, Checkbox, x%x1% y%yIt% vRemoveLibraries, Remove explorer libraries (from directory tree and context menus)
		yIt+=checkboxstep
	}
	yIt+=checkboxstep
	Gui, 1:Add, Text, x%x1% y%yIt%, Windows:
	yIt+=checkboxstep	
	if(A_OSVersion != "WIN_XP" && A_OSVersion != "WIN_VISTA")
	{		
		Gui, 1:Add, Checkbox, x%x1% y%yIt% vHKActivateBehavior, Left click on task group button: cycle through windows	
		yIt+=checkboxstep
	}
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vShowAllTray, Show all tray notification icons
	yIt+=checkboxstep
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vRemoveCrashReporting, Remove crash reporting dialog
	yIt+=checkboxstep
	if(Vista7)
	{		
		Gui, 1:Add, Checkbox, x%x1% y%yIt% vDisableUAC, Disable UAC
		yIt+=checkboxstep
	}
	if(A_OSVersion != "WIN_XP" && A_OSVersion != "WIN_VISTA")
	{		
		y:=yIt+TextBoxTextOffset
		Gui, 1:Add, Text, x%x1% y%y%, Taskbar thumbnail hover time [ms]:
		Gui, 1:Add, Edit, x+10 y%yIt% w%wTBShort% R1 vHoverTime
		yIt+=textboxstep
	}
}
Settings_CreateMisc(ByRef TabCount) {
	global
	local yIt,x1, hWordDelete
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vMiscTab, 
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	x1:=xBase+10
	xhelp:=xBase
	yIt:=yBase
	Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghJoyControl vURL_JoyControl, ?
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vJoyControl, Use joystick/gamepad as remote control when not in fullscreen (optimized for XBOX360 gamepad)
	yIt+=checkboxstep
	Gui, 1:Add, Text, y%yIt% x%xhelp% cBlue ghWordDelete vURL_WordDelete, ?
	Gui, 1:Add, Checkbox, x%x1% y%yIt% vWordDelete hwndhWordDelete, Make CTRL+Backspace and CTRL+Delete work in all textboxes
	AddToolTip(hWordDelete, "Many text boxes in windows have the problem that it's not possible to use CTRL+Backspace to delete a word. Instead, it will write a square character. Enabling this will fix it.")
	yIt+=2*checkboxstep
	y:=yIt+TextBoxTextOffset
	x2:=x1+180
	Gui, 1:Add, Text, x%x1% y%y%, Image compression quality:
	Gui, 1:Add, Edit, x%x2% y%yIt% w%wTBShort% R1 vImageQuality Number
	yIt+=TextBoxStep
	y:=yIt+TextBoxTextOffset
	Gui, 1:Add, Text, x%x1% y%y%, Default image extension:
	Gui, 1:Add, Edit, x%x2% y%yIt% w%wTBShort% R1 vImageExtension
	yIt+=TextBoxStep
	Gui, 1:Add, Text, x%x1% y%yIt%, Many features of 7plus check if there is a fullscreen window active.`nYou can add window class names to include and exclude filters here to influence the fullscreen recognition.
	yIt+=CheckboxStep * 1.5
	y:=yIt+TextBoxTextOffset
	Gui, 1:Add, Text, x%x1% y%y%, Fullscreen detection include list
	Gui, 1:Add, Edit, x%x2% y%yIt% w%wTBHuge% R1 vFullscreenInclude	
	yIt+=TextBoxStep
	y:=yIt+TextBoxTextOffset
	Gui, 1:Add, Text, x%x1% y%y%, Fullscreen detection exclude list
	Gui, 1:Add, Edit, x%x2% y%yIt% w%wTBHuge% R1 vFullscreenExclude
	yIt+=TextBoxStep
}
Settings_CreateAbout(ByRef TabCount) {
	global
	local yIt,x1,x2,x,y,version
	Gui, 1:Add, Tab2, x176 y14 w460 h350 vAboutTab, 
	TabCount++
	AddTab(0, "","SysTabControl32" TabCount)
	yIt:=YBase
	x1:=XBase+10
	x2:=xBase+350
	; if(A_IsCompiled)			
		; Gui, 1:Add, Picture, w128 h128 y%yIt% x%x2% Icon3 vLogo, %A_ScriptFullPath%
	; else
		Gui, 1:Add, Picture, w128 h128 y%yIt% x%x2% vLogo, %A_ScriptDir%\128.png
	
	Gui, 1:font, s20
	version := MajorVersion "." MinorVersion "." BugfixVersion
	Gui, 1:Add, Text, y%yIt% x%x1%, 7plus Version %version%
	Gui, 1:font
	yIt+=hText*3
	x2:=x1+100
	Gui, 1:Add, Text, y%yIt% x%x1% , Project page:
	Gui, 1:Add, Text, y%yIt% x%x2% cBlue gProjectpage vURL_Projectpage, http://code.google.com/p/7plus/
	yIt+=hText
	Gui, 1:Add, Text, y%yIt% x%x1% , Report bugs:
	Gui, 1:Add, Text, y%yIt% x%x2% cBlue gBugtracker vURL_Bugtracker, http://code.google.com/p/7plus/issues/list
	yIt+=hText
	Gui, 1:Add, Text, y%yIt% x%x1% , Author:
	Gui, 1:Add, Text, y%yIt% x%x2% , Christian Sander
	yIt+=hText
	Gui, 1:Add, Text, y%yIt% x%x1% , E-Mail:
	Gui, 1:Add, Text, y%yIt% x%x2% cBlue gMail vURL_Mail, fragman@gmail.com
	yIt+=hText
	Gui, 1:Add, Text, y%yIt% x%x1% , Twitter:
	Gui, 1:Add, Text, y%yIt% x%x2% cBlue gTwitter vURL_Twitter, 7_plus
	yIt+=hText*2
	Gui, 1:Add, Text, y%yIt% x%x1%, To support the development of this project, please donate:
	yIt+=hText*1.5
	; if(A_IsCompiled)			
		; Gui, 1:Add, Picture, y%yIt% x%x1% cBlue gDonate Icon4 vURL_Donate, %A_ScriptFullPath%
	; else
		Gui, 1:Add, Picture, y%yIt% x%x1% cBlue gDonate vURL_Donate, %A_ScriptDir%\Donate.png		
	yIt+=hText*2
	x2:=x1+200
	Gui, 1:Add, Text, y%yIt% x%x1%, Proudly written in Autohotkey
	Gui, 1:Add, Text, y%yIt% x%x2%, Updater uses
	x2+=66
	Gui, 1:Add, Text, y%yIt% x%x2% vURL_7zip g7zip cBlue, 7-Zip
	x2+=24
	Gui, 1:Add, Text, y%yIt% x%x2%, , which is licensed under the 
	x2+=136
	Gui, 1:Add, Text, y%yIt% x%x2% vURL_LGPL gLGPL cBlue, LGPL
	yIt+=hText
	Gui, 1:Add, Text, y%yIt% x%x1% cBlue gAhk vURL_AHK, www.autohotkey.com		
	yIt+=hText*2
	Gui, 1:Add, Text, y%yIt% x%x1% , Licensed under  	
	x2:=x1+100
	Gui, 1:Add, Text, y%yIt% x%x2% cBlue gGPL vURL_GPL, GNU General Public License v3
	yIt+=hText*2
	Gui, 1:Add, Text, y%yIt% x%x1% , Credits for lots of code samples and help go out to:`nSean, HotKeyIt, majkinetor, polyethene, Lexikos, tic, TheGood, PhiLho, Temp01, Laszlo, jballi, Shrinker,`n M@x and the other guys and gals on #ahk and the forums.	
}

;---------------------------------------------------------------------------------------------------------------
; The following functions set the GUI values and are called each time the GUI is shown
;---------------------------------------------------------------------------------------------------------------
Settings_SetupIntroduction() {
	global
	local class
	GuiControl, 1:,HideTrayIcon,% HideTrayIcon = 1
	GuiControl, 1:,AutoUpdate,% AutoUpdate = 1
	;Figure out if Autorun is enabled
	if(!IsPortable)
		GuiControl, 1:, Autorun,% IsAutorunEnabled()
	GuiControl, 1:Choose, RunAsAdmin, %RunAsAdmin%
	class := HWNDToClassNN(hLanguage)
	GuiControl, 1:ChooseString, %class%, % Language.CurrentLanguage.FullName
}
Settings_SetupEvents() {
	global
	WasCritical := A_IsCritical
	Critical
	Gui, 1:Default
	Gui, ListView, GUI_EventsList
	if(!Settings_Events)
	{
		Loop % Events.MaxIndex()
			Events[A_Index].Trigger.PrepareCopy(Events[A_Index])		
		
		Settings_Events := Events.DeepCopy()	
		i := 1
		ControlSetText, ,, ahk_id %EventFilter%
	}
	RecreateTreeView()
	LV_ModifyCol(3, "AutoHdr")
	GuiControl, 1:focus, Gui_EventsList
	GuiControl, 1:enable, GUI_EventsList_Add
	GuiControl, 1:enable, GUI_EventsList_Remove
	GuiControl, 1:enable, GUI_EventsList_Edit
	GuiControl, 1:enable, GUI_EventsList_Copy
	GuiControl, 1:, ShowAdvancedEvents, % ShowAdvancedEvents = 1
	ControlFocus, SysListView321, A
	if(!WasCritical)
		Critical, Off
}
FillEventsList()
{
	global EventFilter, Settings_Events, ShowAdvancedEvents
	WasCritical := A_IsCritical
	Critical
	Gui, 1:Default
	Gui, ListView, GUI_EventsList
	i := LV_GetNext("")
	if(i)
		LV_GetText(SelectedID,i,2)
	LV_Delete()
	selected := TV_GetSelection()
	TV_GetText(Category, selected)
	parent := TV_GetParent(selected)
	if(!parent)
		Category := ""
	ControlGetText, filter,,ahk_id %EventFilter%
	count := 0
	Loop % Settings_Events.MaxIndex()
	{
		id := Settings_Events[A_Index].ID
		DisplayString := Settings_Events[A_Index].Trigger.DisplayString()
		Name := Settings_Events[A_Index].Name
		scroll := false
		if(((!filter || InStr(id, filter) || InStr(DisplayString, Filter) || InStr(Name, filter) || InStr(Settings_Events[A_Index].Description, Filter)) && (filter || !Category || Category = Settings_Events[A_Index].Category)) && (ShowAdvancedEvents || !Settings_Events[A_Index].EventComplexityLevel))
		{
			Gui, ListView, GUI_EventsList
			LV_Add((SelectedID != "" && id = SelectedID  && (scroll := 1) ? "Select Focus" : "") (Settings_Events[A_Index].Enabled ? " Check": " "), "", id, DisplayString, name)
			if(scroll)
				LV_Modify(A_Index, "Vis")
			count++
		}
	}
	if(LV_GetCount("Selected") = 0 && LV_GetCount() > 0)
		LV_Modify(1, "Select")
	if(LV_GetCount("Selected") = 1)
	{
		i:=LV_GetNext("")
		LV_GetText(id,i,2)
		Event := Settings_Events.GetItemWithValue("ID", id)
		GuiControl, 1:, GUI_EventsDescription, % Event.Description
	}
	
	if(!WasCritical)
		Critical, Off
}
RecreateTreeView()
{
	global Settings_Events, SettingsTabList, SuppressTreeViewMessages, ShowAdvancedEvents
	WasCritical := A_IsCritical
	Critical
	Gui, 1:Default
	outputdebug recreatetreeview() listview
	Gui, ListView, GUI_EventsList
	Gui, TreeView, SettingsTreeView
	i := LV_GetNext("")
	LV_GetText(id,i,2)
	selected := TV_GetSelection()
	TV_GetText(CurrentCategory, selected)
	SuppressTreeViewMessages := true
	TV_Delete()
	TV_Add("Introduction") ;Note: Treeview is also created once in initial setup
	EventsTreeViewEntry := TV_Add("All Events", "", "Expand Select Vis" )
	Loop % Settings_Events.Categories.MaxIndex()
	{
		Category := Settings_Events.Categories[A_Index]
		Loop % Settings_Events.MaxIndex()
		{
			if(ShowAdvancedEvents || (Settings_Events[A_Index].Category = Category && !Settings_Events[A_Index].EventComplexityLevel))
			{
				TV_Add(Category, EventsTreeViewEntry, "Sort" (CurrentCategory = Category ? " Select Vis" : ""))
				break
			}
		}
	}
	Loop, Parse, SettingsTabList, |
		if(A_LoopField != "Introduction")
			TV_Add(A_LoopField)
	FillEventsList()
	SuppressTreeViewMessages := false
	if(!WasCritical)
		Critical, Off
}
Settings_SetupAccessorKeywords() {
	global Accessor, Settings_AccessorKeywords
	WasCritical := A_IsCritical
	Critical
	Gui, 1:Default
	Gui, ListView, GUI_AccessorKeywordsList
	Settings_AccessorKeywords := Accessor.Keywords.DeepCopy()
	LV_Delete()
	Loop % Settings_AccessorKeywords.MaxIndex()
	{
		Gui, ListView, GUI_AccessorKeywordsList
		LV_Add(A_Index = 1 ? "Select" : "", A_Index, Settings_AccessorKeywords[A_Index].Key, Settings_AccessorKeywords[A_Index].Command)
	}
	if(Settings_AccessorKeywords.MaxIndex() > 0)
	{
		GuiControl, 1:enable, GUI_AccessorKeywordsKey
		GuiControl, 1:enable, GUI_AccessorKeywordsCommand
	}
	else
	{
		GuiControl, 1:disable, GUI_AccessorKeywordsKey
		GuiControl, 1:disable, GUI_AccessorKeywordsCommand
	}
	if(!WasCritical)
		Critical, Off
}
Settings_SetupAccessor() {
	global AccessorPlugins, Settings_AccessorPlugins
	WasCritical := A_IsCritical
	Critical
	Gui, 1:Default
	outputdebug Settings_SetupAccessor() listview
	Gui, ListView, GUI_AccessorPluginsList
	Settings_AccessorPlugins := Array()
	outputdebug % "setupaccessor count " accessorplugins.MaxIndex()
	LV_Delete()
	Loop % AccessorPlugins.MaxIndex()
	{
		PluginCopy := RichObject()
		PluginCopy.Enabled := AccessorPlugins[A_Index].Enabled
		; PluginCopy.Keyword := AccessorPlugins[A_Index].Keyword
		PluginCopy.Type := AccessorPlugins[A_Index].Type
		PluginCopy.Settings := AccessorPlugins[A_Index].Settings.DeepCopy()
		Settings_AccessorPlugins.Insert(PluginCopy)
		outputdebug Settings_SetupAccessor() add
		Gui, ListView, GUI_AccessorPluginsList
		LV_Add(Settings_AccessorPlugins[A_Index].Enabled ? "Check" : "", "", A_Index, Settings_AccessorPlugins[A_Index].Type)
	}	
	LV_ModifyCol(1, 60)
    LV_ModifyCol(2, 0)
	LV_ModifyCol(3, "AutoHdr")
	if(!WasCritical)
		Critical, Off
}
Settings_SetupExplorer() {
	global
	local temp
	GuiControl, 1:, HKMouseGestures,% HKMouseGestures = 1
	if(A_OSVersion!="WIN_7")
		GuiControl, 1:disable, HKShowSpaceAndSize
	
	if(!Vista7)
		GuiControl, 1:disable, HKAutoCheck
	GuiControl, 1:, ImgName, % Settings.Explorer.PasteImageAsFileName
	GuiControl, 1:, TxtName, % Settings.Explorer.PasteTextAsFileName
	
	;Setup paste text as file
	GuiControl, 1:, Paste text as file, % (Settings.Explorer.PasteTextAsFileName != "")
	GoSub txt
	
	;Setup paste image as file
	GuiControl, 1:,Paste image as file,% (Settings.Explorer.PasteImageAsFileName!="")
	GoSub img
		
	GuiControl, 1:, HKSelectFirstFile, % HKSelectFirstFile = 1
	GuiControl, 1:, HKImproveEnter, % HKImproveEnter = 1
	GuiControl, 1:, ScrollUnderMouse, % ScrollUnderMouse = 1
	GuiControl, 1:, HKShowSpaceAndSize, % HKShowSpaceAndSize = 1
	GuiControl, 1:, HKAutoCheck, % HKAutoCheck = 1
	GuiControl, 1:, RecallExplorerPath, % RecallExplorerPath = 1
	GuiControl, 1:, AlignExplorer, % AlignExplorer = 1
}
Settings_SetupFastFolders() {
	global			
	if(!Vista7)
	{
		GuiControl, 1:disable, HKFolderBand
		GuiControl, 1:disable, HKCleanFolderBand
		GuiControl, 1:disable, FolderBandDescription
	}
	GuiControl, 1:,HKFolderBand,% HKFolderBand = 1
	GuiControl, 1:,HKCleanFolderBand,% HKCleanFolderBand = 1
	GuiControl, 1:,HKPlacesBar,% HKPlacesBar = 1
	GuiControl, 1:,Use Fast Folders,% HKFastFolders = 1
	GoSub FastFolders
}
Settings_SetupTabs() {
	global
	GuiControl, 1:Choose, NewTabPosition, %NewTabPosition%
	GuiControl, 1:Choose, OnTabClose, %OnTabClose%
	GuiControl, 1:,UseTabs,% UseTabs = 1
	GoSub UseTabs
	GuiControl, 1:,ActivateTab,% ActivateTab = 1
	GuiControl, 1:,TabWindowClose,% TabWindowClose = 1
}
Settings_SetupFTPProfiles() {
	global
	local Profiles
	Settings_FTPProfiles := FTPProfiles.DeepCopy()
	Loop % FTPProfiles.MaxIndex()
		Profiles .= "|" A_Index ": " FTPProfiles[A_Index].Hostname (A_Index = 1 ? "|" : "")
	GuiControl, 1:,FTPProfilesDropDownList, %Profiles%
	GuiControl, 1:Choose, FTPProfilesDropDownList, |1 ;| -> trigger g-label
	GuiControlGet, selection,, FTPProfilesDropDownList
	if(selection)
	{
		GuiControl, enable, FTPProfiles_Hostname
		GuiControl, enable, FTPProfiles_Port
		GuiControl, enable, FTPProfiles_User
		GuiControl, enable, FTPProfiles_Password
		GuiControl, enable, FTPProfiles_URL
	}
	else
	{
		GuiControl, disable, FTPProfiles_Hostname
		GuiControl, disable, FTPProfiles_Port
		GuiControl, disable, FTPProfiles_User
		GuiControl, disable, FTPProfiles_Password
		GuiControl, disable, FTPProfiles_URL
	}
}
Settings_SetupHotstrings() {
	global Hotstrings, Settings_Hotstrings
	WasCritical := A_IsCritical
	Critical
	Gui, 1:Default
	Gui, ListView, GUI_HotstringsList
	Settings_Hotstrings := Hotstrings.DeepCopy()
	LV_Delete()
	Loop % Settings_Hotstrings.MaxIndex()
	{
		Gui, ListView, GUI_HotstringsList
		LV_Add(A_Index = 1 ? "Select" : "", A_Index, Settings_Hotstrings[A_Index].Key, Settings_Hotstrings[A_Index].Value)
	}
	if(Settings_Hotstrings.MaxIndex() > 0)
	{
		GuiControl, 1:enable, GUI_HotstringsKey
		GuiControl, 1:enable, GUI_HotstringsValue
	}
	else
	{
		GuiControl, 1:disable, GUI_HotstringsKey
		GuiControl, 1:disable, GUI_HotstringsValue
	}
	if(!WasCritical)
		Critical, Off
}
Settings_SetupWindows() {
	global
	GuiControl, 1:, HKSlideWindows, % HKSlideWindows = 1
	GuiControl, 1:, SlideWinHide, % SlideWinHide = 1
	GuiControl, 1:, SlideWindowSideLimit, % SlideWindowSideLimit = 1
	GuiControl, 1:, SlideWindowRequireMouseUp, % SlideWindowRequireMouseUp = 1
	GuiControl, 1:ChooseString, SlideWindowsModifier, %SlideWindowsModifier% ;| -> trigger g-label
	if(A_OsVersion!="WIN_7")
	{
		if(!IsPortable)
			GuiControl, 1:disable, HKActivateBehavior
	}
	GuiControl, 1:, ShowResizeTooltip, % ShowResizeTooltip = 1
	GuiControl, 1:, AutoCloseWindowsUpdate, % AutoCloseWindowsUpdate = 1
}

Settings_SetupWindowsSettings() {
	global	
	GuiControl, 1:, ShowAllTray, % WindowsSettings_Get_ShowAllTray()
	GuiControl, 1:, RemoveUserDir, %  WindowsSettings_Get_RemoveUserDir()
	GuiControl, 1:, RemoveWMP, %  WindowsSettings_Get_RemoveWMP()
	GuiControl, 1:, RemoveOpenWith, % WindowsSettings_Get_RemoveOpenWith()
	GuiControl, 1:, RemoveCrashReporting, % WindowsSettings_Get_RemoveCrashReporting()
	GuiControl, 1:, ShowExtensions, % WindowsSettings_Get_ShowExtensions()
	GuiControl, 1:, ShowHiddenFiles, % WindowsSettings_Get_ShowHiddenFiles()
	GuiControl, 1:, ShowSystemFiles, % WindowsSettings_Get_ShowSystemFiles()
	if(Vista7)
		GuiControl, 1:, DisableUAC, %  WindowsSettings_Get_DisableUAC()
	if(A_OSVersion = "WIN_XP")
		GuiControl, 1:, ClassicView, %  WindowsSettings_Get_ClassicView()
	else if(A_OSVersion != "WIN_XP" && A_OSVersion != "WIN_VISTA")
	{
		GuiControl, 1:, RemoveLibraries, % WindowsSettings_Get_RemoveLibraries()
		GuiControl, 1:, HKActivateBehavior, % WindowsSettings_Get_ActivateBehavior()
		GuiControl, 1:, HoverTime, % WindowsSettings_Get_HoverTime()
	}
}
Settings_SetupMisc() {
	global
	GuiControl, 1:, ImageQuality, %ImageQuality%
	GuiControl, 1:, ImageExtension, %ImageExtension%
	GuiControl, 1:, FullscreenInclude, %FullscreenInclude%
	GuiControl, 1:, FullscreenExclude, %FullscreenExclude%
	GuiControl, 1:,JoyControl,% JoyControl = 1
	GuiControl, 1:,WordDelete,% WordDelete = 1
}
Settings_SetupAbout() {
	global	
}

AddTab(IconNumber, TabName, TabControl) {
	Gui 1: +LastFound
	VarSetCapacity(TCITEM, 92 + A_PtrSize * 2, 0)
	NumPut(3, TCITEM ,0) ; Mask (3) comes from TCIF_TEXT(1) + TCIF_IMAGE(2).
	StrPut(TabName, &TCITEM + 12) ; pszText
	NumPut(IconNumber - 1, TCITEM ,16 + A_PtrSize) ; iImage: -1 to convert to zero-based.
	SendMessage, 0x1307, 999, &TCITEM, %TabControl%  ; 0x1307 is TCM_INSERTITEM
}
SettingsHandler:
ShowSettings(Settings.General.FirstRun ? "Introduction" : "")
return
ShowSettings(ShowPane="") {
	global
	local x,y,yIt,x1,x2,TabCount,id,text
	Critical, Off
	if(!SettingsActive)
	{
		SettingsActive:=True
		;---------------------------------------------------------------------------------------------------------------
		; Create GUI
		;---------------------------------------------------------------------------------------------------------------
		Gui, 1:Default
		Gui, 1:+OwnDialogs
		if(!SettingsInitialized)
		{
			ybase:=36
			checkboxstep:=20
			textboxstep:=30
			TextBoxCheckBoxOffset:=4
			TextBoxTextOffset:=4
			TextBoxButtonOffset:=-1
			xCheckBoxTextOffset:=17
			yCheckBoxTextOffset:=-6
			hText:=16
			yIt:=yBase
			y:=yIt
			xBase:=190
			xHelp:=xBase
			x1:=xHelp+10
			x2:=xBase+280
			wTBShort:=50
			wTBMedium:=170
			wTBLarge:=210
			wTBHuge:=300
			wButton:=30
			hCheckbox:=16 
			
			SettingsTabList := "Introduction|Accessor Keywords|Accessor Plugins|Explorer|Fast Folders|Explorer Tabs|FTP Profiles|Hotstrings|Windows|Windows Settings|Misc|About"
			;Note: Treeview entries get overwritten in RecreateTreeView()
			Gui, 1:Add, TreeView, x16 y20 w140 h420 AltSubmit gSettingsTreeView vSettingsTreeView -HScroll
			TV_Add("Introduction", "", "First")
			EventsTreeViewEntry := TV_Add("All Events", "", "Expand")
			Loop % Events.Categories.MaxIndex()
				TV_Add(Events.Categories[A_Index], EventsTreeViewEntry, "Sort")
			Loop, Parse, SettingsTabList, |
				if(A_LoopField != "Introduction")
					TV_Add(A_LoopField)
			
			Gui, 1:Add, GroupBox, x176 y14 w650 h420 vGGroupBox , Events
			Gui, 1:Add, Button, x495 y440 w70 h23 vBtnOK gOK, OK
			Gui, 1:Add, Button, x573 y440 w80 h23 vBtnCancel gCancel, Cancel
			Gui, 1:Add, Button, x660 y440 w80 h23 vBtnApply gApply, Apply
			Gui, 1:Add, Text, x16 y445 vTutLabel, Click on ? to see video tutorial help!
			Gui, 1:Add, Text, y445 x330 vWait, Applying settings, please wait!
			Settings_CreateIntroduction(TabCount)
			Settings_CreateEvents(TabCount)
			Settings_CreateAccessorKeywords(TabCount)
			Settings_CreateAccessor(TabCount)
			Settings_CreateExplorer(TabCount)
			Settings_CreateFastFolders(TabCount)
			Settings_CreateTabs(TabCount)
			Settings_CreateFTPProfiles(TabCount)
			Settings_CreateHotstrings(TabCount)
			Settings_CreateWindows(TabCount)
			Settings_CreateWindowsSettings(TabCount)
			Settings_CreateMisc(TabCount)
			Settings_CreateAbout(TabCount)			
			SettingsInitialized := true
		}
		GuiControl, 1:Hide, Wait
		GuiControl, 1:Choose,SettingsTreeView,Events
		GoSub SettingsTreeView
		DetectHiddenWindows, On
		Gui 1:+LastFoundExist
				
		;---------------------------------------------------------------------------------------------------------------
		; Setup Control Status
		;---------------------------------------------------------------------------------------------------------------
		Settings_SetupIntroduction()
		Settings_SetupEvents()
		Settings_SetupAccessorKeywords()
		Settings_SetupAccessor()
		Settings_SetupExplorer()
		Settings_SetupFastFolders()
		Settings_SetupTabs()
		Settings_SetupFTPProfiles()
		Settings_SetupHotstrings()
		Settings_SetupWindows()
		Settings_SetupWindowsSettings()
		Settings_SetupMisc()
		Settings_SetupAbout()

		Gui, 1:Show, x338 y159 h474 w850, 7plus Settings
		
		;Code for URL hand cursor, don't touch :D
		;Hand cursor over controls where the assigned variable starts with URL_
		; Retrieve scripts PID 
		Process, Exist 
		pid_this := ErrorLevel 

		; Retrieve unique ID number (HWND/handle) 
		WinGet, hw_Gui, ID, ahk_class AutoHotkeyGUI ahk_pid %pid_this% 
		
		OnMessage(0x100, "WM_KEYDOWN")
		OnMessage(0x101, "WM_KEYUP")
	
		; Call "HandleMessage" when script receives WM_SETCURSOR message 
		WM_SETCURSOR = 0x20 
		OnMessage(WM_SETCURSOR, "HandleMessage") 

		; Call "HandleMessage" when script receives WM_MOUSEMOVE message 
		WM_MOUSEMOVE = 0x200 
		OnMessage(WM_MOUSEMOVE, "HandleMessage")
	}
	;Activate desired Pane
	if(ShowPane != "")
	{
		Gui, 1:TreeView, SettingsTreeView
		id := 0
		while(id := TV_GetNext(id, "Full"))
		{
			TV_GetText(text, id)
			if(text = ShowPane)
			{
				TV_Modify(id, "Select")
				break
			}
		}
	}
	Return
}

;---------------------------------------------------------------------------------------------------------------
; Control Handlers
;---------------------------------------------------------------------------------------------------------------

SettingsTreeView:
SettingsTreeViewEvents()
return

SettingsTreeViewEvents()
{
	global SettingsTabList,SuppressTreeViewMessages,EventFilter
	static PreviousSelection
	if(SuppressTreeViewMessages)
		return
	if(!(A_GUIEvent = "" || A_GUIEvent = "S" || A_GUIEvent = "Normal"))
		return
	outputdebug tree view event
	selected := TV_GetSelection()
	TV_GetText(SelectedName, selected)
	parent := TV_GetParent(selected)
	if(parent)
		TV_GetText(ParentName, parent)
	Loop, Parse, SettingsTabList, |
	{
		StringReplace, stripped, A_LoopField, %A_Space% , , 1
		StringReplace, stripped, stripped, / , , 1
		stripped .= "Tab"
			outputdebug hide %stripped%
			GuiControl, 1:Hide, %stripped%
	}
	if(ParentName = "All Events" || SelectedName = "All Events")
	{
		outputdebug events tab
		GuiControl, 1:Show, EventsTab
		GuiControl, 1:Text, GGroupBox, %SelectedName%
		ControlSetText,,, ahk_id %EventFilter% ;This will trigger the refreshing of the events listview (EventFilterChange)
	}
	else
	{
		outputdebug regular tab
		GuiControl, 1:Hide, EventsTab
		Loop, Parse, SettingsTabList, |
		{
			StringReplace, stripped, A_LoopField, %A_Space% , , 1
			StringReplace, stripped, stripped, / , , 1
			stripped .= "Tab"
			If (SelectedName = A_LoopField)
			{
				outputdebug show %stripped%
				GuiControl, 1:Show, %stripped%
				GuiControl, 1:Text, GGroupBox, %A_LoopField%
				test:=stripped
				break
			}
		}
		if(SelectedName = "Accessor Keywords" && PreviousSelection != "Accessor Keywords")
			Gui +E0x10
		else if(PreviousSelection = "Accessor Keywords")
			Gui -E0x10
	}
	
	GuiControl, 1:MoveDraw, SettingsTreeView
	GuiControl, 1:Movedraw, GGroupbox
	if(test)
		GuiControl, 1:Movedraw, %test%
	else
		GuiControl, 1:Movedraw, EventsTab
	GuiControl, 1:MoveDraw, BtnOK
	GuiControl, 1:MoveDraw, BtnCancel
	GuiControl, 1:MoveDraw, BtnApply
	GuiControl, 1:MoveDraw, TutLabel
	GuiControl, 1:MoveDraw, Wait
	if(SelectedName != PreviousSelection)
		PreviousSelection := SelectedName
	return
}
EventFilterChange:
FillEventsList()
return
ShowAllEvents:
GuiControlGet,ShowAdvancedEvents
RecreateTreeView()
return
GUI_EventsList_SelectionChange:
GUI_EventsList_Update()
return
GUI_EventsList_Update()
{
	global
	local filter, count, i, checked, ListEvent
	WasCritical := A_IsCritical
	Critical
	ListEvent := Errorlevel
	outputdebug GUI_EventsList_Update() listview
	Gui, ListView, GUI_EventsList
	ControlGetText, filter,, ahk_id %EventFilter%
	count := LV_GetCount("Selected")
	if(A_GuiEvent="I" && InStr(ListEvent, "S", true))
	{	
		GuiControl, 1:enable, GUI_EventsList_Remove
		GuiControl, 1:enable, GUI_EventsList_Copy
		GuiControl, 1:enable, GUI_EventsList_Export
		if(count = 1)
			GuiControl, 1:enable, GUI_EventsList_Edit
		if(count > 1)
			GuiControl, 1:disable, GUI_EventsList_Edit
		if(count != 1)
			return
		i:=LV_GetNext("")
		LV_GetText(id,i,2)
		Event := Settings_Events.GetItemWithValue("ID", id)
		GuiControl, 1:, GUI_EventsDescription, % Event.Description
	}
	else if(A_GuiEvent="I" && InStr(ListEvent, "s", true))
	{
		if(count = 0)
		{
			GuiControl, 1:disable, GUI_EventsList_Edit
			GuiControl, 1:disable, GUI_EventsList_Remove
			GuiControl, 1:disable, GUI_EventsList_Copy
			GuiControl, 1:disable, GUI_EventsList_Export
		}
		else if(count = 1)
			GuiControl, 1:enable, GUI_EventsList_Edit
	}
	if(A_GuiEvent = "I" && InStr(ListEvent, "c")) ;Catch both check and uncheck
	{
		;Update enabled state from listview
		count := LV_GetCount()
		Loop % count
		{
			Checked := LV_GetNext(A_Index-1, "Checked") = A_Index ? 1 : 0
			LV_GetText(id,A_Index,2)
			Event := Settings_Events.GetItemWithValue("ID", id)
			if((!IsPortable && A_IsAdmin) || !Event.Trigger.Is(CExplorerButtonTrigger))
				Event.Enabled := Checked
		}
	}
	else if(A_GuiEvent="DoubleClick")
		GUI_EventsList_Edit()
	if(!WasCritical)
		Critical, Off
	Return
}
GUI_EventsList_Add:
GUI_AddEvent()
Return
GUI_AddEvent()
{
	global Settings_Events, GUI_EventsList
	Gui, ListView, GUI_EventsList
	Event := EventSystem_RegisterEvent(Settings_Events) ;Event is added to Settings_Events here
	LV_Modify(LV_GetNext(""), "-Select")
	LV_Add("Select Check", "", Event.ID, Event.Trigger.DisplayString(), Event.Name)	
	selected := TV_GetSelection()
	TV_GetText(Category,selected)
	if(Category = "All Events")
		Category := "Uncategorized"
	Event.Category := Category
	GUI_EventsList_Edit(1)
}
GUI_EventsList_Remove:
GUI_RemoveEvent()
Return
GUI_RemoveEvent()
{
	global Settings_Events, IsPortable
	outputdebug GUI_RemoveEvent() listview
	Gui, ListView, GUI_EventsList
	count := LV_GetCount()
	ListPos := 1
	Events := Array()
	Loop % count
	{
		if(LV_GetNext(A_Index-1) = A_Index)
		{
			LV_GetText(id,A_Index,2)			
			Events.Insert(id)
		}
	}
	Loop % Events.MaxIndex()
	{
		Event := Settings_Events.GetItemWithValue("ID", Events[A_Index])
		if((!IsPortable && A_IsAdmin) || !Event.Trigger.is(CExplorerButtonTrigger) && !Event.Trigger.Is(CContextMenuTrigger))
		{
			deleted += Settings_Events.Delete(Event, false) ;Settings_Events has special delete function
			Loop % LV_GetCount()
			{
				LV_GetText(id,A_Index,2)
				if(id = Event.ID)
				{
					ListPos := A_Index
					LV_Delete(A_Index)
				}
			}
			continue
		}
	}
	count :=  LV_GetCount()
	if(count)
	{
		ListPos := min(max(ListPos, 1), count)
		LV_Modify(ListPos, "Select")
	}
	outputdebug deleted %deleted%
	if(deleted) ;If a category was deleted
		RecreateTreeView()
}
GUI_EventsList_Edit:
GUI_EventsList_Edit()
return

GUI_EventsList_Edit(Add = 0)
{	
	global Settings_Events, EventFilter, IsPortable
	Critical Off
	outputdebug GUI_EventsList_Edit() listview
	Gui, ListView, GUI_EventsList
	if(LV_GetCount("Selected") != 1)
		return
	i:=LV_GetNext("")
	LV_GetText(id,i,2)
	pos := Settings_Events.FindKeyWithValue("ID", id)
	if((IsPortable || !A_IsAdmin) && Settings_Events[pos].Trigger.Is(CExplorerButtonTrigger))
	{
		Msgbox ExplorerButton trigger events may not be modified in portable or non-admin mode, as this might cause inconsistencies with the registry.
		return
	}
	Suspend, On
	event:=GUI_EditEvent(Settings_Events[pos].DeepCopy())
	Suspend, Off
	if(event && (IsPortable || !A_IsAdmin) && Event.Trigger.Type = "ExplorerButton") ;Explorer buttons may not be added in portable/non-admin mode
	{
		Msgbox ExplorerButton trigger events may not be modified in portable or non-admin mode, as this might cause inconsistencies with the registry.
		if(Add)
			GUI_RemoveEvent()
		return
	}
	if(event)
	{
		ControlSetText,,, ahk_id %EventFilter%
		Settings_Events[pos] := event ;overwrite edited event
		if(event.Category = "")
			event.Category := "Uncategorized"
		if(!Settings_Events.Categories.indexOf(event.Category))
			Settings_Events.Categories.Insert(event.Category)
		Settings_SetupEvents() ;Refresh listview		
	}
	else if(Add)
		GUI_RemoveEvent()
	Return
}
GUI_EventsList_Copy:
GUI_EventsList_Copy()
return

GUI_EventsList_Copy()
{
	global Settings_Events, IsPortable
	Gui, ListView, GUI_EventsList
	count := LV_GetCount()
	if(count=0)
		return
	Settings_Events_Clipboard := Array()
	Loop % count
	{
		if(LV_GetNext(A_Index-1) = A_Index)
		{
			LV_GetText(id,A_Index,2)			
			Event := Settings_Events.GetItemWithValue("ID", id)
			Copy := Event.DeepCopy()
			Copy.Remove("OfficialEvent") ;Make sure that pasted events don't patch existing events
			if((!IsPortable && A_IsAdmin) || Event.Trigger.Type != "ExplorerButton")
				Settings_Events_Clipboard.Insert(copy)
		}
	}
	WriteEventsFile(Settings_Events_Clipboard,A_Temp "/7plus/EventsClipboard.xml")	
	GuiControl, 1:enable, GUI_EventsList_Paste
}

GUI_EventsList_Paste:
GUI_EventsList_Paste()
return

GUI_EventsList_Paste()
{
	global Settings_Events, IsPortable
	Gui, ListView, GUI_EventsList
	if(FileExist(A_Temp "/7plus/EventsClipboard.xml"))
	{
		selected := TV_GetSelection()
		TV_GetText(Category,selected)
		if(Category = "All Events")
			Category := "Uncategorized"
		ReadEventsFile(Settings_Events, A_Temp "/7plus/EventsClipboard.xml", Category)
		Settings_SetupEvents()
	}
}
GUI_EventsList_Import:
GUI_EventsList_Import()
return
GUI_EventsList_Export:
GUI_EventsList_Export()
return
GUI_EventsList_Import()
{
	global Settings_Events
	FileSelectFile, file, 3, , Import Events file, Event files (*.xml)
	oldlen := Settings_Events.MaxIndex()
	if(file)
		ReadEventsFile(Settings_Events, file)
	Settings_SetupEvents()
	
	;Figure out if FTP events were added and notify the user to set the FTP profile assignments
	Loop % Settings_Events.MaxIndex() - oldlen
	{
		pos := A_Index + oldlen
		Loop % Settings_Events[pos].Actions.MaxIndex()
		{
			if(Settings_Events[pos].Actions[A_Index].Type = "Upload")
			{
				found := true
				break
			}
		}
		if(found)
			break
	}
	if(found)
		MsgBox, Note: Make sure to assign the FTP profiles of all imported FTP actions!
}
GUI_EventsList_Export()
{
	global Settings_Events, MajorVersion, MinorVersion
	
	
	;Uncomment the following lines to export all events separated by category to Events\Category.xml instead	
	Loop % Settings_Events.Categories.MaxIndex()
	{
		Category := Settings_Events.Categories[A_Index]
		Events := Array()
		Loop % Settings_Events.MaxIndex()
		{
			if(Settings_Events[A_Index].Category = Category)
				Events.Insert(Settings_Events[A_Index])
		}
		if(Events.MaxIndex() > 0)
			WriteEventsFile(Events, A_ScriptDir "\Events\" Category ".xml")
	}
	WriteEventsFile(Settings_Events, A_ScriptDir "\Events\All Events.xml")
	run % """" A_ScriptDir "\CreateEventPatch.ahk""" " """ A_ScriptDir "\Events\Old Versions\" MajorVersion "." (MinorVersion-1) "." 0 "\All Events.xml"" """ A_ScriptDir "\Events\All Events.xml"" 0" ;Create event patch, assumes that last minor version was incremented by one since last release
	return
	
	
	Gui, ListView, GUI_EventsList
	count := LV_GetCount("Selected")
	if(count > 0)
	{
		FileSelectFile, file, S19, , Export Events file, Event files (*.xml)
		if(file)
		{
			if(!strEndsWith(file, ".xml"))
				file .= ".xml"
			Events := Array()
			Loop % Settings_Events.MaxIndex()
			{
				if(LV_GetNext(A_Index - 1) = A_Index)
				{
					LV_GetText(id,A_Index,2)
					Event := Settings_Events.GetItemWithValue("ID", id)
					Events.Insert(Event)
					if(!FTP && Event.Actions.FindKeyWithValue("Type", "Upload"))
						FTP := true
				}
			}
			WriteEventsFile(Events, file)
			if(FTP)
				MsgBox, Note: FTP profiles won't be exported by this function. To save them, create a backup of FTPProfiles.xml. This file is only updated at program exit!
		}
	}	
}
GUI_EventsList_Help:
OpenWikiPage("EventsOverview")
Return
GUI_SaveEvents()
{
	global Events, Settings_Events
	;Disable all events first (without setting enabled to false, so triggers can decide what they want to do themselves)
	Loop % Events.MaxIndex()
		Events[A_Index].Trigger.Disable(Events[A_Index])	
	;Remove events that were deleted in settings window and refresh the settings copies to consider recent changes in the original events (such as timer state)
	pos := 1
	Settings_Events.IsSettings := true
	Loop % Events.MaxIndex()
	{
		if(!Settings_Events.GetItemWithValue("ID", Events[pos].id)) ;separate destroy routine instead of simple disable is needed for removed events because of hotkey/timer discrepancy
		{
			Events_Delete(Events, Events[pos],false)
			continue
		}
		Events[pos].Trigger.PrepareReplacement(Events[pos], Settings_Events.GetItemWithValue("ID", Events[pos].id))
		pos++
	}
	;Replace the original events with the copies
	Events := Settings_Events.DeepCopy()
	;Update enabled state
	Loop % Events.MaxIndex()
	{
		if(Events[A_Index].Enabled)
			Events[A_Index].Enable()
		else
			Events[A_Index].Disable()
	}
}
GUI_AccessorKeywordsTextChange:
GUI_AccessorKeywordsTextChange()
return
GUI_AccessorKeywordsTextChange()
{
	global GUI_AccessorKeywordsKey, GUI_AccessorKeywordsCommand, Settings_AccessorKeywords
	GuiControlGet, key,, GUI_AccessorKeywordsKey
	GuiControlGet, command,, GUI_AccessorKeywordsCommand
	Gui, ListView, GUI_AccessorKeywordsList
	i:=LV_GetNext("")
	if(!i)
		return
	LV_GetText(pos,i,1)
	LV_Modify(i, "Col2", key)
	LV_Modify(i, "Col3", command)
	Settings_AccessorKeywords[pos].key := key
	Settings_AccessorKeywords[pos].command := command
}
GUI_AccessorKeywords_Add:
GUI_Accessor_AddKeyword()
return
GUI_Accessor_AddKeyword()
{
	global Settings_AccessorKeywords
	Gui, ListView, GUI_AccessorKeywordsList
	Settings_AccessorKeywords.Insert(Object("Key", "Key", "Command", "Command"))
	LV_Add("Select", Settings_AccessorKeywords.MaxIndex(), "Key", "Command")
}
GUI_AccessorKeywords_Delete:
GUI_AccessorKeywords_Delete()
return
GUI_AccessorKeywords_Delete()
{
	global Settings_AccessorKeywords
	Gui, ListView, GUI_AccessorKeywordsList
	if(LV_GetCount("Selected") != 1)
		return
	len := LV_GetCount()
	i:=LV_GetNext("")
	LV_GetText(pos,i,1)
	Settings_AccessorKeywords.Delete(pos)
	LV_Delete(i)
	Loop % len
	{
		LV_GetText(pos,A_Index,1)
		if(pos>i)
			LV_Modify(A_Index, "Col1", pos - 1)
	}
}
GUI_AccessorKeywordsList_Events:
GUI_AccessorKeywordsList_Events()
return
GUI_AccessorKeywordsList_Events()
{
	global
	local count, ListEvent
	WasCritical := A_IsCritical
	Critical
	ListEvent := Errorlevel
	Gui, ListView, GUI_AccessorKeywordsList
	if(A_GuiEvent="I" && InStr(ListEvent, "S", true))
	{	
		Selected := LV_GetNext()		
		LV_GetText(id,Selected,1)
		GuiControl,,GUI_AccessorKeywordsKey, % Settings_AccessorKeywords[id].Key
		GuiControl,,GUI_AccessorKeywordsCommand, % Settings_AccessorKeywords[id].Command
		GuiControl, 1:enable, GUI_AccessorKeywordsKey
		GuiControl, 1:enable, GUI_AccessorKeywordsCommand
		GuiControl, 1:enable, GUI_AccessorKeywords_Delete
	}
	else if(A_GuiEvent="I" && InStr(ListEvent, "s", true))
	{		
		GuiControl, 1:disable, GUI_AccessorKeywordsKey
		GuiControl, 1:disable, GUI_AccessorKeywordsCommand
		GuiControl, 1:disable, GUI_AccessorKeywords_Delete
		GuiControl,,GUI_AccessorKeywordsKey,
		GuiControl,,GUI_AccessorKeywordsCommand,
	}
	if(!WasCritical)
		Critical, Off
	Return
}
return
GUI_SaveAccessorKeywords()
{
	global Accessor, Settings_AccessorKeywords
	pos := 1
	len := Settings_AccessorKeywords.MaxIndex()
	Loop % len
	{
		keywordentry := Settings_AccessorKeywords[A_Index]
		Loop % Settings_AccessorKeywords.MaxIndex()
		{
			if(pos != A_Index && Settings_AccessorKeywords[A_Index].Key = keywordentry.Key)
			{
				Settings_AccessorKeywords.Delete(pos)
				keywordentry := ""
				break
			}
		}
		if(IsObject(keywordentry))
			pos++
	}
	Accessor.Keywords := Settings_AccessorKeywords.DeepCopy()
}
GuiDropFiles:
DropAccessorFiles()
return
DropAccessorFiles()
{
	global Settings_AccessorKeywords
	Gui, ListView, GUI_AccessorKeywordsList
	Loop, parse, A_GuiEvent, `n
	{
		SplitPath, A_LoopField, ,,,name
		Settings_AccessorKeywords.Insert(Object("Key", name, "Command", A_LoopField))
		LV_Add("Select", Settings_AccessorKeywords.MaxIndex(), name, A_LoopField)
	}
	return
}
GUI_Accessor_Help:
OpenWikiPage("docsAccessor")
return
GUI_AccessorSettings:
ShowAccessorSettings()
return
ShowAccessorSettings()
{
	global Settings_AccessorPlugins, AccessorPlugins
	outputdebug ShowAccessorSettings() ListView
	Gui, ListView, GUI_AccessorPluginsList
	if(LV_GetCount("Selected") != 1)
		return
	i:=LV_GetNext("")
	LV_GetText(pos,i,2)
	if(!AccessorPlugins[pos].HasSettings)
		return
	outputdebug % "type 1 " Settings_AccessorPlugins[pos].type
	PluginSettings:=GUI_EditAccessorPlugin(Settings_AccessorPlugins[pos].DeepCopy())
	
	if(PluginSettings)
		Settings_AccessorPlugins[pos] := PluginSettings
}
GUI_AccessorPluginsList_Events:
GUI_AccessorPluginsList_Events()
return
GUI_AccessorPluginsList_Events()
{
	global
	local count, ListEvent
	WasCritical := A_IsCritical
	Critical
	ListEvent := Errorlevel
	outputdebug GUI_AccessorPluginsList_Events() listview
	Gui, ListView, GUI_AccessorPluginsList
	if(A_GuiEvent="I" && InStr(ListEvent, "S", true))
	{	
		Selected := LV_GetNext()		
		LV_GetText(id,Selected,2)
		if(AccessorPlugins[id].HasSettings)
			GuiControl, 1:enable, GUI_AccessorSettings
		else
			GuiControl, 1:disable, GUI_AccessorSettings
	}
	else if(A_GuiEvent="I" && InStr(ListEvent, "s", true))
		GuiControl, 1:Disable, GUI_AccessorSettings
	if(A_GuiEvent = "I" && InStr(ListEvent, "c")) ;Catch both check and uncheck
	{
		;Update enabled state from listview
		count := LV_GetCount()
		Loop % count
		{
			Checked := LV_GetNext(A_Index-1, "Checked") = A_Index ? 1 : 0
			LV_GetText(id,A_Index,2)
			Settings_AccessorPlugins[id].Enabled := Checked
			outputdebug % Settings_AccessorPlugins[id].Type "enabled " checked
		}
	}
	else if(A_GuiEvent="DoubleClick")
		ShowAccessorSettings()
	if(!WasCritical)
		Critical, Off
	Return
}
return
GUI_SaveAccessorPluginSettings()
{
	global AccessorPlugins, Settings_AccessorPlugins	
	outputdebug save accessor settings
	Loop % AccessorPlugins.MaxIndex()
	{
		Plugin := AccessorPlugins[A_Index]
		Settings_Plugin := Settings_AccessorPlugins[A_Index]
		Plugin.Enabled := Settings_Plugin.Enabled
		Plugin.Settings := Settings_Plugin.Settings.DeepCopy()
		outputdebug % Plugin.Type " save enabled " Plugin.Enabled
	}
}
txt:
GuiControlGet, enabled ,1: , Paste text as file
GuiControl, 1:enable%enabled%,TxtName
Return

img:
GuiControlGet, enabled ,1: , Paste image as file
GuiControl, 1:enable%enabled%,ImgName
Return

FastFolders:
GuiControl, 1:enable%Vista7%, HKFolderBand
GuiControl, 1:enable%Vista7%, HKCleanFolderBand
Return

UseTabs:
GuiControlGet, enabled ,1: , UseTabs
GuiControl, 1:enable%enabled%, NewTabPosition
GuiControl, 1:enable%enabled%, TabStartupPath
GuiControl, 1:enable%enabled%, ActivateTab
GuiControl, 1:enable%enabled%, TabWindowClose
GuiControl, 1:enable%enabled%, OnTabClose
GuiControl, 1:enable%enabled%, ShowSingleTab
GuiControl, 1:enable%enabled%, TabStartupPathBrowse
GuiControl, 1:enable%enabled%, TabLabel1
GuiControl, 1:enable%enabled%, TabLabel2
GuiControl, 1:enable%enabled%, TabLabel3
Return

TabStartupPathBrowse:
Gui 1:+OwnDialogs
path:=COMObjCreate("Shell.Application").BrowseForFolder(0, "Enter Path to add as button", 0).Self.Path
if(path!="")
	GuiControl, , 1:TabStartupPath,%path%
return

FTPProfilesDropDownList:
FTPProfilesDropDownList()
return
FTPProfilesDropDownList()
{
	global Settings_FTPProfiles	
	SavePreviousFTPProfile()
	GuiControlGet, selection,, FTPProfilesDropDownList
	outputdebug % "Old ID: " Settings_FTPProfiles.CurrentID
	ShowCurrentFTPProfile()
	Settings_FTPProfiles.CurrentId := SubStr(selection, 1, Instr(selection, ": ") - 1)
	outputdebug % "Current ID: " Settings_FTPProfiles.CurrentID
}
FTPProfiles_Hostname:
FTPProfiles_Hostname()
return
FTPProfiles_Hostname()
{
	global Settings_FTPProfiles
	GuiControlGet, hostname,, FTPProfiles_Hostname
	GuiControlGet, selection,, FTPProfilesDropDownList
	if(SubStr(selection, 1, Instr(selection, ": ") - 1) != Settings_FTPProfiles.CurrentID)
		return
	Loop % Settings_FTPProfiles.MaxIndex()
		Profiles .= "|" A_Index ": " (A_Index = Settings_FTPProfiles.CurrentID ? hostname : Settings_FTPProfiles[A_Index].Hostname) (A_Index = Settings_FTPProfiles.CurrentID ? "|" : "")
	if(strEndsWith(Profiles, "|"))
		Profiles .= "|"
	GuiControl, 1:,FTPProfilesDropDownList, %Profiles%
}
SavePreviousFTPProfile()
{
	global Settings_FTPProfiles, FTPProfiles_Hostname, FTPProfiles_Port, FTPProfiles_User, FTPProfiles_Password, FTPProfiles_URL
	CurrentID := Settings_FTPProfiles.CurrentID
	if(CurrentID)
	{
		CurrentProfile := Settings_FTPProfiles[CurrentID]
		GuiControlGet, Hostname, , FTPProfiles_Hostname
		CurrentProfile.Hostname := strTrimRight(Hostname, "/")
		GuiControlGet, Port, , FTPProfiles_Port
		CurrentProfile.Port := Port
		GuiControlGet, User, , FTPProfiles_User	
		CurrentProfile.User := User
		GuiControlGet, Password, , FTPProfiles_Password
		CurrentProfile.Password := Encrypt(Password)
		GuiControlGet, URL, , FTPProfiles_URL
		CurrentProfile.URL := strTrimRight(URL, "/")
	}
}
ShowCurrentFTPProfile()
{
	global Settings_FTPProfiles, FTPProfilesDropDownList, FTPProfiles_Hostname, FTPProfiles_Port, FTPProfiles_User, FTPProfiles_Password, FTPProfiles_URL
	GuiControlGet, selection,, FTPProfilesDropDownList
	id := SubStr(selection, 1, Instr(selection, ": ") - 1)
	;If id is invalid (because there are no profiles), fields will be cleared
	
	Password := Decrypt(Settings_FTPProfiles[id].Password)
	GuiControl,,FTPProfiles_Hostname, % Settings_FTPProfiles[id].Hostname
	GuiControl,,FTPProfiles_Port, % Settings_FTPProfiles[id].Port
	GuiControl,,FTPProfiles_User, % Settings_FTPProfiles[id].User
	GuiControl,,FTPProfiles_Password, %Password%
	GuiControl,,FTPProfiles_URL, % Settings_FTPProfiles[id].URL
}
FTPProfiles_Add:
FTPProfiles_Add()
return
FTPProfiles_Add()
{
	global Settings_FTPProfiles, FTPProfilesDropDownList
	Settings_FTPProfiles.Insert(Object("Hostname", "Hostname.com", "Port", 21, "User", "SomeUser", "Password", "", "URL", "http://somehost.com"))
	len := Settings_FTPProfiles.MaxIndex()
	string := len ": " Settings_FTPProfiles[Settings_FTPProfiles.MaxIndex()].Hostname
	GuiControl, , FTPProfilesDropDownList, %string%
	GuiControl, Choose,FTPProfilesDropDownList, |%len% ;|->trigger g-label
	GuiControl, enable, FTPProfiles_Hostname
	GuiControl, enable, FTPProfiles_Port
	GuiControl, enable, FTPProfiles_User
	GuiControl, enable, FTPProfiles_Password
	GuiControl, enable, FTPProfiles_URL
}
FTPProfiles_Delete:
FTPProfiles_Delete()
return
FTPProfiles_Delete()
{
	global Settings_FTPProfiles, FTPProfilesDropDownList
	GuiControlGet, selection,, FTPProfilesDropDownList
	id := SubStr(selection, 1, Instr(selection, ": ") - 1)
	Settings_FTPProfiles.Delete(id)
	Settings_FTPProfiles.CurrentID := 0 ;Don't save current values
	Loop % Settings_FTPProfiles.MaxIndex()
		Profiles .= "|" A_Index ": " Settings_FTPProfiles[A_Index].Hostname (A_Index = 1 ? "|" : "")
	outputdebug new list: %Profiles%
	GuiControl, 1:,FTPProfilesDropDownList, %Profiles%
	GuiControl, 1:Choose, FTPProfilesDropDownList, |1 ;| -> trigger g-label
	GuiControlGet, selection,, FTPProfilesDropDownList
	if(!selection) ;No entries anymore, disable controls
	{
		GuiControl, disable, FTPProfiles_Hostname
		GuiControl, disable, FTPProfiles_Port
		GuiControl, disable, FTPProfiles_User
		GuiControl, disable, FTPProfiles_Password
		GuiControl, disable, FTPProfiles_URL
	}
	MsgBox, Make sure to update any FTP event profile assignments that pointed to the deleted profile!
}

GUI_HotstringsTextChange:
GUI_HotstringsTextChange()
return
GUI_HotstringsTextChange()
{
	global GUI_HotstringsKey, GUI_HotstringsValue, Settings_Hotstrings
	GuiControlGet, key,, GUI_HotstringsKey
	GuiControlGet, value,, GUI_HotstringsValue
	Gui, ListView, GUI_HotstringsList
	i:=LV_GetNext("")
	if(!i)
		return
	LV_GetText(pos,i,1)
	LV_Modify(i, "Col2", key)
	LV_Modify(i, "Col3", value)
	Settings_Hotstrings[pos].key := key
	Settings_Hotstrings[pos].value := value
}
GUI_Hotstrings_Add:
GUI_Hotstrings_Add()
return
GUI_Hotstrings_Add()
{
	global Settings_Hotstrings
	Gui, ListView, GUI_HotstringsList
	Settings_Hotstrings.Insert(Object("Key", "key", "value", "value"))
	LV_Add("Select", Settings_Hotstrings.MaxIndex(), "key", "value")
}
GUI_Hotstrings_Delete:
GUI_Hotstrings_Delete()
return
GUI_Hotstrings_Delete()
{
	global Settings_Hotstrings
	Gui, ListView, GUI_HotstringsList
	if(LV_GetCount("Selected") != 1)
		return
	len := LV_GetCount()
	i:=LV_GetNext("")
	LV_GetText(pos,i,1)
	Settings_Hotstrings.Delete(pos)
	LV_Delete(i)
	Loop % len
	{
		LV_GetText(pos,A_Index,1)
		if(pos>i)
			LV_Modify(A_Index, "Col1", pos - 1)
	}
}
GUI_Hotstrings_RegexHelp:
run http://www.autohotkey.com/docs/misc/RegEx-QuickRef.htm
return

GUI_HotstringsList_Events:
GUI_HotstringsList_Events()
return
GUI_HotstringsList_Events()
{
	global
	local count, ListEvent
	WasCritical := A_IsCritical
	Critical
	ListEvent := Errorlevel
	Gui, ListView, GUI_HotstringsList
	if(A_GuiEvent="I" && InStr(ListEvent, "S", true))
	{	
		Selected := LV_GetNext()		
		LV_GetText(id,Selected,1)
		GuiControl,,GUI_HotstringsKey, % Settings_Hotstrings[id].Key
		GuiControl,,GUI_HotstringsValue, % Settings_Hotstrings[id].Value
		GuiControl, 1:enable, GUI_HotstringsKey
		GuiControl, 1:enable, GUI_HotstringsValue
		GuiControl, 1:enable, GUI_Hotstrings_Delete
	}
	else if(A_GuiEvent="I" && InStr(ListEvent, "s", true))
	{		
		GuiControl, 1:disable, GUI_HotstringsKey
		GuiControl, 1:disable, GUI_HotstringsValue
		GuiControl, 1:disable, GUI_Hotstrings_Delete
		GuiControl,,GUI_HotstringsKey,
		GuiControl,,GUI_HotstringsValue
	}
	if(!WasCritical)
		Critical, Off
	Return
}
return
GUI_SaveHotstrings()
{
	global Hotstrings, Settings_Hotstrings
	pos := 1
	len := Settings_Hotstrings.MaxIndex()
	Loop % len
	{
		Hotstring := Settings_Hotstrings[A_Index]
		;Remove duplicates and empty keys
		Loop % Settings_Hotstrings.MaxIndex()
		{
			if(pos != A_Index && Settings_Hotstrings[A_Index].Key = Hotstring.Key)
			{
				Settings_Hotstrings.Delete(pos)
				Hotstring := ""
				break
			}
			if(Settings_Hotstrings[A_Index].Key = "")
			{
				Settings_Hotstrings.Delete(pos)
				Hotstring := ""
				break
			}
		}
		if(IsObject(Hotstring))
			pos++
	}
	len := Hotstrings.MaxIndex()
	Loop % len
		RemoveHotstring(Hotstrings[1].key)
	Loop % Settings_Hotstrings.MaxIndex()
		AddHotstring(Settings_Hotstrings[A_Index].key, Settings_Hotstrings[A_Index].value)
}

SlideWindow:
GuiControlGet, enabled,1: , HKSlideWindows
GuiControl, 1:enable%enabled%, SlideWinHide
return

WM_KEYDOWN(wParam, lParam)
{
	Critical
	global EventFilter, SettingsActive
	outputdebug wm_keydown()
	if(A_GUI = 1 && SettingsActive)
	{
		outputdebug settings
		/*
		if(A_GuiControl = "CustomHotkeysList" && wParam = 0x2E) ;Delete key pressed on CustomHotkeysList
		{
			Gui, ListView, CustomHotkeysList
			i:=LV_GetNext("")
			LV_Delete(i)
		}
		*/
		if(A_GuiControl = "GUI_EventsList" && wParam = 0x2E) ;Delete key pressed on events list
			GUI_RemoveEvent()
		else if(A_GuiControl = "GUI_EventsList")
		{
			send := true
			if(wParam = 17 || (wParam > 32 && wParam < 41)) ;CTRL, arrow keys, home, end, page up/down
				send := false
			outputdebug wParam %wParam% lParam %lParam%
			if(GetKeyState("Control", "P")) ;Don't send when CTRL is down
			{
				send := false
				outputdebug ctrl
				if(GetKeyState("A", "P"))
				{
					outputdebug ctrl+a
					Loop % LV_GetCount()
						LV_Modify(A_Index, "Select")
				}
			}
			
			if(send)
			{
				outputdebug send keydown %wparam% to %EventFilter%
				PostMessage, 0x100, %wParam%, %lParam%,,ahk_id %EventFilter%
				Critical, Off
				return true
			}
		}
		else if(A_GuiControl = "GUI_AccessorKeywordsList" && wParam = 0x2E)
			GUI_AccessorKeywords_Delete()
		else if(A_GuiControl = "GUI_HotstringsList" && wParam = 0x2E)
			GUI_Hotstrings_Delete()
	}
	if(A_GUI = 4 && (A_GuiControl = "EditEventConditions" || A_GuiControl = "EditEventActions") && wParam = 0x2E)
		GUI_EditEvent("","RemoveSubEvent")
}
WM_KEYUP(wParam, lParam)
{
	global EventFilter
	outputdebug wm_keyup()
	if(A_GUI = 1 && A_GuiControl = "GUI_EventsList")
	{
		if(wParam = 17 || (wParam > 32 && wParam < 41)) ;CTRL, arrow keys, home, end, page up/down
			return false
		if(GetKeyState("Control", "P")) ;Don't send when CTRL is down
			return false

		outputdebug send keyup %wparam% to %EventFilter%
		PostMessage, 0x101, %wParam%, %lParam%,,ahk_id %EventFilter%
		return true
	}
}
WM_COMMAND(wParam, lParam)
{
	outputdebug WM_COMMAND wParam %wParam% lParam %lParam%
}

#if WinActive("7plus Settings ahk_class AutohotkeyGUI") && IsControlActive("SysListView321")
^a::
GUI, ListView, GUI_EventsList
Loop % LV_GetCount()
	LV_Modify(A_Index, "Select")
return
#if
;---------------------------------------------------------------------------------------------------------------
; Help Links
;---------------------------------------------------------------------------------------------------------------
hPasteAsFile:
run http://www.youtube.com/watch?v=yOJ8evyuVhY
return
hNavigation:
run http://www.youtube.com/watch?v=RZOdgDl2ujU
return
hExplorer1dot1:
run http://www.youtube.com/watch?v=xKmWOEbMI1Q
return
hScrollUnderMouse:
run http://www.youtube.com/watch?v=qJ_u4C3EuhU
return
hSelectFirstFile:
run http://www.youtube.com/watch?v=Bih7HEtpk0A
return
hShowSpaceAndSize:
run http://www.youtube.com/watch?v=-fnOBf3Ggoc
return
hApplyOperation:
run http://www.youtube.com/watch?v=flBnx2NETlc
return
hFastFolders1:
run http://www.youtube.com/watch?v=dTIGxue6WCY
return
hFastFolders2:
run http://www.youtube.com/watch?v=cC6cnG87j2M
return
hTabs:
msgbox Not recorded yet
return
hTaskbar:
run http://www.youtube.com/watch?v=v__ZiHFt7NE
return
hWindow:
run http://www.youtube.com/watch?v=JJ-kqjRY910
return
hCapslock:
run http://www.youtube.com/watch?v=im088NYiSvw
return
hSlideWindow:
run http://www.youtube.com/watch?v=e0yLqr8mjsg
return
hWindow1dot1:
run http://www.youtube.com/watch?v=1vutsoA3j7Y
return
hImproveConsole:
run http://www.youtube.com/watch?v=irMu69t3kEg
return
hJoyControl:
run http://www.youtube.com/watch?v=MZiK7E98hOU
return
hWordDelete:
msgbox Not recorded yet
return
GPL:
run http://www.gnu.org/licenses/gpl.html
return
Mail:
run mailto://fragman@gmail.com
return
Twitter:
run http://www.twitter.com/7_plus
return
Ahk:
run http://www.autohotkey.com
return
Projectpage:
run http://code.google.com/p/7plus/
return
Bugtracker:
run http://code.google.com/p/7plus/issues/list
return
7zip:
run http://www.7-zip.org
Return
LGPL:
run http://www.gnu.org/licenses/lgpl.html
return
Donate:
run https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=CCDPER7Z2CHZW
return

;---------------------------------------------------------------------------------------------------------------
; OK/Cancel/Close/Apply
;---------------------------------------------------------------------------------------------------------------
OK:
ApplySettings(1)
return

Apply:
ApplySettings(0)
return

ApplySettings(Close = 0)
{
	global
	local active, enabled, pastename, flip, path, temp
	;First process variables which require comparison with previous values

	;Store Fast Folders settings and make everything consistent by backing up and restoring reg keys
	GuiControl, 1:Show, Wait
	GuiControl, 1:MoveDraw, Wait
	if(!IsPortable || !A_IsAdmin)
	{
		GuiControlGet, active ,1: , HKFolderBand
		if(active && !HKFolderBand)
			PrepareFolderBand()
		else if(HKFolderBand && !active)
			RestoreFolderBand()

		GuiControlGet, active ,1: , HKCleanFolderBand
		if(active && !HKCleanFolderBand)
			BackupAndRemoveFolderBandButtons()
		else if(HKCleanFolderBand && !active)
			RestoreFolderBandButtons()
				
		GuiControlGet, active ,1: , HKPlacesBar
		if(active && !HKPlacesBar)
			BackupPlacesBar()
		else if(HKPlacesBar && !active)
			RestorePlacesBar()
	}
	
	Settings_WindowsSettings_Submit(1)
	;Store variables which can be stored directly
	Gui 1:Submit, NoHide
	Language.CurrentLanguage := Language.Languages.GetItemWithValue("FullName", ControlGetText("", "ahk_id " hLanguage))
	Action := Settings_WindowsSettings_Submit(0)
	;SaveEvents relies on SettingsActive to be false so that enabling/dissabling the events doesn't refresh them in settings window
	SettingsActive := false
	GUI_SaveEvents()
	SettingsActive := true
	GUI_SaveAccessorKeywords()
	GUI_SaveAccessorPluginSettings()
	GUI_SaveHotstrings()
	if(JoyControl)
		JoystickStart()
	else
		JoystickStop()
	;Store paste text as file filename
	GuiControlGet, enabled ,1: , Paste text as file
	GuiControlGet, pastename ,1: , TxtName
	if(enabled)
	{
		TxtName:=pastename
		temp_txt:=A_Temp . "\" . TxtName
	}
	else
	{
		TxtName:=""
		temp_txt:=""
	}

	;Store paste image as file filename
	GuiControlGet, enabled ,1: , Paste image as file
	GuiControlGet, pastename ,1: , ImgName
	if(enabled)
	{
		ImgName:=pastename
		temp_img:=A_Temp . "\" . ImgName
	}
	else
	{
		ImgName:=""
		temp_img:=""
	}
	SavePreviousFTPProfile()
	FTPProfiles := Array()
	Loop % Settings_FTPProfiles.MaxIndex()
		FTPProfiles.Insert(Settings_FTPProfiles[A_Index])

	if(!(ImageQuality > 0 && ImageQuality <= 100))
		ImageQuality := 95
	if(!IsPortable)
	{
		;Store Autorun setting
		if(Autorun)
			EnableAutorun()
		else
			DisableAutorun()
	}
	if(HideTrayIcon)
	{
		MsgBox You have chosen to hide the tray icon. This means that you will only be able to access the settings dialog by pressing WIN + H (Default settings). Also, the program can only be ended by using the task manager then.
		Menu, Tray, NoIcon
	}
	else
		Menu, Tray, Icon
	Settings.Save()
	if(Action & 1)
	{
		MsgBox, 4, Restart Windows, You need to restart windows to apply a setting. Do you want to do this now?
		IfMsgBox Yes
			Shutdown, 2
	}
	else if(Action & 2)
	{
		MsgBox, 4, Restart Explorer, You need to restart explorer to apply a setting. Do you want to do this now?
		IfMsgBox Yes
		{
			Runwait, taskkill /im explorer.exe /f, , Hide
			Run, %a_windir%\explorer.exe
		}
	}
	GuiControl, 1:Hide, Wait
	if(Close)
		GoSub Cancel
	Return
}

Settings_WindowsSettings_Submit(PreSubmit)
{
	global
	static  PreRemoveWMP, PreHKActivateBehavior, PreDisableUAC, PreRemoveLibraries
	if(PreSubmit)
	{
		PreRemoveWMP := RemoveWMP
		PreHKActivateBehavior := HKActivateBehavior
		PreDisableUAC := DisableUAC
		PreRemoveLibraries := RemoveLibraries
	}
	else
	{
		RequiredAction := 0
		RequiredAction |= WindowsSettings_ShowAllTray(ShowAllTray)
		RequiredAction |= WindowsSettings_RemoveUserDir(RemoveUserDir)
		if(RemoveWMP != PreRemoveWMP)
			RequiredAction |= WindowsSettings_RemoveWMP(RemoveWMP)
		RequiredAction |= WindowsSettings_RemoveOpenWith(RemoveOpenWith)
		RequiredAction |= WindowsSettings_RemoveCrashReporting(RemoveCrashReporting)
		RequiredAction |= WindowsSettings_ShowExtensions(ShowExtensions)
		RequiredAction |= WindowsSettings_ShowHiddenFiles(ShowHiddenFiles)
		RequiredAction |= WindowsSettings_ShowSystemFiles(ShowSystemFiles)
		if(DisableUAC != PreDisableUAC)
				RequiredAction |= WindowsSettings_DisableUAC(DisableUAC)
		RequiredAction |= WindowsSettings_ClassicView(ClassicView)
		RequiredAction |= WindowsSettings_RemoveLibraries(RemoveLibraries)
		RequiredAction |= WindowsSettings_ActivateBehavior(HKActivateBehavior)
		RequiredAction |= WindowsSettings_HoverTime(HoverTime)
	}
}


GuiEscape:
Cancel:
GuiClose:
SettingsActive:=False
Settings_FTPProfiles := ""
Settings_Events := ""
Settings_AccessorKeywords := ""
Settings_AccessorPlugins := ""
Gui 1:Cancel
Return

;Link hand cursor handling
HandleMessage(p_w, p_l, p_m, p_hw) 
{ 
  global   WM_SETCURSOR, WM_MOUSEMOVE, 
  static   URL_hover, h_cursor_hand, h_old_cursor, CtrlIsURL, LastCtrl 
  If (p_m = WM_SETCURSOR) 
    { 
      If URL_hover 
        Return, true 
    } 
  Else If (p_m = WM_MOUSEMOVE) 
    { 
      ; Mouse cursor hovers URL text control 
      StringLeft, CtrlIsURL, A_GuiControl, 3 
      If (CtrlIsURL = "URL") 
        { 
          If URL_hover= 
            { 
              Gui, 1:Font, cBlue underline 
              GuiControl, 1:Font, %A_GuiControl% 
              LastCtrl = %A_GuiControl% 
              
              h_cursor_hand := DllCall("LoadCursor", "Ptr", 0, "uint", 32649, "Ptr") 
              
              URL_hover := true 
            }
            h_old_cursor := DllCall("SetCursor", "Ptr", h_cursor_hand, "Ptr") 
        } 
      ; Mouse cursor doesn't hover URL text control 
      Else 
        { 
          If URL_hover 
            { 
              Gui, 1:Font, norm cBlue 
              GuiControl, 1:Font, %LastCtrl% 
              
              DllCall("SetCursor", "Ptr", h_old_cursor) 
              
              URL_hover= 
            } 
        } 
    } 
}
