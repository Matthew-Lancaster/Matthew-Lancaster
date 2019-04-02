#include %A_ScriptDir%\Lib\FTP.ahk
Class CFTPUploadAction Extends CAction
{
	static Type := RegisterType(CFTPUploadAction, "Upload to FTP")
	static Category := RegisterCategory(CFTPUploadAction, "File")
	static SourceFiles := "${SelNM}" ;All upload actions need to have SourceFiles property (used in ImageConverter)
	static TargetFolder := ""
	static TargetFile := ""
	static Silent := 0
	static Clipboard := 1
	static FTPProfile := 1
	Startup()
	{
		if(FTPProfiles := this.ReadFTPProfiles())
			this.FTPProfiles := FTPProfiles
		else
			this.FTPProfiles := Array({Hostname : "Hostname.com", Password : "", Port : 21, URL : "http://somehost.com", User : "SomeUser"})
	}
	OnExit()
	{
		this.WriteFTPProfiles()
		this.Remove("FTPProfiles")
	}
	ReadFTPProfiles()
	{
		FTPProfiles := Array()
		FileRead, xml, % Settings.ConfigPath "\FTPProfiles.xml"
		if(!xml)
			return
		XMLObject := XML_Read(xml)
		
		;Convert to Array
		if(!XMLObject.List.MaxIndex())
			XMLObject.List := IsObject(XMLObject.List) ? Array(XMLObject.List) : Array()
		Loop % XMLObject.List.MaxIndex()
		{
			ListEntry := XMLObject.List[A_Index]
			FTPProfiles.Insert(Object("Hostname", ListEntry.Hostname, "Port", ListEntry.Port, "User", ListEntry.User, "Password", ListEntry.Password, "URL", ListEntry.URL, "NumberOfFTPSubDirs", ListEntry.NumberOfFTPSubDirs))
		}
		return FTPProfiles
	}
	WriteFTPProfiles()
	{
		ConfigPath := Settings.ConfigPath
		SplitPath, ConfigPath,,path
		path .= "\FTPProfiles.xml"
		FileDelete, %ConfigPath%\FTPProfiles.xml
		
		XMLObject := Object("List",Array())
		Loop % this.FTPProfiles.MaxIndex()
		{
			ListEntry := this.FTPProfiles[A_Index]
			XMLObject.List.Insert(Object("Hostname", ListEntry.Hostname, "Port", ListEntry.Port, "User", ListEntry.User, "Password", ListEntry.Password, "URL", ListEntry.URL, "NumberOfFTPSubDirs", ListEntry.NumberOfFTPSubDirs))
		}
		XML_Save(XMLObject, ConfigPath "\FTPProfiles.xml")
	}
	Execute(Event)
	{
		global FTP
		SourceFiles := ToArray(Event.ExpandPlaceholders(this.SourceFiles))
		TargetFolder := Event.ExpandPlaceholders(this.TargetFolder)
		TargetFile := Event.ExpandPlaceholders(this.TargetFile)
		files := ToArray(SourceFiles)
		;Process target filenames
		targets := Array()
		for index, File in SourceFiles
		{
			Splitpath, file, filename, , fileextension, filenamenoextension
			SplitPath, TargetFile, , , targetfileextension, targetfilenamenoextension
			if(targetfilenamenoextension && targetfileextension)
				targets.Insert(targetfilenamenoextension "." targetfileextension)
			else if(targetfilenamenoextension)
				targets.Insert(targetfilenamenoextension "." fileextension)
			else if(targetfileextension)
				targets.Insert(filenamenoextension "." targetfileextension)
			else
				targets.Insert(filename)
			file1 := targets[index]
			SplitPath, file1, ,,CheckExtension, CheckFilenameNoExtension
			number := 1
			Loop % index - 1 ;add (Number) to avoid duplicate filenames
				if(targets[index] = CheckFilenameNoExtension (number > 1 ? " (" number ")" : "") "." CheckExtension)
					number++
			targets[index] := CheckFilenameNoExtension (number > 1 ? " (" number ")" : "") "." CheckExtension
		}
		if(files.MaxIndex() > 0)
		{
			this.GetFTPVariables(this.FTPProfile, Hostname, Port, User, Password, URL, NumberOfFTPSubDirs)
			if(!Hostname || Hostname = "Hostname.com")
			{
				Notify("FTP profile not set", "The FTP profile was not created yet or is invalid. Click here to enter a valid FTP login.", "5", "GC=555555 AC=FTP_Notify_Error TC=White MC=White",NotifyIcons.Error)
				return 0
			}
			decrypted:=Decrypt(Password)
			cliptext := ""
			result := 1
			; connect to FTP server 
			FTP := FTP_Init()
			FTP.Port := Port
			FTP.Hostname := Hostname
			if !FTP.Open(Hostname, User, decrypted) 
			{ 
				if(!this.Silent)
					Notify("Connection Error", "Couldn't connect to " Hostname ". Correct host/username/password?", "5", "GC=555555 AC=FTP_Notify_Error TC=White MC=White",NotifyIcons.Error)
				result := 0
			}
			else
			{
				;go into target directory, optionally creating directories along the way.
				if(TargetFolder != "" && ftp.SetCurrentDirectory(TargetFolder) != true)
				{
					Loop, Parse, TargetFolder, /
					{
						;Skip current level if it exists
						if(ftp.SetCurrentDirectory(A_LoopField))
							continue
						;Try to create current level
						if(ftp.CreateDirectory(A_LoopField) != true)
						{
							if(!this.Silent)
								Notify("FTP Error", "Couldn't create target directory " A_LoopField ". Check permissions!", "5", "GC=555555 AC=FTP_Notify_Error TC=White MC=White",NotifyIcons.Error)
							result := 0
							break
						}
						;Try to go into newly created directory
						if(ftp.SetCurrentDirectory(A_LoopField) != true)
						{
							if(!this.Silent)
								Notify("FTP Error", "Couldn't switch to created target directory" A_LoopField ". Check permissions!", "5", "GC=555555 AC=FTP_Notify_Error TC=White MC=White",NotifyIcons.Error)
							result := 0
							break
						}
					}
				}
				if(result != 0)
				{
					;The url is sometimes mapped differently on FTP vs. Web.
					;FTP might have more directories while the webserver only mirrors a part of the directory structure.
					;The code below allows skipping some directories
					URLTargetFolder := TargetFolder
					Loop % NumberOfFTPSubDirs
					{
						if(pos := InStr(URLTargetFolder, "/"))
							URLTargetFolder := SubStr(URLTargetFolder, pos + 1)
						else
						{
							URLTargetFolder := ""
							break
						}
					}
					success := 0
					FTP.NumFiles := files.MaxIndex()
					Loop % files.MaxIndex()
					{
						FullPath := files[A_Index]
						result := FTP.InternetWriteFile(FullPath,  targets[A_Index], "Action_FTPUpload_Progress")
						if(!success && result)
							success := result
						if(result=0 && !this.Silent)
							Notify("Couldn't upload file", "Couldn't upload "  targets[A_Index] " properly. Make sure you have write rights and the path exists", "5", "GC=555555 AC=FTP_Notify_Error TC=White MC=White",NotifyIcons.Error)
						else if(result != 0 && URL && this.Clipboard)
							cliptext .= (A_Index = 1 ? "" : "`r`n") URL "/" URLTargetFolder (URLTargetFolder ? "/" : "") StringReplace(targets[A_Index], " ", "%20", 1)
					}
					FTP.Close()
					if(URL && this.Clipboard && cliptext)
						clipboard:=cliptext
					if(!this.Silent && success)
					{
						Notify("","",0, "Wait",FTP.NotifyID)
						Notify("Transfer finished", "File uploaded", 2, "GC=555555 TC=White MC=White",NotifyIcons.Success)
						SoundBeep
					}
					result := 1
				}
			}
			return result
		}
	}
	DisplayString()
	{
		this.GetFTPVariables(this.FTPProfile, Hostname, "", "", "", "")
		return "Upload " this.SourceFiles " to " Hostname (this.TargetFolder ? "/" : "") this.TargetFolder
	}

	GuiShow(GUI, GoToLabel = "")
	{
		static sGUI
		if(GoToLabel = "")
		{
			sGUI := GUI
			this.AddControl(GUI, "Text", "Desc", "This action can upload files to FTP servers.")
			this.AddControl(GUI, "Edit", "SourceFiles", "", "", "Source files:", "Browse", "Action_Upload_Browse", "Placeholders", "Action_Upload_Placeholders_SourceFiles")
			for index, Profile in SettingsWindow.FTPProfiles
				Profiles .= "|" index ": " Profile.Hostname
			this.AddControl(GUI, "DropDownList", "FTPProfile", Profiles, "", "FTP profile:","","","","","FTP profiles are created on their specific sub page in the settings window.")
			this.AddControl(GUI, "Edit", "TargetFolder", "", "", "Target folder:", "Placeholders", "Action_Upload_Placeholders_TargetFolder")
			this.AddControl(GUI, "Edit", "TargetFile", "", "", "Target files:", "Placeholders", "Action_Upload_Placeholders_TargetFile", "", "", "Filename(+ optionally .extension) / or empty for original filename")
			this.AddControl(GUI, "Checkbox", "Silent", "Silent", "", "")
			this.AddControl(GUI, "Checkbox", "Clipboard", "Copy links to clipboard", "", "")
		}
		else if(GoToLabel = "Placeholders_SourceFiles")
			ShowPlaceholderMenu(sGUI, "SourceFiles")
		else if(GoToLabel = "Browse")
			this.SelectFile(sGUI, "SourceFiles")
		else if(GoToLabel = "Placeholders_TargetFolder")
			ShowPlaceholderMenu(sGUI, "TargetFolder")
		else if(GoToLabel = "Placeholders_TargetFile")
			ShowPlaceholderMenu(sGUI, "TargetFile")
	}
	
	GetFTPVariables(id, ByRef Hostname, ByRef Port, ByRef User, ByRef Password, ByRef URL, ByRef NumberOfFTPSubDirs)
	{
		Hostname := this.FTPProfiles[id].Hostname
		Port := this.FTPProfiles[id].Port
		User := this.FTPProfiles[id].User
		Password := this.FTPProfiles[id].Password
		URL := this.FTPProfiles[id].URL
		NumberOfFTPSubDirs := this.FTPProfiles[id].NumberOfFTPSubDirs
	}
}
Action_FTPUpload_Progress()
{
	global FTP
	my := FTP.File
	done := my.BytesTransfered
	total := my.BytesTotal
	if(!FTP.init)
	{
		FTP.NotifyID := Notify("Uploading " FTP.NumFiles " file" (FTP.NumFiles = 1 ? "":"s") " to " FTP.Hostname,my.RemoteName " - " FormatFileSize(done) " / " FormatFileSize(total),"","PG=100 GC=555555 TC=White MC=White",NotifyIcons.Internet)
		FTP.init := 1
		return 1
	}
	Notify("","",done/total*100, "Progress",FTP.NotifyID)
	Notify("","",my.RemoteName " - " FormatFileSize(done) " / " FormatFileSize(total), "Text",FTP.NotifyID)
}
FTP_Notify_Error:
ShowSettings("FTP Profiles")
return

Action_Upload_Placeholders_SourceFiles:
GetCurrentSubEvent().GuiShow("", "Placeholders_SourceFiles")
return
Action_Upload_Browse:
GetCurrentSubEvent().GuiShow("", "Browse")
return
Action_Upload_Placeholders_TargetFolder:
GetCurrentSubEvent().GuiShow("", "Placeholders_TargetFolder")
return
Action_Upload_Placeholders_TargetFile:
GetCurrentSubEvent().GuiShow("", "Placeholders_TargetFile")
return