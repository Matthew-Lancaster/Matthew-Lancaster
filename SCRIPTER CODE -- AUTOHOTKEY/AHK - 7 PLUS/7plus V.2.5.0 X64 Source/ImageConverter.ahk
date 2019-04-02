Class CImageConverter extends CGUI
{
	/*
	Vars:
	Files := Array()
	AspectRatio
	*/
	__New(Action)
	{
		Base.__New()
		this.Title := "Image Converter & Uploader"
		this.Files := Array()
		this.TemporaryFiles := Action.TemporaryFiles
		this.ReuseWindow := Action.ReuseWindow
		
		this.ListView := this.AddControl("ListView", "ListView", "-Multi NoSort r19 w300", "File|Target Filename")
		this.ListView.IndependentSorting := true
		this.ListView.ModifyCol(1,150)
		this.ListView.ModifyCol(2,"AutoHdr")
		this.Picture := this.AddControl("Picture", "Picture", "x+10 w560 h350 +0xE +0x40", "") ; +0xE is needed for setting the picture to a hbitmap
		this.Picture.Tooltip := "Click the image to edit it in the registered image editing program.`n Save it and quit the program to refresh it here."
		
		this.txtPath := this.AddControl("Text", "txtPath", "x10 y375 Section", "Target Path:")
		this.editPath := this.AddControl("Edit", "editPath", "x+14 ys-4 w196", "")
		this.btnBrowse := this.AddControl("Button", "btnBrowse", "x+5 ys-6 w26", "...")
		this.txtName := this.AddControl("Text", "txtName", "x10 ys+30", "Target Name:")
		this.editName := this.AddControl("Edit", "editName", "x+8 ys+26 w196", "")
		this.txtTargetFormat := this.AddControl("Text", "txtTargetFormat", "x10 ys+60", "Target Format:")
		this.ddlTargetExtension := this.AddControl("DropDownList", "ddlTargetExtension", "Choose1 x+4 ys+56", "Keep Extension|bmp|dib|rle|jpg|jpeg|jpe|jfif|gif|tif|tiff|png")
		
		;Make sure to show quality setting only on specific items
		for index, item in this.ddlTargetExtension.Items
		{
			if(InStr("JPG,JPEG,JPE,JFIF", item.text))
			{
				if(!this.txtQuality)
				{
					this.txtQuality := item.AddControl("Text", "txtQuality", "x+10 ys+60", "Quality:")
					this.editQuality := item.AddControl("Edit", "EditQuality", "x+10 ys+56", Settings.Misc.ImageQuality)
				}
				else
				{
					item.Controls.Insert(this.txtQuality)
					item.Controls.Insert(this.editQuality)
				}
			}
		}
		this.chkOverwriteFiles := this.AddControl("Checkbox", "chkOverwriteFiles", "xs ys+90", "Overwrite existing files")
		this.chkDeleteSourceFiles := this.AddControl("Checkbox", "chkDeleteSourceFiles", "x+4" (Action.TemporaryFiles ? " Checked Disabled" : ""), "Delete source files")
		
		
		this.GrpResize := this.AddControl("GroupBox", "GrpResize", "x320 y360 w300 h100 Section", "Resize")
		this.Radio1 := this.AddControl("Radio", "Radio1","xs+10 ys+20", "Absolute [px]")
		this.Radio2 := this.AddControl("Radio", "Radio2", "xs+10 ys+50 Checked", "Relative  [`%]")
		this.txtWidth1 := this.Radio1.AddControl("Text", "txtWidth1", "xs+100 ys+20", "Width:", 1)
		this.txtWidth2 := this.Radio2.AddControl("Text", "txtWidth2", "xs+100 ys+50", "Width:", 1)
		this.txtHeight1 := this.Radio1.AddControl("Text", "txtHeight1", "xs+200 ys+20", "Height:", 1)
		this.txtHeight2 := this.Radio2.AddControl("Text", "txtHeight2", "xs+200 ys+50", "Height:", 1)
		
		this.editAbsWidth := this.Radio1.AddControl("Edit", "editAbsWidth", "xs+140 ys+16 w50 disabled", "0", 1)
		this.editAbsHeight := this.Radio1.AddControl("Edit", "editAbsHeight", "xs+240 ys+16 w50 disabled", "0", 1)
		this.editRelWidth := this.Radio2.AddControl("Edit", "editRelWidth", "xs+140 ys+46 w50", "0", 1)
		this.editRelHeight := this.Radio2.AddControl("Edit", "editRelHeight", "xs+240 ys+46 w50", "0", 1)
		this.chkKeepAspectRatio := this.AddControl("Checkbox", "chkKeepAspectRatio", "xs+10 ys+75 Checked", "Keep aspect ratio")
		
		
		this.grpCrop := this.AddControl("GroupBox", "grpCrop", "x630 y360 w210 h100 Section", "Crop pixels")
		this.txtLeft := this.AddControl("Text", "txtLeft", "xs+10 ys+20", "Left:")
		this.txtRight := this.AddControl("Text", "txtRight", "xs+110 ys+20", "Right:")
		this.txtTop := this.AddControl("Text", "txtTop", "xs+10 ys+50", "Top:")
		this.txtBottom := this.AddControl("Text", "txtBottom", "xs+110 ys+50", "Bottom:")
		this.editCropLeft := this.AddControl("Edit", "editCropLeft", "xs+50 y376 w50", "0")
		this.editCropRight := this.AddControl("Edit", "editCropRight", "xs+150 y376 w50", "0")
		this.editCropTop := this.AddControl("Edit", "editCropTop", "xs+50 y406 w50", "0")
		this.editCropBottom := this.AddControl("Edit", "editCropBottom", "xs+150 y406 w50", "0")
		this.txtUpload := this.AddControl("Text", "txtUpload", "x10 y486", "Upload")
		this.ddlWhichFiles := this.AddControl("DropDownList", "ddlWhichFiles", "x+10 y482 w70", "selected||all")
		this.txtFilesTo := this.AddControl("Text", "txtFilesTo", "x+10 y486", "files to:")
		for index, FTPProfile in  CFTPUploadAction.FTPProfiles
			Hosters .= (index != 1 ? "|" : "") index ": " FTPProfile.Hostname (Action.Hoster = index ? "|" : "")
		Hosters .= "|" GetImageHosterList().ToString("|")
		if(!IsNumeric(Action.Hoster))
			Hosters := RegexReplace(Hosters, Action.Hoster "\|?", Action.Hoster "||")
		this.ddlHoster := this.AddControl("DropDownList", "ddlHoster", "x+10 y482 w140", Hosters)
		
		;Make ftp target directory controls enabled only when an ftp server is selected
		for index, item in this.ddlHoster.Items
		{
			if(IsNumeric(SubStr(item.text, 1, 1)))
			{
				if(!this.txtDirectory)
				{
					this.txtDirectory := item.AddControl("Text", "txtDirectory", "x+10 y486", "Directory:", 1)
					this.editFTPTargetDir := item.AddControl("Edit", "editFTPTargetDir", "x+10 y482 w120", Action.FTPTargetDir, 1)
				}
				else
				{
					item.Controls.Insert(this.txtDirectory)
					item.Controls.Insert(this.editFTPTargetDir)
				}
			}
		}
		this.btnUpload := this.AddControl("Button", "btnUpload", "x+5 y481 w69", "&Upload")
		this.btnCopyToClipboard := this.AddControl("Button", "btnCopyToClipboard", "x+10 y481 w95", "&Copy to Clipboard")
		this.btnConvertAndSave := this.AddControl("Button", "btnConvertAndSave", "x+10 y481 w124", "Convert && &Save && Close")
		this.btnCancel := this.AddControl("Button", "btnCancel", "x+5 y481 w60", "Cancel")
		this.OnMessage(WM_COMMAND:=0x111, "MessageHandler") ;Might not be working right now
		LV_Modify(1,"Select")
		this.CloseOnEscape := true
		this.DestroyOnClose := true
		this.Show("Autosize")
	}
	PreClose()
	{
		;Possibly delete old files
		;Get list of involved files
		Files := []
		TargetDir := this.editPath.Text
		for index, Item in this.ListView.Items
		{
			SplitPath(this.Files[index].SourceFile, "", SourceDir)
			Target := (TargetDir ? TargetDir : SourceDir) "\" Item[2]
			;Strategy: Delete all files from temp directory (from screenshots) and source files if the user chooses to delete them after the conversion process.
			;		   Images which couldn't be converted (and therefore don't exist in their target destination) are left untouched, same for overwritten files
			if((this.chkDeleteSourceFiles.Checked && FileExist(Target) && this.Files[index].SourceFile != Target) || InStr(this.Files[index].SourceFile, A_Temp "\7plus\"))
				Files.Insert(this.Files[index].SourceFile)
		}
		;Check if an upload is still running and possibly attach an action to it to delete the involved files
		;Enable Critical so the event processing doesn't interfere here (should be unlikely but better be sure)
		Critical, On
		if(this.tmpQueuedUploadEvent && EventSystem.EventSchedule.IndexOf(this.tmpQueuedUploadEvent))
		{
			Action := new CFileDeleteAction()
			Action.SourceFile := this.ddlWhichFiles.Text = "All" ? Files : this.tmpQueuedUploadEvent.Actions[1].SourceFiles
			Action.Silent := 1
			this.tmpQueuedUploadEvent.Actions.Insert(Action)
			for index, File in Action.SourceFile ;Remove all files which will be delete when the upload action finishes so that all other files can be deleted here
				Files.Delete(File)
		}
		Critical, Off
		;Delete all other files
		for index, File in Files
			FileDelete, % File
	}
	;Adds files to the list of images
	AddFiles(Files)
	{
		Files := ToArray(Files)
		if(Files.MaxIndex() < 1)
		{
			Msgbox Image Converter: No files selected!
			return
		}
		Added := false
		for index, file in Files
		{
			SplitPath(file, Filename, "", Extension)
			if Extension not in BMP,DIB,RLE,JPG,JPEG,JPE,JFIF,GIF,TIF,TIFF,PNG
			{
				Msgbox % file " is no supported image format!"
				continue
			}
			if(!this.Files.GetItemWithValue("SourceFile", file)) ;Append files which aren't yet in the list
			{
				this.ListView.Items.Add("", Filename)
				this.Files.Insert({SourceFile : File})
				Added := true
			}
		}
		if(Added)
		{
			this.FillTargetFilenames()
			this.ListView.SelectedIndex := this.Files.MaxIndex() - Files.MaxIndex() + 1
			if(this.TemporaryFiles && this.Files.MaxIndex() = Files.MaxIndex()) ;Select first file if list was empty before
			{
				SplitPath(this.Files[1].SourceFile,"",Path)
				this.editPath.Text := Path
			}
			this.Show()
		}
		else
			Msgbox Image converter: No new files!
		return
	}
	
	;Calculates target filenames and shows them
	FillTargetFilenames()
	{
		Targets := Array()
		for index, item in this.ListView.Items
		{
			SplitPath(this.Files[A_Index].SourceFile, "", dir, Extension, Filename)
			if(this.Files[A_Index].HasKey("TargetFilename") && this.Files[A_Index].TargetFilename != "")
				Filename := this.Files[A_Index].TargetFilename
			if(this.ddlTargetExtension.Text != "Keep extension")
				Extension := this.ddlTargetExtension.Text
				
			Testpath := (this.editPath.Text ? this.editPath.Text : dir) "\" Filename "." Extension
			i:=1 ;Find free Filename
			while((!this.chkOverwriteFiles.Checked && FileExist(TestPath)) || (Targets.IndexOf(TestPath) > 0 && Targets.IndexOf(TestPath) < index))
			{
				i++
				Testpath:=dir "\" Filename " (" i ")." Extension
			}
			SplitPath(Testpath,Filename)
			
			Targets[index] := TestPath
			this.ListView.Items[A_Index][2] := Filename
		}
	}
	LoadImage()
	{
		Selected := this.Files[this.ListView.SelectedIndex ? this.ListView.SelectedIndex : this.ListView.FocusedIndex].SourceFile
		if(Selected != this.Picture.Picture)
		{
			this.PreviousPicturePath := this.Picture.Picture
			SplitPath(this.Picture.Picture, cFilename, cPath)
			WatchDirectory(cPath "|" cFilename "\", "")
		}
		pBitmap := Gdip_CreateBitmapFromFile(Selected)
		Width := Gdip_GetImageWidth(pBitmap)
		Height := Gdip_GetImageHeight(pBitmap)
		this.AspectRatio := Width/Height
		sw := min(560 / Width, 1)
		sh := min(350 / Height, 1)
		s := min(sw, sh)
		pThumbnail := Gdip_CreateBitmap(560,350)
		pGraphics := Gdip_GraphicsFromImage(pThumbnail)
		Gdip_SetInterpolationMode(pGraphics, 7)
		Gdip_DrawImage(pGraphics, pBitmap, 0, 0, Width*s, Height*s)
		Gdip_DeleteGraphics(pGraphics)
		Gdip_DisposeImage(pBitmap)
		hBitmap := Gdip_CreateHBITMAPFromBitmap(pThumbnail)
		this.Picture.SetImageFromHBitmap(hBitmap)
		this.Picture._.Picture := Selected ;Using internal details of classes isn't the nice way, but I need to store this path here :P
		this.Picture._.PictureWidth := Width
		this.Picture._.PictureHeight := Height
		; this.Picture.Picture := Selected ;Set picture by path so its path property can be retrieved later.   ;;;;Unfortunately this won't preserve the aspect ratio
		DeleteObject(hBitmap)
		Gdip_DisposeImage(pThumbnail)
		this.editAbsWidth.Text := this.Picture.PictureWidth
		this.editAbsHeight.Text := this.Picture.PictureHeight
		this.editRelWidth.Text := "100"
		this.editRelHeight.Text := "100"
		SplitPath(Selected, Filename, Path)
		WatchDirectory(Path "|" Filename "\", "ImageConverter_OpenedFileChange")
	}
	ListView_SelectionChanged(Row)
	{
		if(this.ListView.SelectedIndex && this.ListView.SelectedIndex != this.ListView.PreviouslySelectedIndex)
		{
			this.LoadImage()
			SplitPath(this.ListView.SelectedItem[2], "", "", "", FileName)
			this.editName.Text := FileName
			this.ActiveControl := this.ListView
		}
	}
	editName_TextChanged()
	{
		this.Files[(index := this.ListView.SelectedIndex) ? index : this.ListView.FocusedIndex].TargetFilename := this.editName.Text
		this.FillTargetFilenames()
	}
	Picture_Click()
	{
		; DllCall("shell32\ShellExecute"uint, 0, str, "Edit"str, this.Files[Selected].SourceFile, str, "", str, "", int, 1)
		run, % "edit """ this.Picture.Picture """",,UseErrorLevel
	}
	editAbsWidth_TextChanged()
	{
		if(this.Lock<=0)
		{
			this.LastChanged := "Width"
			this.Lock := 1
			this.editRelWidth.Text := Round(this.editAbsWidth.Text / this.Picture.PictureWidth * 100)			
			if(this.chkKeepAspectRatio.Checked)
			{
				this.Lock += 2
				this.editAbsHeight.Text := Round(this.editAbsWidth.Text / this.AspectRatio)
				this.editRelHeight.Text := this.editRelWidth.Text
			}
		}
		else
			this.Lock--
	}
	editAbsHeight_TextChanged()
	{
		if(this.Lock<=0)
		{
			this.LastChanged := "Height"
			this.Lock := 1
			this.editRelHeight.Text := Round(this.editAbsHeight.Text / this.Picture.PictureHeight * 100)
			if(this.chkKeepAspectRatio.Checked)
			{
				this.Lock += 2
				this.editAbsWidth.Text := Round(this.editAbsHeight.Text * this.AspectRatio)			
				this.editRelWidth.Text := this.editRelHeight.Text
			}
		}
		else
			this.Lock--
	}
	editRelWidth_TextChanged()
	{
		if(this.Lock<=0)
		{
			LastChanged := "Width"
			this.Lock := 1
			this.editAbsWidth.Text := Round(this.Picture.PictureWidth * this.editRelWidth.Text / 100)
			if(this.chkKeepAspectRatio.Checked)
			{
				this.Lock += 2
				this.editRelHeight.Text := this.editRelWidth.Text
				this.editAbsHeight.Text := Round(this.Picture.PictureHeight * this.editRelHeight.Text / 100)
			}
		}
		else
			this.Lock--
	}
	editRelHeight_TextChanged()
	{
		if(this.Lock<=0)
		{
			LastChanged := "Height"
			this.Lock := 1
			this.editAbsHeight.Text := Round(this.Picture.PictureHeight * this.editRelHeight.Text / 100)
			if(this.chkKeepAspectRatio.Checked)
			{
				this.Lock += 2
				this.editRelWidth.Text := this.editRelHeight.Text
				this.editAbsWidth.Text := Round(this.Picture.PictureWidth * this.editRelWidth.Text / 100)
			}
		}
		else
			this.Lock--
	}
	chkKeepAspectRatio_CheckedChanged()
	{
		if(this.chkKeepAspectRatio.Checked)
		{
			if(LastChanged = "Width")
			{
				if(this.Radio1.Checked)
					this.editAbsWidth_TextChanged()
				else
					this.editRelWidth_TextChanged()
			}
			else
			{
				if(this.Radio1.Checked)
					this.editAbsHeight_TextChanged()
				else
					this.editRelHeight_TextChanged()
			}
		}
	}
	
	btnCopyToClipboard_Click()
	{	
		SplitPath(this.Picture.Picture, "", FilePath)
		FilePath .= "\" FileName
		pConverted := this.ConvertSingleImage(this.Picture.Picture, FilePath, changed)
		if(pConverted)
		{
			Gdip_SetBitmapToClipboard(pConverted)
			Gdip_DisposeImage(pConverted)
		}
		else
			Msgbox Failed to convert image!
	}
	btnUpload_Click()
	{
		this.Upload()
	}
	btnConvertAndSave_Click()
	{
		this.ConvertAndSave()
	}
	btnCancel_Click()
	{
		this.Close()
	}
	chkOverwriteFiles_CheckedChanged()
	{
		this.FillTargetFilenames()
	}
	btnBrowse_Click()
	{
		FolderDialog := new CFolderDialog()
		SplitPath(this.editPath.Text ? this.editPath.Text : this.Files[this.ListView.SelectedIndex ? this.Listview.SelectedIndex : this.ListView.FocusedIndex].SourceFile, "", Path)
		FolderDialog.Folder := Path
		FolderDialog.Title := "Select output directory"
		if(!FolderDialog.Show())
			return false
		this.editPath.Text := FolderDialog.Folder
		this.FillTargetFilenames()
		return true
	}
	;Converts the images contained in the ListView and saves them at their target paths.
	;ConvertedImages and FailedImages are output parameters from this function and should be initiliazed as arrays before calling this function.
	;If TemporaryFiles is set the function will save the images to a temporary location so they can easily be deleted after uploading.
	;If the conversion process is canceled this function returns false, otherwise true.
	ConvertImages(ConvertedImages, FailedImages, SelectedOnly = 0, TemporaryFiles = 0)
	{
		if(TemporaryFiles)
			FileCreateDir, % A_Temp "\7plus\Upload"
		else if(this.TemporaryFiles && !this.editPath.Text)
		{
			if(!this.btnBrowse_Click())
				return false
		}
		for index, item in this.ListView.Items
		{
			if(SelectedOnly && !item.Selected)
				continue
			Source := this.Files[A_Index].SourceFile
			SplitPath(Source,"", SourceDir)
			Target := TemporaryFiles ? A_Temp "\7plus\Upload\" item[1] : ((this.editPath.Text ? this.editPath.Text : (this.TemporaryFiles ? FolderDialog.Folder : SourceDir)) "\" item[2])
			pConverted := this.ConvertSingleImage(Source, Target, changed)
			if(pConverted)
			{
				if((!changed && Source = Target) || Gdip_SaveBitmapToFile(pConverted, Target, this.editQuality.Text) = 0)
					ConvertedImages.Insert(Target)
				else
					FailedImages.Insert(Source)
				Gdip_DisposeImage(pConverted)
			}
			else
				FailedImages.Insert(Source)
		}
		return true
	}
	ConvertAndSave()
	{
		if(this.ConvertImages(ConvertedImages := Array(), FailedImages := Array()))
		{
			if(FailedImages.MaxIndex() > 0)
			{
				for index, image in FailedImages
					Files .= (index := 1 ? "" : "`n") Image
				Notify("Image Conversion failed!", "Failed to convert these files:`n" Files, 5, "GC=555555 TC=White MC=White", NotifyIcons.Error)
			}
			else
				Notify("Image Conversion completed!", "Successfully converted " ConvertedImages.MaxIndex() " files.", 5, "GC=555555 TC=White MC=White", NotifyIcons.Success)
			this.Close()
		}
	}
	
	Upload()
	{
		this.ConvertImages(ConvertedImages := Array(), FailedImages := Array(), this.ddlWhichFiles.Text = "Selected", 1)
		;Let's build an event that uploads the files using the selected hoster and deletes them afterwards (if they are temporary (screenshot) files located in temp dir)
		Event := new CEvent()
		if(IsNumeric(SubStr(this.ddlHoster.Text, 1,max(InStr(this.ddlHoster.Text, ":") - 1, 1))))
		{
			Event.Actions.Insert(new CFTPUploadAction())
			Event.Actions[1].SourceFiles := ConvertedImages
			Event.Actions[1].TargetFolder := this.editFTPTargetDir.Text
			Event.Actions[1].FTPProfile := SubStr(this.ddlHoster.Text, 1,max(InStr(this.ddlHoster.Text, ":") - 1, 1))
		}
		else
		{
			Event.Actions.Insert(new CImageUploadAction())
			Event.Actions[1].SourceFiles := ConvertedImages
		}
		Event.Actions.Insert(new CFileDeleteAction())
		Event.Actions[2].SourceFile := ConvertedImages
		Event.Actions[2].Silent := 1
		EventSystem.TemporaryEvents.RegisterEvent(Event)
		this.tmpQueuedUploadEvent := Event.TriggerThisEvent()
	}
	
	ConvertSingleImage(OldFile, NewFile, ByRef changed)
	{
		UseAbsoluteSizes := this.Radio1.Checked
		Width := UseAbsoluteSizes ? this.editAbsWidth.Text : this.editRelWidth.Text
		Height := UseAbsoluteSizes ? this.editAbsHeight.Text : this.editRelHeight.Text
		KeepAspectRatio := this.chkKeepAspectRatio.Checked
		CropLeft := abs(this.editCropLeft.Text)
		CropRight := abs(this.editCropRight.Text)
		CropTop := abs(this.editCropTop.Text)
		CropBottom := abs(this.editCropBottom.Text)
		Assert(OldFile && FileExist(OldFile))
		Changed := false
		pBitmap := Gdip_CreateBitmapFromFile(OldFile)
		if(pBitmap > 0)
		{
			if(!(Width > 0))
				Width := UseAbsoluteSizes ? Gdip_GetImageWidth(pBitmap) : 100
			if(!(Height > 0))
				Height := UseAbsoluteSizes ? Gdip_GetImageHeight(pBitmap) : 100
			SplitPath(OldFile,"","",OldExt)		
			SplitPath(NewFile,"","",NewExt)
			;Calculate sizes
			w_old := Gdip_GetImageWidth(pBitmap) - (CropLeft + CropRight)
			h_old := Gdip_GetImageHeight(pBitmap) - (CropTop + CropBottom)
			AR_old := w_old/h_old
			if(UseAbsoluteSizes)
			{
				w_new := Width
				h_new := Height
				;Since multiple files can be processed, the aspect ratio might be different.
				;The approach used here looks for the width or height scaling factor that is closer to 1 and applies it to the whole image.
				if(KeepAspectRatio)
				{
					AR_w := w_new / w_old
					AR_h := h_new / h_old
					if(Abs(AR_w - 1) < Abs(AR_h - 1))
					{
						w_new := w_old * AR_w
						h_new := h_old * AR_w
					}
					else
					{
						w_new := w_old * AR_h
						h_new := h_old * AR_h
					}
				}
			}
			else
			{
				w_new := w_old * Width / 100
				h_new := h_old * Height / 100
			}
			;Save image
			if(w_new != w_old || h_new != h_old || CropLeft > 0 || CropRight > 0 || CropTop > 0 || CropBottom > 0)
			{
				pConverted := Gdip_CreateBitmap(w_new,h_new)
				pGraphics := Gdip_GraphicsFromImage(pConverted)			
				Gdip_SetInterpolationMode(pGraphics, 7)
				Gdip_DrawImage(pGraphics, pBitmap, 0, 0, w_new, h_new, CropLeft, CropTop, w_old, h_old)
				Gdip_DeleteGraphics(pGraphics)
				Gdip_DisposeImage(pBitmap)
				changed := true
			}
			else
			{
				pConverted := pBitmap
				pBitmap := ""
			}
			if(OldFile != NewFile)
				changed := true
			return pConverted
		}
		return 0
	}
	
	;Not working right now
	MessageHandler(wParam, lParam, msg, hwnd)
	{
		global
		static aw,ah,rw,rh,quality		
		local Word1:=(wParam&0xFFFF0000)>>16
		local WM_COMMAND:=0x111
		local EN_SETFOCUS:=0x100 
		local EN_KILLFOCUS:=0x200
		local aw_h,ah_h,rw_h,rh_h,quality_h
		GuiControlGet, aw_h, Hwnd, this.editAbsWidth.Text
		GuiControlGet, ah_h, Hwnd, this.editAbsHeight.Text
		GuiControlGet, rw_h, Hwnd, this.editRelWidth.Text
		GuiControlGet, rh_h, Hwnd, this.editRelHeight.Text
		GuiControlGet, quality_h, Hwnd, ImageConverter_Quality
		if(msg = WM_COMMAND)
		{
			if(lParam = aw_h || lParam = ah_h || lParam = rw_h || lParam = rh_h || lParam = quality_h) ;Handle Focus messages to validate input
			{
				if(Word1 = EN_SETFOCUS)
				{
					GUI, Submit, NoHide
					aw := this.editAbsWidth.Text
					ah := this.editAbsHeight.Text
					rw := this.editRelWidth.Text
					rh := this.editRelHeight.Text
					quality := ImageConverter_Quality
				}
				else if(Word1 = EN_KILLFOCUS)
				{
					local text
					ControlGetText, Text, ,ahk_id %lParam%
					if((!IsNumeric(text) || text <= 0) && lParam != quality_h)
					{
						CGUI.GUIList.GetItemWithValue("hwnd", hwnd).Lock := 4
						GuiControl, , this.editAbsWidth.Text, %aw%
						GuiControl, , this.editAbsHeight.Text, %ah%
						GuiControl, , this.editRelWidth.Text, %rw%
						GuiControl, , this.editRelHeight.Text, %rh%
					}
					else if(lParam = quality_h && (!IsNumeric(text) || text <= 0 || text > 100))					
						GuiControl, , ImageConverter_Quality, %quality%
				}
			}
		}
	}
	;Not working right now
	ImageConverter_MessageHandler(wParam, lParam, msg, hwnd)
	{
		for index, Event in EventSystem.EventSchedule
			for index2, Action in Event
				if(Action.tmpImageConverterClass.GUINum = A_GUI)
					Action.tmpImageConverterClass.MessageHandler(wParam, lParam, msg, hwnd)
	}
}

ImageConverter_OpenedFileChange(from, to) ;This gets called when a file that is being watched was modified
{
	for index, ImageConverter in CImageConverter.Instances
	{
		if(ImageConverter.Picture.Picture = from)
			ImageConverter.LoadImage()
	}
	; outputdebug file change %from% %to%
	; ImageConverter("","","ImageConverter_LoadPicture")
	; WatchDirectory(from, "")
}