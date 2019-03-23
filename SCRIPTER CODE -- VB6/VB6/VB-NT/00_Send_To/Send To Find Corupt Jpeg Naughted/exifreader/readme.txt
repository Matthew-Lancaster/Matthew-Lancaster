Thanks for downloading the ExifReader class for VB! Any help/suggestions you have for making this class better are VERY welcomed and encouraged!

Public Methods:
---------------
	Load(Optional picFile as String)
		Will load the JPG file into memory and retrieve Exif information.


Public Properties:
------------------
	picFile
		Optional. Sets the JPG file location (path). Use before calling the Load method.

	Tag(Optional ExifTag as Long)
		ExifTag is an enumeration that will list all Exif tags. You can use custom tags by using a hex number for 				ExifTag.
		Example: objExifReader.Tag(&H8769&) will have the same result as objExifReader.Tag(ExifOffset)

	MakerNoteTag(Optional MakerTag as Long)
		Under construction. Use at your own risk.



Below is an example using this class in VB 6:

	Dim objExif as New ExifReader
	Dim txtExifInfo as String
	
	objExif.Load "C:\Path_To_Jpg.jpg"
	txtExifInfo = objExif.Tag(DateTimeOriginal)
	MsgBox txtExifInfo
