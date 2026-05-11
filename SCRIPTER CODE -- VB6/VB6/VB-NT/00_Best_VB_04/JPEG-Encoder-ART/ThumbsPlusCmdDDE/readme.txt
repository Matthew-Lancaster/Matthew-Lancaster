TPCMD.EXE, Version 2.0b, 10-May-2001
©1993-2001 Cerious Software Inc.
===================================

TPCMD is an command line utility which invokes ThumbsPlus to perform many common 
ThumbsPlus functions (any function which can be performed via DDE).

Cerious Software does not provide official support for TPCMD, but of course
we'll investigate if you have problems!

Command Line Syntax for TPCMD.EXE:
==================================

tpcmd [<command>|-f <scriptfile>] -m -l <logfile> -v -s

Options:
        <command>       DDE command to pass to ThumbsPlus for
			execution.
        -f <scriptfile> Execute the DDE commands in the script
			file named <scriptfile>.
        -m              Minimize ThumbsPlus during execution.
        -l <logfile>    Record any errors and output to the log
			file named <logfile>.
        -v              Verbose output mode.
        -s              Silent operation - no output.

        -h|-?           This help screen.

Notes:
1. You can download the new TPCMD.EXE, a command line utility for executing 
   ThumbsPlus DDE commands from a command line (DOS box) from:
        < ftp://ftp.cerious.com/pub/cerious/tpcmd.zip >.
2. If ThumbsPlus is already running, TPCMD does not start a new instance.
3. You can use multiple commands in a BAT or CMD (on NT) file, or start them 
   from other programs.
4. Many commands accept a file or directory name as a parameter. You may enclose
   the name in quotes if you wish, but this is not required. When the parameter
   is optional for a command, it is identified with the 'optional-' prefix. 
5. You should always use the complete path when specifying files and
   directories.
6. If you use a UNC name (i.e., \\PHILLIP\C\TEMP) for OpenDir or LocateFile,
   ThumbsPlus will automatically map a drive letter for you.
7. Each of the DDE commands covered in this document (with the exception of 
   ExportThumb() and Convert()) are in the ThumbsPlus Help file.


Examples:
=========
  tpcmd -m Open(c:\testing\myimage.jpg)
        Opens and views myimage.jpg. The ThumbsPlus main window will be 
        minimized.

  tpcmd OpenDir(c:\testing)
        Positions to the "c:\testing" directory in ThumbsPlus.

  tpcmd ScanTree(d:\)
        Scans the entire D: drive and makes or updates thumbnails.

  tpcmd -m LocateFile(c:\testing\*.*)
  tpcmd Remove();
  tpcmd MakeThumb();
  tpcmd Exit();
        Opens the Testing directory, removes all thumbnails, then makes them
        again, then exits ThumbsPlus.



DDE Commands Reference
======================

Close(filename)
---------------
Abbreviation: C(filename)

Closes any view window for the file specified.
	Parameters:
		filename	Filename to close view window.


CopyClipboard(optional-filename)	
--------------------------------
Abbreviation:	B(optional-filename)

Copies the currently selected file (if no parameter) or a specific file to the
clipboard. If a file is specified, the directory tree and file list are 
positioned, and the file becomes selected.
	Parameters:
		optional-filename	Filename to copy to clipboard.


Exit()
------
Abbreviation:	X()

Exits the program (closes ThumbsPlus).


Find(keyword-list)
------------------
Abbreviation:	F(keyword-list)

Finds files associated with particular keywords. 

	Parameters:
		keyword-list 	Formatted list of keywords with inclusion specifier.
	Format for keyword-list:
		[how]keyword1;keyword2;
			how 	A single character to indicate whether all ('&'), 
				any ('|'), or most ('*') keywords must match. 
			keyword1;keyword2;...	
				Semicolon delimited list of keywords.

Example: Find all images with keywords: 'raster', '.jpg', and 'truecolor'.

Find("&raster;.jpg;truecolor") 


Keyword(keyword-list|filename)
------------------------------
Abbreviation:	K(keyword-list|filename)

Assigns keywords to or removes keywords from a file.

	Parameters:
		keyword-list 	Formatted list of keywords with inclusion specifier.
		filename	Filename to use.
	Format for keyword-list:
		[+-]keyword1;[+-]keyword2;...
				Semicolon delimited list of keywords. Each
				keyword may be preceded by a '-' or '+' to 
				indicate that the keyword should be removed 
				or added. The default is to add.

Example: Adds keywords 'large', 'dog', 'animal' and removes keyword 'cat' to 
and from file c:\images\animals\dog.jpg.

Keyword("+large;+dog;+animal;-cat|c:\images\animals\dog.jpg") 


LocateFile(filename)
--------------------
Abbreviation:	L(filename)

Locates a file. The directory list is positioned at the file's directory, and
the file itself is selected. You can use a file mask to specify a set of files
to select.

	Parameters:
		filename		Filename or mask to find.

Example: Find and select any files in the C:\Temp directory.

LocateFile(C:\Temp\*.*)


MakeThumb(optional-filename)
----------------------------
Abbreviation:	M(optional-filename)

Creates thumbnails for a specific file or a set of files. If the 
optional-filename parameter is not specified, thumbnails are made for any 
currently selected files. Like LocateFile, this command will accept a mask for 
the file name. The directory list will be repositioned if the file is in a 
different directory from the current one.

	Parameters:
		optional-filename	Filename or mask to thumbnail.

ExportThumb(type;extension;path)
--------------------------------
Abbreviation:	E(type;extension;path)

Extract all selected files' thumbnails as images and save them to files. 

	Parameters:
		type		Represents the file type to write
		extension 	Represents the file extension to use
		path		Provides the path for output files

Example: Export selected thumbnails and saves them as jpeg files with the 
(silly) file extension .any.

ExportThumb(jpg;any;C:\Exported\Jpegs) 


Open(filename)
--------------
Alternate:	FileOpen(filename)
Abbreviation:	O(filename)

Opens a specific file in a view window. The current directory position and file 
selections are not modified by this command, and this command does NOT accept a 
file mask.

	Parameters:
		filename	Filename to view (open).


OpenDB(filename)
----------------
Abbreviation:	DB(filename)

Closes the currently opened ThumbsPlus database and opens another.

	Parameters:
		filename	Filename of a ThumbsPlus database file.


OpenDir(path)
-------------
Abbreviation:	D(path)

Positions the ThumbsPlus directory list in a specific directory.

	Parameters:
		path		Directory path for positioning.


Print(filename)
---------------
Abbreviation:	P(filename)

Prints a specific file. The current directory position and file selections are 
not modified by this command, and it does NOT accept a file mask.

	Parameters:
		filename	Filename to print.


RefreshTree()
-------------
Abbreviation:		T()

Causes ThumbsPlus to re-read the current directory tree, allowing updates from 
other programs to be visible. It's also useful to initialize the tree when 
ThumbsPlus is started minimized (ThumbsPlus does not read the tree by default 
until the main window is visible -- this improves performance when the "simple" 
DDE commands are used, such as Open and Print).


RemoveThumb(optional-filename)
------------------------------
Alternate:	Remove(optional-filename)
Abbreviation:	R(optional-filename)

Removes the thumbnail for the currently selected file(s) (if no parameter) or for 
a specific file. This also removes any associated keywords. Like LocateFile, 
this command will accept a mask for the file name. The directory list will be 
repositioned if the file is in a different directory from the current one.  
Also, if a mask is specified, only the first matching file will have it's 
thumbnail removed from the database.

	Parameters:
		optional-filename	Filename to remove thumbnail.

RemoveTree(optional-path)
-------------------------
Abbreviation:	V(optional-path)

Removes thumbnails from the current directory (if no parameter) or from a 
specific directory, and it's subdirectories.

	Parameters:
		optional-path	Directory to remove thumbnails.

** ThumbsPlus Version 4.15 and higher compatible syntax: 
RemoveTree(optional-path,wait)
	Parameters:
		optional-path	Directory to remove thumbnails.
		wait		Instruct ThumbsPlus to wait for the removal to
				complete before returning control to TPCMD

ScanTree(optional-path)
-----------------------
Abbreviation:	S(optional-path)

Scans the current directory (if no parameter) or a specific directory path and 
creates thumbnails for any found files. 

	Parameters:
		optional-path	Directory to scan and create thumbnails.

** ThumbsPlus Version 4.15 and higher compatible syntax: 
ScanTree(optional-path,wait)
	Parameters:
		optional-path	Directory to scan and create thumbnails.
		wait		Instruct ThumbsPlus to wait for the scan to
				complete before returning control to TPCMD
SlideShow(optional-path)
------------------------
Abbreviation:	W(optional-path)

This function runs a slide show from the current directory (if no parmeter) or 
in a specified directory. If a directory is specified, the directory list is 
repositioned in that directory.

	Parameters:
		optional-path	Directory for slide show.

UpdateAll(optional-path)
------------------------
Abbreviation:	U(optional-path)

Updates all thumbnails in the current directory (if no parameter), or in a 
specified directory. If a directory is specified, the directory list is 
repositioned in that directory.

	Parameters:
		optional-path	Directory for thumbnail update.

Convert(format;extension;retain-comments;reset-date;format-options;output-path)
-------------------------------------------------------------------------------
Abbreviation:	
	N(format;extension;retain-comments;reset-date;format-options;output-path)

Converts any selected files to the specified format. For GIF, TIFF and JPEG 
formats, you may specify various options. Files with the same name are 
overwritten.

You can also specify whether comments should be retained or dropped, and set 
the converted files' creation time to that of the original.

	Parameters:
		format          The 3 character file extension for the output 
				format for the converted files. NOTE: This format
				must be one of the output formats supported by 
				ThumbsPlus.
		                Examples: bmp pcx tga wmf jpg gif tif

		extension       The file extension to use (usually the same as 
				the format).

		retain-comments Specify 0 (numeric zero) to remove image 
				comments, or one 1 (numeric one) to retain them.

		reset-date      When 1 (numeric one), the converted file will
				receive the same date and time as the original 
				file.

		format-options  Options required for some file types: 
				(GIF, JPG, TIF).  See below.

		output-path     The complete directory path where ThumbsPlus will
				create the converted files.

'format-options' for JPEG files:
================================
Parameter       Valid values	Description
---------       ------------    -----------
Quality         1-100           The quality level for JPEG compression. 75
                                is a good default; values below 50 usually
                                generate unacceptable artifacts, and values
                                above 90 usually cause larger files without
                                appreciable visible improvement, so I
                                recommend a number between 50 and 90.

Smoothing       0-50            The amount of smoothing to apply before
                                compression. This is useful for converting
                                dithered (i.e., GIF) images to JPEG. For
                                photographic originals, you should specify
                                zero (0).

Progressive     0, 1            Set this option to one (1) to enable
                                progressive JPEG generations. Note that many
                                programs do not yet support progressive JPEG
                                files.

'format-options' for GIF files:
===============================

Parameter       Valid values	Description
---------       ------------    -----------
Interlaced      0, 1            Set to one (1) to generate an interlaced GIF
                                file.

Minimize Size   0, 1            Set to one (1) to optimize the color table
                                and remove duplicate colors, thus minimizing
                                the file size.

Transparent     0, 1            Set to one (1) to generate a transparent GIF.
                                Use the Transparent Color parameter to
                                specify the transparent color.

Transp. Color   0,0,0           Specify the transparent color when generating
                - 255,255,255   transparent GIF files. This is an RGB value
                                (red, green, blue), and should match one of
                                the pixel values of the GIF file.

'format-options' for TIFF files
===============================

Parameter       Valid values	Description
---------       ------------    -----------
Compression     lzw jpeg fax3   Specifies the compression method for the
                fax4 pack zip   resulting TIFF file.
                none

Quality         1-100           When using JPEG compression in TIFF files,
                                specifies the quality level. See the notes
                                above about JPEG files for further info.

Separate        0, 1            When set to one (1), specifies that image
                                color planes should be stored separately.
                                This often increases compressibility of
                                24-bit image files.

Examples:

Convert the currently selected files to JPG with a quality level of 75, no 
smoothing, non-progressive, while removing comments and setting the file date 
to that of the original file:

	Convert(jpg;jpg;0;1;75;0;0;c:\temp)

Convert the currently selected files to interlaced, non-transparent GIF files. 
Retain image comments and set the date to when the file was converted.

	Convert(gif;gif;1;0;1;1;0;0,0,0;c:\temp)

Convert the currently selected files to Windows Bitmaps. Remove comments and 
leave the date as the converted date.

	Convert(bmp;bmp;0;0;c:\temp)

Convert the currently selected files to TIFF with LZW compression and separated 
image channels. Retain comments and set the TIFF files to the original files' 
dates.

	Convert(tif;abc;1;1;lzw;0;1;c:\temp)

Select all files in the directory c:\camera\pictures\kdc, convert them to JPEG 
with a quality level of 75, eliminating comments and setting the file dates to 
the original KDC file dates, and write the output JPEG files in the 
c:\camera\pictures\jpg directory. Then, use the UpdateDir command to create 
thumbnails for the converted files.

	LocateFile(c:\camera\pictures\kdc\*)
	Convert(jpg;jpg;0;1;75;0;0;c:\camera\pictures\jpg)
	UpdateDir(c:\camera\pictures\jpg)


ThumbsPlus Version 4.15 and Higher
==================================

BatchProcess(set-name;output-path;{input-path;selection-option;file-mask;output-type;overwrite-flag;subdir-flag})
-------------------------------------------------------------------------------
Abbreviation:	
	Z(set-name;output-path;{input-path;selection-option;file-mask;output-type;overwrite-flag;subdir-flag})

Executes a defined batch process.


	Parameters:
		set-name	This value can be any batch set name defined 
				in ThumbsPlus.  This name is case-sensitive.

		output-path     The complete directory path where ThumbsPlus will
				create the converted files.

	Optional Parameters (If omitted, ThumbsPlus will use the settings from the specified batch set):
		input-path	The complete directory path where ThumbsPlus will
				look for the input files.
		selection-option	
				0 - Selected files 
				1 - current folder 
				2 - Current folder and sub-folders

		file-mask	Any valid file mask string.  Eg. *.tif;*.gif
		
		output-type	Any valid file extension for file types that 
				ThumbsPlus can write.  Eg. bmp, jpg, png, etc.

		overwrite-flag	0 - Don't overwrite files existing in the output-path
				1 - Do overwrite existing files

		subdirs-flag	0 - Don't recreate the input-path subfolder structure.  
					Puts all output files directly under output-path, 
					regardless of their input-path subfolder. 
				1 - Do create the same input-path subfolder structure 
					under the output-path.
