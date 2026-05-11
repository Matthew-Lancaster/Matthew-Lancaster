VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "ClipBoard Logger"
   ClientHeight    =   5196
   ClientLeft      =   3696
   ClientTop       =   6096
   ClientWidth     =   14820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5196
   ScaleWidth      =   14820
   WindowState     =   1  'Minimized
   Begin RichTextLib.RichTextBox RTB_CLIPPER_FORMAT_CONVERTING 
      Height          =   960
      Left            =   1968
      TabIndex        =   11
      Top             =   1620
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   1693
      _Version        =   393217
      TextRTF         =   $"ClipBoard Logger.frx":0000
   End
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   7356
      Top             =   228
   End
   Begin VB.Timer Timer_API_RESET_INFO 
      Interval        =   1000
      Left            =   4770
      Top             =   1665
   End
   Begin VB.Timer Timer_APP_BEGIN_TIMER 
      Interval        =   1000
      Left            =   6690
      Top             =   2370
   End
   Begin VB.Timer Timer_UNLOAD_FORM 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6690
      Top             =   1875
   End
   Begin VB.Timer Timer_KEYBOARD_Active 
      Interval        =   50
      Left            =   5650
      Top             =   2325
   End
   Begin VB.Timer TIMER_KEYBOARD_AND_MOUSE_ACTIVE 
      Enabled         =   0   'False
      Interval        =   59000
      Left            =   5650
      Top             =   330
   End
   Begin VB.FileListBox File3 
      Height          =   252
      Left            =   2450
      TabIndex        =   10
      Top             =   800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Timer TIMER_VB_PROJECT_DATE 
      Interval        =   4000
      Left            =   5360
      Top             =   40
   End
   Begin VB.Timer TIMER_RETRY_WRITE_INFO_UNTIL_DONE1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5360
      Top             =   330
   End
   Begin VB.Timer TIMER_RETRY_WRITE_INFO_UNTIL_DONE2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5360
      Top             =   630
   End
   Begin VB.Timer Timer_MENU_HEIGHT_CHANGED 
      Interval        =   100
      Left            =   6120
      Top             =   1490
   End
   Begin VB.Timer Timer_MOUSE_5_MINUTE 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5650
      Top             =   1780
   End
   Begin VB.Timer Timer_Idle_Few_Second 
      Interval        =   1000
      Left            =   6110
      Top             =   1200
   End
   Begin VB.Timer Timer_MOUSE_4_MINUTE 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5650
      Top             =   1490
   End
   Begin VB.Timer Timer_MOUSE_3_MINUTE 
      Interval        =   59990
      Left            =   5650
      Top             =   1200
   End
   Begin VB.Timer TIMER_OutClipChunck_Txt 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5060
      Top             =   910
   End
   Begin VB.Timer Timer_INFORAPID_MSGBOX 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5060
      Top             =   620
   End
   Begin VB.Timer Timer_MOUSE_2_MINUTE 
      Enabled         =   0   'False
      Interval        =   59990
      Left            =   5660
      Top             =   910
   End
   Begin VB.Timer Timer_MOUSE_1_MINUTE 
      Enabled         =   0   'False
      Interval        =   59990
      Left            =   5650
      Top             =   620
   End
   Begin VB.Timer Timer_EXE_DAY_AGE 
      Interval        =   60000
      Left            =   5060
      Top             =   50
   End
   Begin VB.Timer Timer_API_OKAY_COLOUR 
      Interval        =   1
      Left            =   4760
      Top             =   1200
   End
   Begin VB.Timer Timer_API_Test 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4760
      Top             =   910
   End
   Begin VB.Timer Timer_WEATHER_URL 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6120
      Top             =   40
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   3330
      TabIndex        =   7
      Top             =   980
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Timer Timer_EXPLORER_GONE 
      Interval        =   1000
      Left            =   4760
      Top             =   330
   End
   Begin VB.Timer Timer_MOUSE_KEYBOARD 
      Interval        =   1000
      Left            =   5650
      Top             =   2070
   End
   Begin VB.Timer Timer_SCREEN_SHOT 
      Interval        =   100
      Left            =   4760
      Top             =   620
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   330
   End
   Begin VB.Timer Timer_MOUSE 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5650
      Top             =   40
   End
   Begin VB.Timer Timer_FORM_RESIZE 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   5060
      Top             =   330
   End
   Begin VB.Timer TimerSCROLL 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   6120
      Top             =   630
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4460
      Top             =   330
   End
   Begin VB.Timer TIMER_JOYPAD 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6120
      Top             =   910
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   3920
      ScaleHeight     =   372
      ScaleWidth      =   420
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   3340
      ScaleHeight     =   384
      ScaleWidth      =   468
      TabIndex        =   4
      Top             =   490
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Picture2 
      Height          =   420
      Left            =   3910
      ScaleHeight     =   372
      ScaleWidth      =   444
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   405
      Left            =   3330
      ScaleHeight     =   360
      ScaleWidth      =   492
      TabIndex        =   2
      Top             =   40
      Visible         =   0   'False
      Width           =   540
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   30
      TabIndex        =   0
      Top             =   820
      Visible         =   0   'False
      Width           =   2360
      _ExtentX        =   4995
      _ExtentY        =   593
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4460
      Top             =   40
   End
   Begin MCI.MMControl MMControl2 
      Height          =   330
      Left            =   30
      TabIndex        =   6
      Top             =   1170
      Visible         =   0   'False
      Width           =   2360
      _ExtentX        =   4995
      _ExtentY        =   572
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   408
      Left            =   36
      TabIndex        =   1
      Top             =   360
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   699
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"ClipBoard Logger.frx":0079
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "TIMERS NOT USED"
      Height          =   1720
      Left            =   6420
      TabIndex        =   9
      Top             =   40
      Visible         =   0   'False
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "Label1 -- COLOUR BAR"
      Height          =   255
      Left            =   45
      TabIndex        =   8
      Top             =   60
      Width           =   3225
   End
   Begin VB.Menu Mnu_Exit 
      Caption         =   "Exit"
   End
   Begin VB.Menu MNU_AUDIO_WANT_ON 
      Caption         =   "AUDIO WANT ON"
   End
   Begin VB.Menu MNU_SCROLL_BOTTOM_OFF__WANT_ON_ASK_PRESS 
      Caption         =   "SCROLL BOTTOM OFF -- WANT ON ASK PRESS"
   End
   Begin VB.Menu MNU_API_RESET 
      Caption         =   "API"
   End
   Begin VB.Menu MNU_SLECTOR_WHOLE_LINE_MODE 
      Caption         =   "SELECTOR WHOLE LINE MODE=FALSE"
   End
   Begin VB.Menu AUTO_CLIPBOARD_SELECTOR 
      Caption         =   "AUTO_CLIPBOARD_SELECTOR_>2"
   End
   Begin VB.Menu MNU_EXE_LANUCHER 
      Caption         =   "EXE MENU"
      Begin VB.Menu MNU_TSF 
         Caption         =   "TREE SIZE FREE"
      End
      Begin VB.Menu MNU_TSP 
         Caption         =   "TREESIZE PRO"
      End
      Begin VB.Menu MNU_TSS 
         Caption         =   "TREESIZE SEARCH"
      End
      Begin VB.Menu MNU_FLICKR_YAHOO 
         Caption         =   "FLICKR YAHOO"
      End
      Begin VB.Menu MNU_FLCIKR_SYNC 
         Caption         =   "FLCIKR SYNC"
      End
      Begin VB.Menu MNU_MULTI_MENU 
         Caption         =   "VB MULTI MENU"
      End
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_EXE_NEWER_IN_MIRROR 
      Caption         =   "**** VB EXE NEWER IN MIRROR ****"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_CLIP_TIME 
      Caption         =   "GIVE ME TIME -->"
   End
   Begin VB.Menu MNU_COMPARE 
      Caption         =   "COMPARE 1 2"
   End
   Begin VB.Menu MNU_AUDIO_01_MISSING 
      Caption         =   "AUDIO 01 MISSING"
   End
   Begin VB.Menu MNU_AUDIO_02_MISSING 
      Caption         =   "AUDIO 02 MISSING"
   End
   Begin VB.Menu MNU_COUNT_JUMP_UP 
      Caption         =   "<-- UP"
   End
   Begin VB.Menu MNU_COUNT_JUMP_DOWN 
      Caption         =   "DOWN"
   End
   Begin VB.Menu MNU_COUNT_JUMP_TOP 
      Caption         =   "TOP"
   End
   Begin VB.Menu MNU_COUNT_JUMP_BOTTOM 
      Caption         =   "BOTTOM -->"
   End
   Begin VB.Menu MNU_BOTTOM 
      Caption         =   "BOTTOM NEED MORE WORK"
   End
   Begin VB.Menu Mnu_Minimize 
      Caption         =   "<**** Minimize"
   End
   Begin VB.Menu Mnu_Norm 
      Caption         =   "* NORM *"
   End
   Begin VB.Menu Mnu_Center 
      Caption         =   "* CENTER *"
   End
   Begin VB.Menu Mnu_Max 
      Caption         =   "* Maximized ****>"
   End
   Begin VB.Menu MNU_NOTEPAD__ 
      Caption         =   "<**  NOTEPAD++  **>"
   End
   Begin VB.Menu MNU_REG_JUMP 
      Caption         =   "<**  REG JUMP **>"
   End
   Begin VB.Menu MNU_NIRSOFT_2 
      Caption         =   "** NIRSOFT **"
      Begin VB.Menu MNU_NIRSOFT 
         Caption         =   "NIR1"
         Index           =   1
      End
      Begin VB.Menu MNU_NIRSOFT 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu MNU_NIRSOFT 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu MNU_NIRSOFT 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu MNU_NIRSOFT 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu MNU_NIRSOFT 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu MNU_NIRSOFT 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu MNU_NIRSOFT 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu MNU_NIRSOFT 
         Caption         =   ""
         Index           =   9
      End
   End
   Begin VB.Menu MNU_JUMP_ANY_SPECIAL_FOLDER 
      Caption         =   "<** ANY SPECIAL DIR **>"
   End
   Begin VB.Menu MNU_TOOL 
      Caption         =   "* EXPLORER FOLDER TOOL SHOP MENU*"
      Begin VB.Menu MNU_CODE_MENU 
         Caption         =   "CODE_MENU"
         Begin VB.Menu MNU_TEST_STOP_ALL_TIMER 
            Caption         =   "TEST STOP ALL TIMER"
         End
      End
      Begin VB.Menu MNU_EXPLORER_ME_VB 
         Caption         =   "EXPLORER -- ME_VB"
      End
      Begin VB.Menu MNU_SEND_TO_SYSTEM_FOLDER 
         Caption         =   "EXPLORER -- SEND TO SYSTEM FOLDER"
      End
      Begin VB.Menu MNU_SEND_TO_FAT32_FOLDER 
         Caption         =   "EXPLORER -- SEND TO FAT 32 FOLDER"
      End
      Begin VB.Menu MNU_SPECIAL_FOLDER_PIN_TO_START_MENU 
         Caption         =   "EXPLORER -- PIN TO START MENU - FOLDER"
      End
      Begin VB.Menu MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC 
         Caption         =   "EXPLORER -- DESKTOP FOLDER - PUBLIC"
      End
      Begin VB.Menu MNU_SPECIAL_FOLDER_DESKTOP_USER 
         Caption         =   "EXPLORER -- DESKTOP FOLDER - USER"
      End
      Begin VB.Menu MNU_MIRROR_SEND_TO_OPERATING_SYSTEM 
         Caption         =   "MIRROR COPY CUSTOM SEND_TO FOLDER TO ANY OPERATING SYSTEM SEND_TO SPECIAL FOLDER"
      End
      Begin VB.Menu MNU_ABORT_SHUTDOWN 
         Caption         =   "ABORT SHUTDOWN \system32\shutdown.exe /a"
      End
   End
   Begin VB.Menu Mnu_Options 
      Caption         =   "* OPTIONS MENU *"
      Begin VB.Menu Mnu_SoundOn 
         Caption         =   "AUDIO ON"
      End
      Begin VB.Menu MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM 
         Caption         =   "TOGGLE NOT SCROLL BACK TO BOTTOM WHILE WORKING"
      End
      Begin VB.Menu MNU_PLAY_SOUND_ON_LOAD 
         Caption         =   ""
      End
      Begin VB.Menu MNU_Audio_Only_With_Text_and_Picture_Clip_Sound_Bug_Acer 
         Caption         =   "Audio Only With Text and Picture Clip -- PROBLEM -- Sound Bug Acer"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Edit_Sound 
         Caption         =   "EDIT AUDIO 1 -- WITH COOLEDIT 2000"
      End
      Begin VB.Menu Mnu_Find_New_Sounds 
         Caption         =   "Refresh - Find New In Folder For AUDIO 1"
      End
      Begin VB.Menu Mnu_Reset_MMControl 
         Caption         =   "Reset AUDIO All Media Control in Case All Drive Handles were Unlocked"
      End
      Begin VB.Menu Mnu_Reload_Audio 
         Caption         =   "Reload Usual Selection From Audio"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Explorer_Sound_1 
         Caption         =   "Explorer AUDIO Folder 1"
      End
      Begin VB.Menu Mnu_Explorer_Sound_2 
         Caption         =   "Explorer AUDIO Folder 2"
      End
      Begin VB.Menu Mnu_Play_Sound_1 
         Caption         =   "Play Current AUDIO 1"
      End
      Begin VB.Menu Mnu_Play_Sound_2 
         Caption         =   "Play Current AUDIO 2"
      End
      Begin VB.Menu MNU_SOUND_OPTION 
         Caption         =   "MNU_SOUND_OPTION"
         Index           =   1
      End
      Begin VB.Menu MNU_SOUND_OPTION 
         Caption         =   "MNU_SOUND_OPTION"
         Index           =   2
      End
      Begin VB.Menu MNU_SOUND_OPTION 
         Caption         =   "MNU_SOUND_OPTION"
         Index           =   3
      End
      Begin VB.Menu MNU_SOUND_OPTION 
         Caption         =   "MNU_SOUND_OPTION"
         Index           =   4
      End
      Begin VB.Menu MNU_SOUND_OPTION 
         Caption         =   "MNU_SOUND_OPTION"
         Index           =   5
      End
      Begin VB.Menu MNU_SOUND_OPTION 
         Caption         =   "MNU_SOUND_OPTION"
         Index           =   6
      End
      Begin VB.Menu MNU_SOUND_2 
         Caption         =   "MNU_SOUND_2"
      End
   End
   Begin VB.Menu MNU_INFO 
      Caption         =   "* INFO MENU *"
      Begin VB.Menu Mnu_APP_BEGIN_TIME_TIMER 
         Caption         =   "APP BEGIN TIME"
      End
      Begin VB.Menu MNU_BRing_Front 
         Caption         =   "Bring All Windows Front"
      End
      Begin VB.Menu MNU_TIME_API_FUNCTION_ACCESS 
         Caption         =   "TIME API SUB FUCTION LAST ACCESSED -- NOT A TIME YET"
      End
      Begin VB.Menu MNU_CLIPBOARD_API_PUBLIC_VAR_HOOK 
         Caption         =   "FORM CLIPBOAD API - PUBLIC VAR = TRUE SHOWS LOADED IS = "
         Visible         =   0   'False
      End
      Begin VB.Menu MNU_EXECUTE_TIMER_ENABLED 
         Caption         =   "EXECUTE_TIMER_ENABLED - TRUE OR FALSE"
      End
      Begin VB.Menu Mnu_Missing_Link_API_Test 
         Caption         =   "Missing Link Detector Check Clipboard API Unloaded = "
      End
      Begin VB.Menu MNU_MIRROR_EXE_DRY_RUN 
         Caption         =   "FORCE DRY RUN DRILL DOWN -- OF RUN MIRROR EXE ROUTINE"
      End
      Begin VB.Menu Mnu_Project_Source_Code 
         Caption         =   "Check For More Up-to-date Project Source Code"
      End
      Begin VB.Menu Mnu_Project_Source_Code_MIRROR 
         Caption         =   "Check For More Up-to-date Project Source Code MIRROR"
      End
      Begin VB.Menu Mnu_API_Reload_Status 
         Caption         =   "The API Form Clipper Logger Sub Call Loaded @ "
      End
      Begin VB.Menu Mnu_API_UNLoad_Status 
         Caption         =   "The API Form Clipper Logger Sub Call UN-Loaded @ "
      End
      Begin VB.Menu Mnu_API_Unload_Reload 
         Caption         =   "Test Unload && Reload Quick Test the API Clipboard Form"
      End
      Begin VB.Menu Mnu_API_Unload 
         Caption         =   "Test Unload the API Clipboard Form"
      End
      Begin VB.Menu Mnu_API_Reload 
         Caption         =   "Test Reload the API Clipboard Form"
      End
      Begin VB.Menu MNU_CLIPBOARD_TEST 
         Caption         =   "Test Code to Test This Program Is Working ClipBoarding Through the API -- Trigger MsgBox Result Happen"
      End
      Begin VB.Menu MNU_MESSAGE_BOXES 
         Caption         =   "STOP REPEAT NAGGING MESSAGE BOXES - TOGGLE"
      End
   End
   Begin VB.Menu Mnu_Loggs 
      Caption         =   "* LOGG MENU *"
      Begin VB.Menu Mnu_LoggExplorer 
         Caption         =   "Open Logg Explorer"
      End
      Begin VB.Menu MNU__ML_11 
         Caption         =   "----"
      End
      Begin VB.Menu MNU_INFRO_RAPID_ALL 
         Caption         =   "INFO RAPID ALL FOLDERS"
      End
      Begin VB.Menu MNU_INFRO_RAPID_ALL_SMALL_FILES 
         Caption         =   "INFO RAPID ALL FOLDERS AND SMALL FILES BEGIN ClipBoard-*.TXT"
      End
      Begin VB.Menu MNU_INFO_RAPID_MYUSER 
         Caption         =   "INFO RAPID MY USER"
      End
      Begin VB.Menu MNU_INFO_RAPID 
         Caption         =   "INFO RAPID TRIM LOGG"
      End
      Begin VB.Menu MNU__ML_12 
         Caption         =   "----"
      End
      Begin VB.Menu Mnu_Open_Recent_Hex 
         Caption         =   "Hex Open Recent Trim Logg"
      End
      Begin VB.Menu MNU__ML_13 
         Caption         =   "----"
      End
      Begin VB.Menu Mnu_Open_Recent 
         Caption         =   "Edit Recent Trim Logg"
      End
      Begin VB.Menu Mnu_Open_Logg 
         Caption         =   "Edit This Logg"
      End
      Begin VB.Menu Mnu_Open_Total 
         Caption         =   "Edit Total Logg"
      End
      Begin VB.Menu Mnu_StripLogg 
         Caption         =   "Edit Strip Logg"
      End
      Begin VB.Menu MNU__ML_14 
         Caption         =   "----"
      End
      Begin VB.Menu Mnu_TEXT_Open_Recent 
         Caption         =   "Text View Recent Trim Logg"
      End
      Begin VB.Menu Mnu_TEXT_Open_Logg 
         Caption         =   "Text View This Logg"
      End
      Begin VB.Menu Mnu_TEXT_Open_Total 
         Caption         =   "Text View Total Logg"
      End
      Begin VB.Menu Mnu_TEXT_StripLogg 
         Caption         =   "Text View Strip Logg"
      End
      Begin VB.Menu MNU__ML_15 
         Caption         =   "----"
      End
      Begin VB.Menu Mnu_ShellT_Todays 
         Caption         =   "Shell T -- This Logg"
      End
      Begin VB.Menu Mnu_Shell_T 
         Caption         =   "Shell T -- Total Logg"
      End
      Begin VB.Menu MNU__ML_17 
         Caption         =   "----"
      End
      Begin VB.Menu Mnu_Clips 
         Caption         =   "Clips"
         Begin VB.Menu Mnu_ClipAll 
            Caption         =   "Clip All"
         End
         Begin VB.Menu Mnu_ClearClipBoard 
            Caption         =   "Clear ClipBoard"
         End
         Begin VB.Menu Mnu_Clear_Text 
            Caption         =   "Clear Text"
         End
         Begin VB.Menu Mnu_Test_Clip 
            Caption         =   "Test Clip"
         End
      End
   End
   Begin VB.Menu MNU_LAST_ART_PIC2 
      Caption         =   "* IMAGE MENU *"
      Begin VB.Menu MNU_SCREEN_SHOT 
         Caption         =   "OPEN ART LOGG FOLDER"
      End
      Begin VB.Menu MNU__LAP_1 
         Caption         =   "----"
      End
      Begin VB.Menu MNU_LAST_ART_PIC_SELECTOR_EXPLORER 
         Caption         =   "EXPLORER SELECTOR __ LAST CLIPBOARD ONE"
      End
      Begin VB.Menu MNU_LAST_ART_PIC 
         Caption         =   "EXPLORER LAST CLIP-BOARDED HOT KEY IMAGE"
      End
      Begin VB.Menu MNU_LAST_ART_PIC_IVIEW 
         Caption         =   "IVIEW ---- LAST CLIP-BOARDED HOT KEY IMAGE"
      End
      Begin VB.Menu MNU__LAP_2 
         Caption         =   "----"
      End
      Begin VB.Menu Mnu_Explorer_Screen_Capture 
         Caption         =   "Explorer Auto Screen Capture"
      End
      Begin VB.Menu Mnu_IVIEW_Screen_Capture 
         Caption         =   "IVIEW - Auto Screen Capture"
      End
      Begin VB.Menu MNU__LAP_3 
         Caption         =   "----"
      End
      Begin VB.Menu Mnu_Explorer_Form_Capture 
         Caption         =   "Explorer Auto Current Form Capture"
      End
      Begin VB.Menu Mnu_IVIEW_Form_Capture 
         Caption         =   "IVIEW - Auto Current Form Capture"
      End
      Begin VB.Menu MNU__LAP_4 
         Caption         =   "----"
      End
      Begin VB.Menu MNU_LAST_WEBCAM_PIC 
         Caption         =   "EXPLORER LAST WEB CAM"
      End
      Begin VB.Menu MNU_IVIEW_LAST_WEBCAM_PIC 
         Caption         =   "IVIEW ---  LAST WEB CAM"
      End
      Begin VB.Menu MNU___LAP_5 
         Caption         =   "----"
      End
      Begin VB.Menu MNU_URL_SCREEN_SCRAPER 
         Caption         =   "EXPLORER LAST SCREEN SCRAPER FROM URL SHOT LOGGER"
      End
   End
   Begin VB.Menu MNU_SHOW_IMAGE 
      Caption         =   "* /\/\ SHOW IMAGE /\/\ *"
   End
   Begin VB.Menu MNU_FILE_LOCATOR_IMAGE 
      Caption         =   "FILE_LOCATOR.EXE _IMAGE_"
   End
   Begin VB.Menu MNU_SCANPATH_COUNTER 
      Caption         =   "Image Counter Numeric Info"
   End
   Begin VB.Menu MNU_F_O_M 
      Caption         =   "* FILE OPERATIONS MENU *"
      Begin VB.Menu MNU_RENAME_MP3_TAG_MUSIC_FOLDER 
         Caption         =   "RENAME SOME MP3 TAG IN MUSIC FOLEDR"
      End
   End
   Begin VB.Menu MNU_FORMAT_TEXT 
      Caption         =   "* FORMAT TEXT MENU >> *"
      Begin VB.Menu MNU_LAST_GRAB_ALL_CAPS 
         Caption         =   "LAST GRAB - ALL CAPS - ON ALL CLIPBOARD'S IF ENABLED"
      End
      Begin VB.Menu MNU_LAST_GRAB_CAPS 
         Caption         =   "LAST GRAB - CAPS - ONCE"
      End
      Begin VB.Menu MNU_PROCAPS 
         Caption         =   "LAST GRAB - PRO - CAPS"
      End
      Begin VB.Menu MNU_FORMAT_MINE 
         Caption         =   "SORT MINE PRO CAPS + MORE PROPER"
      End
      Begin VB.Menu MNU_ENTER_LARGE 
         Caption         =   "ENTER LARGE TEXT INTO LOGGER - Hardcoded - BluetoothLogView - To Allow"
      End
      Begin VB.Menu MNU_FORMAT_PLAIN_TEXT 
         Caption         =   "REFORMAT TO PLAIN TEXT"
      End
      Begin VB.Menu MNU_FORMAT_PLAIN_TEXT_LARGE_CLIPBOARD 
         Caption         =   "REFORMAT TO PLAIN TEXT - LARGE CLIPBOARD"
      End
      Begin VB.Menu MNU_SPACE 
         Caption         =   "ClipBoard a SPACE"
      End
      Begin VB.Menu MNU_REFORMAT_ADD_A_DASH 
         Caption         =   "ADD A DOUBLE DASH BEFORE EVERY CR-LINEFEED - TEXTBOXES THE REMOVE CR-LINEFEED  PROBLEM"
      End
      Begin VB.Menu MNU_REFORMAT_REMOVE_THE_DASH 
         Caption         =   "REMOVE THE  DOUBLE DASH BEFORE EVERY CR-LINEFEED"
      End
      Begin VB.Menu mnu_Keyboard_Logger_Remove_Doubel_Char_Into_One 
         Caption         =   "Keyboard Logger -- Remove Doudle Char Into One"
      End
   End
   Begin VB.Menu MNU_LAST_GRAB_LOW_CASE_TOP_BAR 
      Caption         =   "<-- TXT LOW"
   End
   Begin VB.Menu MNU_LAST_GRAB_CAPS_TOP_BAR 
      Caption         =   "TXT CAPS"
   End
   Begin VB.Menu MNU_PROCAPS_TOP_BAR 
      Caption         =   "TXT PRO CAPS"
   End
   Begin VB.Menu MNU_PROPER_CAPS_TOP_BAR 
      Caption         =   "TXT PROPER CAPS"
   End
   Begin VB.Menu MNU_FORMAT_PLAIN_TEXT_TOP_BAR 
      Caption         =   "TXT PLAIN"
   End
   Begin VB.Menu MNU_TEXT_HEX_2_TEXT 
      Caption         =   "TXT HEX 2 TEXT"
   End
   Begin VB.Menu MNU_CLIP_TO_FILE 
      Caption         =   "*  CLIP TO -- FILE _*MENU*_ *"
      Begin VB.Menu MNU_CLIP_TO_FILE_CODE 
         Caption         =   "CLIP TO TEMPORAY FILE AND EDITOR -"
      End
      Begin VB.Menu MNU_CLIP_TO_FILE_LAST 
         Caption         =   "LAST - CLIPPER TO TEMPORAY FILE -"
      End
   End
   Begin VB.Menu MNU_SCROLL_DOWN 
      Caption         =   "SCROLL DOWN"
      Visible         =   0   'False
      Begin VB.Menu MNU_800 
         Caption         =   "800"
      End
      Begin VB.Menu MNU_500 
         Caption         =   "500"
      End
      Begin VB.Menu MNU_400 
         Caption         =   "400"
      End
      Begin VB.Menu MNU_300 
         Caption         =   "300"
      End
      Begin VB.Menu MNU_200 
         Caption         =   "200"
      End
      Begin VB.Menu MNU_100 
         Caption         =   "100"
      End
      Begin VB.Menu MNU_SCROLL_OFF 
         Caption         =   "OFF"
      End
   End
   Begin VB.Menu Mnu_Clipboard_Info_Spacer 
      Caption         =   "ClipBoard INFO"
   End
   Begin VB.Menu Mnu_LAST_CLIP_TIME 
      Caption         =   "Last Clip Time"
   End
   Begin VB.Menu Mnu_LAST_CLIP_TIME_Second 
      Caption         =   "Last Clip Time Second"
   End
   Begin VB.Menu MNU_CB_SIZE 
      Caption         =   "CB SIZE 0.0 MEG  && Options"
      Begin VB.Menu Mnu_CB 
         Caption         =   "Hitt Menu The Option Below to Copy and Paste"
      End
      Begin VB.Menu Mnu_Line_Spacer 
         Caption         =   "------------------------------------------"
      End
      Begin VB.Menu MNU_CB_SIZE_MEG 
         Caption         =   "CB SIZE MEG"
      End
      Begin VB.Menu MNU_CB_SIZE_BYTE 
         Caption         =   "CB SIZE BYTE"
      End
   End
   Begin VB.Menu MNU_CB_SIZE_BYTE_OVERSIZE 
      Caption         =   "CB SIZE OVERSIZE LIMIT"
   End
   Begin VB.Menu Mnu_Clip_Description 
      Caption         =   "Clip Description"
   End
   Begin VB.Menu Mnu_Clip_Status 
      Caption         =   "Clip Status"
   End
   Begin VB.Menu Mnu_Run_Status 
      Caption         =   "Run Status"
   End
   Begin VB.Menu MNU_IDLE_ACTIVE 
      Caption         =   "IDLE <> ACTIVE"
   End
   Begin VB.Menu Mnu_Last_Clipboard_Timer 
      Caption         =   "Last Clipboard Timer"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_FILE_STUCK_IN_USE 
      Caption         =   "FILE STUCK IN USE"
   End
   Begin VB.Menu Mnu_URL 
      Caption         =   "* URL / FILE MENU *"
      Begin VB.Menu MNU_CLIPBOARD_EXPLORER_FILE_FOLDER 
         Caption         =   "OPEN FILE FOLDER PATH CLIPBOARD IN EXPLORER /SELECT"
      End
      Begin VB.Menu MNU_URL_21 
         Caption         =   "----"
      End
      Begin VB.Menu Mnu_URL_Browser 
         Caption         =   "URL Launch In Browser"
      End
      Begin VB.Menu MNU_URL_22 
         Caption         =   "----"
      End
      Begin VB.Menu Mnu_HTML_URL 
         Caption         =   "URL ADD WRAPPER  HTML <a href=""""..."
      End
      Begin VB.Menu MNU_URL_23 
         Caption         =   "----"
      End
      Begin VB.Menu MNU_CPC 
         Caption         =   "URL CPC WEB Page Offer Search 100 Web Pages 100%"
      End
   End
   Begin VB.Menu MNU_CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY 
      Caption         =   "URL_WEB_FOLDER_FILE_REG_KEY_CPC"
   End
   Begin VB.Menu MNU_CPC_TOP_MENU 
      Caption         =   "* CPC 20 80 90 *"
   End
   Begin VB.Menu MNU_404_CPC_PAGES 
      Caption         =   "* CPC 404 PAGE REMOVE *"
   End
   Begin VB.Menu MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS 
      Caption         =   "*--- NOT SCROLL IN  OPTION MENU---*"
   End
   Begin VB.Menu MNU_RUN_TIME 
      Caption         =   "RUN TIME"
   End
   Begin VB.Menu Mnu_Show_Error_Script_Frm_Msgbox 
      Caption         =   "** Show Error Script **"
   End
   Begin VB.Menu MNU_ERROR_INFO 
      Caption         =   "ANY ERROR MSG OR INFO"
   End
   Begin VB.Menu MNU_EBAY 
      Caption         =   "* EBAY GOLD TELEPHONE # * "
      Begin VB.Menu Mnu_Line_Spacer2 
         Caption         =   "--------------------"
      End
      Begin VB.Menu MNU_TITLE_EBAY_MENU_TAB 
         Caption         =   "EBAY -- FIND A GOLD TELEPHONE NUMBER -- NOT FINISHED YET"
      End
      Begin VB.Menu MNU_EBAY_NOTEPAD 
         Caption         =   "EBAY -- BUY FILTER RESULTS -- LOAD NOTEPAD++"
      End
      Begin VB.Menu MNU_EBAY_BEGIN_FILTER_SAVE_EDITOR 
         Caption         =   "EBAY -- FILTER RESULT OF BASIC LINE TEXT SAVE AND LOAD EDITOR"
      End
      Begin VB.Menu MNU_EBAY_FILTER_WILDCARD 
         Caption         =   "EBAY -- FILTER LINE INPUT - * - WILDCARD"
      End
      Begin VB.Menu MNU_EBAY_FILTER_NUMERIC_SIX 
         Caption         =   "EBAY -- FILTER LINE INPUT - NUMERIC - SIX "
      End
      Begin VB.Menu MNU_EBAY2 
         Caption         =   "EBAY_MORE"
      End
   End
   Begin VB.Menu Mnu_Audio_ON 
      Caption         =   "** /\/\ Audio Is Off /\/\ ** Hitt On Here for ON ***"
   End
   Begin VB.Menu MNU_SELECTED 
      Caption         =   "SELECTED"
   End
   Begin VB.Menu MNU_COPY 
      Caption         =   "<*** COPY ***>"
   End
   Begin VB.Menu MNU_CLIPPER_REMOVE_DOUBLE_VBCRLF 
      Caption         =   "REMOVE DOUBLE VBCRLF FROM CLIPPER"
   End
   Begin VB.Menu MNU_CLIPPER_WRAPPER_HTML_LINK_HREF 
      Caption         =   "WRAP ALL HTML LINK IN A WRAPPER <HREF>"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim DIE()


Dim O_Text1_SelLength
Dim LAST_CLIPBOARD_FILE

Const FULL_LINE = "RUN FULL LINE SELECT"

Dim PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY As String
Dim PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO___ As String
Dim PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP As String
Dim PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA As String
Dim PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA As String
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'THE PROJECT WAS DOWN FORM AUGUST 2KSIXTEEN UNTIL AROUND MARCH RECENT
'THE MENU GOT CURPUTED FROM A COMPARE SPLICE TOGETHER
'WAS UNTIL I WANTED A SNIPPET LEARN FROM I FIXED IT UP A BIT
'AND ONLY TODAY GOT A COMPILER GOING WITH ALLOW MAKE CHANGE
'1ST THE MENU WAS OVER 240 ITEM MAXIMUM
'2ND TODAY TWO ITEM APPEAR TO BE THE SAME
'AS RUN TROUGHT REMOVE SOME
'AS ERROR REPORT HAD NOT ALLOW ADD ANY MORE MAXIMUM
'Mon 10 April 2017 17:01:58----------
'TODAY ALSO THE DAY BEGIN MY LEARNING VBS SCRIPTER
'AND MADE THE LOADER FRONT END THING
'TO BRING THE MIRROR EXE OVER
'AND THEN RUN EASYIER
'ORIGINALY 1ST WORKED ON
'AND ALSO A VBS SCRIPT BEFORE TO DELETE MP3 CREATED ON BULK
'AT ---- KAT MP3 PROGRAM
'ALSO BOTH THE KAT MP3 PROGRAM
'--------------------------------------------------------------------
'--------------------------------------------------------------------







'WORK LEFT TO DO AT
'BOOKMARK HERE

'FS 'NEXT WORK AUGUST
'1.. SHIFT AND CURSOR CHUNCK MOVE UP DOWN AT A TIME
'2.. SEARCH BOX REQUIRED
'3.. EDIT BOX REQUIRED
'4.. WINMERGE COMPARE LAST TWO CLIPBOARD
'5.. IDLE ACTIVE LOGGING

'MR AND MRS
'MARRIED
'DR MARRIED FEMALE OR MALE
'CUSTOMER SERVICE -- DEAR SIR MADAM -- WHO ARE THEY -- HE SHE
'Actress Said to the Bishop
'

'NEXT WORK STARTED ALREADY TO DO

'LOT MORE OF THESE
'    EXECUTE_TIMER_ENABLED = False
'    Clipboard.Clear
'    Clipboard.SetText AD$
'    EXECUTE_TIMER_ENABLED = True

'NEED FINSIH MESSAGE LIST BOX RATHER THAN MSGBOX-ING

'---------------------------------
'Tue 05 July 2016 04:46:34
'NEW SPEEDER FOR FORM DESIGNER
'Check Disable Desktop Composition
'----
'[RESOLVED] VB6 build slow on Win 7-VBForums
'http://www.vbforums.com/showthread.php?712373-RESOLVED-VB6-build-slow-on-Win-7
'----
'---------------------------------




'Private Sub Mnu_Clip_Status_Click()
'Mnu_Clip_Status
'Private Sub MNU_CLIP_TO_FILE_CODE_Click()

'--------------------------------------------
'JOB # 01
'WORK TO FINISH ON THE SEND TO SPECIAL FOLDER
'2016 MAY 02 MON BANK HOLIDAY
'PROBLEM WITH MY LINKS WITH SENT TO IN WIN 10 - DON'T SEEM TO WANT TO WORK THEM LINKS
'--------------------------------------------
'JOB # 02
'WORK TO FINISH ON THE RESCAN TWICE OF AUDIO EFFECT AND IT LOSE ABILTY METHOD
'2016 MAY 02 MON BANK HOLIDAY
'-------------------------------
''SUB SETUP_SOUND_FILE(VARSOUND)
'-------------------------------
''WORK TO DO HERE
''LAST BREAK POINT WHILE WORK WAS SET HERE
'--------------------------------------------
'--------------------------------------------
'JOB # 03
'ADD SUPPORT FOR NETWORK CLIPBOARD SHARE
'2016 MAY 02 MON BANK HOLIDAY
'--------------------------------------------
'--------------------------------------------
'JOB # 04
'ADD IMPROVE WITH A SEARCH BOX
'2015
'--------------------------------------------
'--------------------------------------------
'JOB # 04
'MORE TO DO SORT MIRROR VERSION BETTER -- NOT SEEN IT WORKING MUCH
'2016 MAY 02 MON BANK HOLIDAY
'--------------------------------------------
'--------------------------------------------
'JOB # 05
'MORE INFO ON WHAT WAS STORED ABOUT CLIPBOARD HOLDING WHEN AT LOAD TIME
'2016 MAY 02 MON BANK HOLIDAY
'--------------------------------------------
'--------------------------------------------
'JOB # 07
'ADD FORWARD AND BACK THROUGHT LAST CLIPBOARD
'AT LEAST GET TO TOP OF LAST ONE WHEN BIG
'2016 MAY 31 DAY AFTER BANK HOLIDAY
'--------------------------------------------

'---------------------------------------------------------------
'WED 01 JUNE
'---------------------------------------------------------------
'JOB
'IS SOMEWHERE THAT REQUIRE A FORM BOX QUESTION RATHER THAN
'MSGBOX -- THINK IT IS PHOTO CAMERA
'---------------------------------------------------------------
'JOB
'NEED A LOT MORE JOB OF START UP FOLDER AND COMMON ONES
'ESPECIALLY NEW INSTALL GET GOING
'---------------------------------------------------------------
'JOB
'CHOP END FROM DUPE CHECK STOP DOING BIG FILE JOB END
'UNTIL ERROR SORTED
'---------------------------------------------------------------



'--------------------
'JOBS
'--------------------

'--------------------
'WORK
'--------------------


'--------------------------------------------
'WORK DONE FRI 13 MAY 20SIXTEEN
'--------------------------------------------
'MENU ITEM TO SHOW WHEN OVERSIZE LIMITED IS HELF IN CLIPBOARD
'AND THEN THE MENU ITEM TO ALLOW LARGE TEXT IS SET AND DISPLAY
'ONE TOP MENU ITEM IS NOT SHOWN WHEN NOT NEED
'
'THE WIN 7 HAD PROBLEM UNLESS FINDPART LOAD VB IDE
'HAD NOT A EXTRA RIGHT SQUARE BRACKET AT END SEARCH STRING
'AND THEN MORE CODE TO LOAD ANYWAY OF MSGBOX QUESTION IF WANT IN
'OTHER FAIL MIGHT HAPPEN
'MIGHT REMOVE THE QEUSTION LATER
'--------------------------------------------
'TWO WORK JOB ' CAN'T REMEMBER ANY MORE
'TO 2:15 PM ON LESS QUICK COMPUTER
'WITHOUT CHAIR - STANDING
'NOT A BIG SCREEN YET
'--------------------------------------------
'TWO ORGINAL IMPROVEMENT TO MAKE
'--------------------------------------------
'MORE NIGHT WORK
'--------------------------------------------
'ADD INFO OF EXE MIRROR DATE
'--------------------------------------------------
'CLEAN UP ALL ERROR OF FILE LOCK WHEN DRIVE LOCK
'EXAMPLE
'LOAD THE IDE - RUN PROGRAM -- BREAK IT ANYWHERE TO
'STOP EXECUTION
'RUN CHECK DISK CHKDSK AND LOCK THE DRIVE
'WONT LOCK IF RUNNING AND ACCESS
'AND THEN CONTINUE EXECUTION AGAIN
'THAT WAY WILL FIND ALL LOCK DRIVE OPEN FILE BUG
'ABOUT THREE BUG FOUND
'----------------------------------------------------------
'WARNING SEEM TO BE STILL A BUG OF OPEN FILE ACCESS PROBLEM
'AT LOAD TIME -- CRASHING -- NEED SORT VITAL
'AND IT WIPE THE HISTORY TEXT WHEN HAPPEN
'CHECK AFTER SEE THIS RUN
'HARD TO FIX UNLESS RUN CHKDSK QUICK WITHIN CODE FORM_LOAD
'----------------------------------------------------------



'--------------------------------------------
'WORK DONE SAT 14 MAY
'EARLY HOUR
'WORK DO EXE COMPILED AGAIN AND EXE MIRROR FOLDER
'THE CHANGE THAT HAPPEN REPLACE IS NOT MINIMIZED AFTER
'OPTION TO FORCE THE EXE FROM MIRROR  BUT NOT WORK WELL AS OLDER VERSION CAN REPLACE
'--------------------------------------------
'DO WORK -- SPECIAL FOLDER FOR SEND TO -- TO WORK WIN XP WIN 07 AND WIN 10
'WIN 7 AND WIN 10 SAME MAJOR BUILD NUMBER VER SIX
'--------------------------------------------
'USE THEM WIN XP 7 AND 10 TO DO SCREEN SIZE -- GOOD
'AND SCREEN SIZE MIN NORM CENTER AND MAX -- GOOD
'--------------------------------------------

'--------------------------------------------
'WORK DONE SUN 15 MAY
'--------------------------------------------
'END WITH LOT MENU WORK
'-----------------------
'GIVE ME TIME HAS MSGBOX
'-----------------------
'ADD NOT ONLY URL TO BROWSER AND NOW ALSO FILE OR FOLDER TO EXPLORER
'-----------------------
'FIX IDLE TIMER ALL GOOD
'REMOVE KEYBOARD F5 FROM IDLE TIMER GOOD IDEA AS DEBUG RESUME HARDER
'LOT WORK IN CATCH MISSING CLIPBOARD EVERY 5 SEC AFTER IDLE AND RESUME FROM IDLE AND 1 MIN
'NOT DUPLICATE MESSAGE
'-----------------------
'NOT DIPLAY FINDWINOW FOR CLIPBOARD CURRENT NOT ALWAYS
'ERROR IN CODER
'-----------------------
'WORK NINE O SEVEN AM OCLOCK TO 11:47 AM SUNDAY
'LONGER THAN A HOUR ON THE TELEPHONE
'-----------------------
'BEOFRE WAS WORK DUPE-LICATE CRC CHECKER
'AND MANY CODE FROM ABOUT 6PM SATURDAY TO 2AM
'WITH
'GOODSYNC SETTING UP
'AND
'NEW BLUETOOTH CODE SOUND AUDIO
'




'DO A LOAD WITH HEX VIEW FOR CODES
'HERE
'C:\Program Files\XVI32\XVI32.exe

'IMPROVE SCREEN CENTERING WHEN BOTTOM TASKBAR UP TWO LEVEL
'LIKE RUNAS AND CIDRUNME

'MORE TO DO - WITH MSGBOX STAY UPS
'FORM_STAY_UP_WITH_MSGBOX = True
'itech = MsgBox(Format(Now, "DD-MM-YYYY -- HH:MM:SS ") + vbCrLf + "1 MINUTE IDLE CHECK -- Problem Clipboard Not Logging TEXT BY COMPARE VARIABLE" + vbCrLf + "Put the Test Flag *ON* TEST COPY PASTE - And THEN Test Options Menu Unload and Reload the API Form is THE Next Workaround OPtion", vbMsgBoxSetForeground)

'MORE TO DO - WITH THESE MSGBOX
'MSGBOX2 = "VB IDE ICON CHANGED PROPERTY PROPERTIE" ', vbMsgBoxSetForeground
'FRM_MSGBOX.Timer1.Enabled = True

'DO THIS -- MNU_INFRO_RAPID_ALL_Click()
'UPGRADE ALL PATHS TO HAVE APP.PATH
'APP.PATH
'APATHY

'FINISH 'AD_DATE = Now -- WHY

'ADD A LITTLE EDITOR BOX


'Public HOOK_CLIPBOARD_API_lOADED As Boolean

'THIS CODE REPLACE MSGBOX
'EXAMPLE
'MSGBOX2 = "INFO - " + VB_ICON_TEXT_VAR
'FRM_MSGBOX.Timer1.Enabled = True

Dim XVB_DATE

Public O_VAL_MINUTE_API
Dim oMNU_IDLE_ACTIVE

Dim COMPARE_EXE_1, COMPARE_EXE_2, COMPARE_EXE_3


Dim VAR_MNU_404_HITT_ONCE

Dim GO3

Dim I5()


Dim TEXTBOX1_SELECTION_CHANGE


Dim RCPC
Dim RESUME_GO_CPC


Dim TIMER_CODE_TEST(200)

Dim VB_DATE_REMOTE

Dim OIK2, IK2

'Public DONE_COUNT_FORM_CAPTURE
Dim FF_COUNT_FORM_CAPTURE1
Dim FF_COUNT_FORM_CAPTURE2
Dim FF_COUNT_FORM_CAPTURE3
Dim FF_COUNT_FORM_CAPTURE4


Public SCREEN_SHOT_VAR_MOUSE_KEY_ACTIVE
Public FF_FORM1, FF_FORM2, FF_FORM3, FF_FORM4


Dim X22

Dim PATH_FILE_NAME1
Dim PATH_FILE_NAME2

Dim SIMU_TEST

Dim KEYBOARD_OR_MOUSE_ACTIVE
Dim KEYBOARD_OR_MOUSE_ACTIVE_LATCH
Dim PICTURE_INFO_TEXT


Dim OMenu_Height

Public IRFANVIEW_PATH
Public IRFANVIEW_PATH2
Public IRFANVIEW_PATH3

Dim O_RESIZE
Dim VAR_FLAG_EXPLORER_LOAD_NOT

Dim Mouse_Keyboard_Active_Time_Before
Dim Mouse_Keyboard_Active_Time

Dim FOLDER_PIN_TO_START_MENU
Dim FOLDER_SENDTO_SYSTEM
Dim FOLDER_SENDTO_FAT32_STORE


Dim MNU_IDLE_ACTIVE_VAR
Dim Mnu_Exit_VAR

Dim FILENAME1 As String

Dim HOOKSTATold


Dim KASC_TRIGGER

Dim DONT_RESIZE_WHILE_SETUP
Dim DONT_RESIZE_RUN_ONCE_OR_NORM

Dim vPathSOUND2 As String, vFileSOUND2
Dim vPathSOUND1 As String, vFileSOUND1
Dim ADTEST$
Dim ADTEST_BEFORE$
Dim TIMER1_LAST_DATE As Date

Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
Const SW_NORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const SW_MAXIMIZE = 3
Const SW_SHOWNOACTIVATE = 4
Const SW_SHOW = 5
Const SW_MINIMIZE = 6
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA = 8
Const SW_RESTORE = 9
Const SW_SHOWDEFAULT = 10
Const SW_FORCEMINIMIZE = 11
Const SW_MAX = 11



'PID HAS TO BE -1 -- Process Id Return in PID
'Var - False if Not Exist or PID Remain -1
'----
'USAGE = ----
'PID = -1
'Var = cProcesses.GetEXEID(PID, App.Path + "\" + App.EXEName + ".exe")
'If PID <> -1 Then
'    Call Process_HIGH_PRIORITY_CLASS(PID)
'End If
'USAGE = ----
'----
Public cProcesses As New clsCnProc
Dim PID As Long

'Public Kill_Usage_Hogs_DATE1
'Dim Kill_Usage_Hogs2_CHECKGATE
'Dim Kill_Usage_Hogs3_CHECKGATE

Dim APPEXEDATE, ADATE_MIRROR_EXE_DATE
Dim ADATE_APP_BEGIN_DATE, ADATE_APP_BEGIN_DATE2
Dim WTrue1, WTrue2, WTrue3, TW1, TW2, TW3

Dim Timer_API_Test_Logick_Var1_Missing_Count
Dim Timer_API_Test_Logick_Var1_OLD
Dim Timer_API_Test_Logick_Var2_OLD

Dim FDS2
Dim FORM_STAY_UP_WITH_MSGBOX
Dim FirstRun

Dim LimitClipSize
'Dim TimerCheckIntegrityOfProgram
'Dim DateTimerCheckIntegrityOfProgram
Dim ClipFormatDescription
Dim XArchiveXClipFormatDescription
Dim FlagTestClipBoardRoutine

Dim ExplorerGone

Dim FOLDERNAME_AUTO
Dim OverRideOnceExtra
Dim FOLDERNAME1, FOLDERNAME2
Dim OX22, OX24 As Long, OX25 As Long
Dim DUPE_CLIPPER_AT_LOAD_FORM

Dim PICXX$, ADLEN

Dim Xmp5, Ymp5, MouseClicks, Mousey
Dim ADATE3_BEFORE
Dim ENTER_LARGE_IN_LOGGER

Dim OCheckQ As String

Dim iMessageResultRECompile, Pro2

Dim OLTLWH, RESET_RRR
Dim DUPE_IMAGE_AT_LOAD_FORM
'Dim RESIZE_AT_LOAD, RESIZE_AT_LOAD2

Public ART$, ART2$, GGSize

Public QQ2$, QQ4$
Public ARCHIVE_Path_Of_Sound_File

Public Path_Of_Sound_File, MODIFY_SOUND_SELECTION, LATCH_RUN_ONCE
Dim Path_Of_Sound()
Dim KCODE


Private FS, FS2

' Public Sub SET_VAR_FS()
' Set FS = CreateObject("Scripting.FileSystemObject")



Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer



Private Type POINTAPI
        x As Long
        y As Long
End Type


Dim OFH, A1 As String, B1 As String, f1 As String

Private Declare Function FlatSB_SetScrollPos Lib "comctl32" (ByVal hwnd As Long, ByVal code As Long, ByVal nPos As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32" (ByVal hwnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO) As Boolean
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32" (ByVal hwnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO, ByVal fRedraw As Boolean) As Long

Const SB_HORZ = 0
Const SB_VERT = 1
Const SB_BOTH = 3
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Const SIF_RANGE = &H1
Const SIF_PAGE = &H2
Const SIF_POS = &H4
Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS)


Public RRR, XXMOUSEMOVE, OLDTTS 'Picture3W, Picture3H,
Public Star, RR$, RrS$, CountR2$

Public AD$, TRemGG$
Public CountR
Public start
Public Pic1$, Pic2$
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetParent _
        Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Const GW_CHILD = 5

Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lClose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Dim Filexxx$, CurProcHwnd, TTT
Dim Rect1 As RECT

Dim TS1 As String, PATH_TO_CLIPBOARD_AND_FILENAME As String, TS3 As String, QQ, Txw$, ii, FF$, XX, YY
'Private Declare Function GetForegroundWindow Lib "user32" () As Long

'Public cProcesses As New clsCnProc
    '#### Functions/Consts used for GetHWnd() (part of Convert())
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Private Const GW_HWNDNEXT = 2
'Private Const GW_CHILD = 5
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long



Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Dim TJax, GJax, ET, TBak, TY, Tx, HH

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long

Private mCapHwnd As Long

Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054

'declarations
Dim P() As Long
Dim POn() As Boolean

Dim inten As Integer

Dim I As Integer, j As Integer

Dim Ri As Long, Wo As Long
Dim RealRi As Long

Dim C As Long, c2 As Long

Dim r As Integer, G As Integer, B As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

Dim Tppx As Single, Tppy As Single
Dim Tolerance As Integer

Dim RealMov As Integer

Dim Counter As Integer

Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim LastTime As Long, TT




Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CYCAPTION = 4
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'http://www.vbforums.com/showthread.php?673134-RESOLVED-Minimum-height-for-menu-bar-to-be-visible
'-------------- MENU SIZE DECLARE
'Private Type rect
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
Private Type MENUBARINFO
  cbSize As Long
  rcBar As RECT
  hMenu As Long
  hwndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hwnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
'Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long


Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long



Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'----

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long  'MODULE 1135

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40





'------------------------------------------------

Property Get TitleBarHeight() As Long
    TitleBarHeight = GetSystemMetrics(SM_CYCAPTION)
End Property



'On Error Resume Next
'MkDir App.Path + "\# DATA\"+GetComputerName+ "\Data\"+GetComputerName
'On Error GoTo 0



'Private Sub Command1_Click()
'
'Clipboard.Clear
'Clipboard.SetText Text1.Text
'
'End Sub
'
'Private Sub Command2_Click()
'Clipboard.Clear
'
'End Sub

'Private Sub Command3_Click()
'
'    Text1.Text = ""
'    CountR = 0
'
'End Sub

'Private Sub Command4_Click()
'Clipboard.Clear
'Clipboard.SetText Format$((Now), "DDD DD-MMM-YY HH:MM:SS")
'End Sub


Private Sub Form_Load()

FirstRun = True
HOOKSTATold = -10

MNU_SLECTOR_WHOLE_LINE_MODE.Caption = FULL_LINE + "=FALSE"


LimitClipSize = 200000
ADATE_APP_BEGIN_DATE = Now

'For r = 1 To 127

'    Debug.Print Str(r) + " -- " + GetSpecialfolder(r)

'Next

'Exit Sub


If App.PrevInstance = True Then End

Set FS = CreateObject("Scripting.FileSystemObject")


'Debug.Print "PROGRAM BEGIN__"



'PID HAS TO BE -1 -- Process Id Return in PID
'Var - False if Not Exist or PID Remain -1

If IsIDE = False Then
    PID = -1
    Var = cProcesses.GetEXEID(PID, App.Path + "\" + App.EXEName + ".exe")
    If PID <> -1 Then
        Call Process_HIGH_PRIORITY_CLASS(PID)
    End If
End If


'If IsIDE = False Then
    If Dir("C:\Program Files (x86)\PicPick\picpick.exe") <> "" Then
        PID = -1
        Var = cProcesses.GetEXEID(PID, "C:\Program Files (x86)\PicPick\picpick.exe")
        If PID = -1 Then
            Shell "C:\Program Files (x86)\PicPick\picpick.exe", vbNormalNoFocus
        End If
    End If
    If Dir("C:\Program Files\PicPick\picpick.exe") <> "" Then
        PID = -1
        Var = cProcesses.GetEXEID(PID, "C:\Program Files (x86)\PicPick\picpick.exe")
        If PID = -1 Then
            Shell "C:\Program Files\PicPick\picpick.exe", vbNormalNoFocus
        End If
    End If
'End If


Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
On Error Resume Next
EXECUTE_FILE_NAME = "E:\01 Start Menu\#_7-ASUS-GL522VW\Programs\Startup-01-Delayed\7-ASUS-GL522VW--MATT 01\Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk.lnk"
WSHShell.Run """" + EXECUTE_FILE_NAME + """"
EXECUTE_FILE_NAME = "E:\01 Start Menu\#_7-ASUS-GL522VW\Programs\Startup-01-Delayed\7-ASUS-GL522VW--MATT 01\Autokey -- 10-READ MOUSE CURSOR ICON STATE AND BEEPER WHEN NOT BUSY HOUR GLASS OVER.ahk.lnk"
WSHShell.Run """" + EXECUTE_FILE_NAME + """"
Set WSHShell = Nothing
On Error GoTo 0


If IsIDE = False Then
    PID = -1
    Var = cProcesses.GetEXEID(PID, "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe")
    If PID = -1 Then
        Shell "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe", vbNormalNoFocus
    End If
End If


MNU_SCANPATH_COUNTER.Visible = False

'Mnu_Minimize.Caption = "MIN"
'Mnu_Max.Caption = "MAX"
'Mnu_Norm.Caption = "NORM"

'Mnu_Height.Visible = False
'FrmJoy.Show

Xmp5 = -1: Ymp5 = -1

Call ScanPath.SET_VAR_FS
'Set FS = CreateObject("Scripting.FileSystemObject")

'Picture2.Picture = LoadPicture(App.Path + "\# DATA\"+GetComputerName+ "\VBIcon4.bmp")
    
'Me.Show
'DoEvents

IRFANVIEW_PATH3 = "C:\Program Files (X86)\IrfanView\i_view32.exe"
If Dir(IRFANVIEW_PATH3) <> "" Then IRFANVIEW_PATH = IRFANVIEW_PATH3
IRFANVIEW_PATH2 = "C:\Program Files (X86)\IrfanView\i_view32.exe"
If Dir(IRFANVIEW_PATH2) <> "" Then IRFANVIEW_PATH = IRFANVIEW_PATH2
    
    
    
    
'--------------------------------------------------------------------
Call SET_FOLDER_CLIPBOARD_LOGGER

FileName_er = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP + "\HotKey-Shot % Pic Data Location.txt"
If Dir(FileName_er) <> "" Then
    FR1 = FreeFile
    Open FileName_er For Input As #FR1
        Line Input #FR1, LAST_CLIPBOARD_FILE
    Close #FR1
End If
'--------------------------------------------------------------------
    
    
Dim FileSpec, TT As Long

'PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA

If FS.FolderExists(App.Path + "\Sound_Wav's") = False Then
    MkDir App.Path + "\Sound_Wav's"
End If

If FS.FolderExists(App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBDataNoTArchive") = False Then
    MkDir App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBDataNoTArchive"
End If
'--------------------------------------------------------------------

If FS.FolderExists("D:\# MY DOCS") = False Then
    MkDir "D:\# MY DOCS"
End If
If FS.FolderExists("D:\# MY DOCS\# 01 My Documents") = False Then
    MkDir "D:\# MY DOCS\# 01 My Documents"
End If
If FS.FolderExists("D:\# MY DOCS\# 01 My Documents\03 NotePad Files") = False Then
    MkDir "D:\# MY DOCS\# 01 My Documents\03 NotePad Files"
End If

'--------------------------------------------------------------------






'--------------------------------------------------------------------

On Error GoTo 0

'Not Any Need to Write Code as Call zzSave_Checks -- Does the Job
'Mnu_SoundOn.Checked = True

FileSpec = App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\" + App.EXEName + ".vbp"
If IsIDE = False And Dir$(FileSpec) <> "" Then
    'TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + FileSpec + """", vbMinimizedNoFocus)
    'End
End If
    
If Dir(App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBIcon4.bmp") <> "" Then
    FR1 = FreeFile
    Open App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBIcon4.bmp" For Binary As #1
        'Open App.Path + "\VBIcon4.bmp" For Binary As #FR1
        'Open App.Path + "\VBIcon.bmp" For Binary As #1
        Pic1$ = Space$(LOF(FR1))
        Get #FR1, 1, Pic1$
    Close #FR1
End If

Open PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total--TRIM.txt" For Binary As #1
If LOF(1) > 1 * 1024 ^ 2 Then
    Pic12$ = Space$((1 * 1024 ^ 2) + 1)
    Seek 1, LOF(1) - (1 * 1024 ^ 2)
    Get #1, , Pic12$
    Close #1
    
    '---------------------
    'Count =
    
    'ii = InStr(Pic12$, "---------------------" + vbLf + "Count =")
    ii = InStr(Pic12$, "---------------------" + vbcflf + "Count =")
    If ii = 0 Then
        ii = InStr(Pic12$, "---------------------" + vbLf + "Count =")
    End If
    If ii = 0 Then
        ii = 1
    End If
    
    Pic12$ = Mid$(Pic12$, ii)
    
    Kill PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total--TRIM.txt"
    Open PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total--TRIM.txt" For Binary As #1
        Put #1, , Pic12$
    Close #1
    
    'Simple Copy File

    A1 = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total--TRIM.txt"
    B1 = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\00_ClipBoard_Total--TRIM-" + GetComputerName + "-" + GetUserName + ".txt"
    'fs.CopyFile ,
    ShellFileCopy A1, B1

End If
Close #1

'------------
FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#OutClipChunck.Txt"
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Binary As #FR1
    AD$ = Input$(LOF(FR1), FR1)
Close #FR1


ADTEST$ = AD$
ADTEST_BEFORE$ = AD$
'------------

'COMPARE_EXE_2 = AD$

'MNU_COMPARE.Caption = "COMPARE" + Str(Len(COMPARE_EXE_2)) + " 0"

Call COMPARE_FOR_EXE


Call CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY(False)


DUPE_CLIPPER_AT_LOAD_FORM = False

If Clipboard.GetFormat(vbCFText) = True Then
    Mnu_Clip_Description.Caption = "Clip Format:- " + "Text (.txt file)"
    Mnu_Clip_Description.Caption = "Clip Format:- " + "Text" ' (.txt file)"
    
    'MEMORY TEXT AT STARTER
    '----------------------
    If Trim(Clipboard.GetText) <> "" Then
        If AD$ = Clipboard.GetText Then
            DUPE_CLIPPER_AT_LOAD_FORM = True
            Mnu_Clip_Status.Caption = "Text in Buffer at Starter is Duplicate"
        Else
            Mnu_Clip_Status.Caption = "Text in Buffer at Starter Status is Not Duplicate"
            'Delicate Duplicate
        
        End If
    End If
    GetFormat_And_Display
    xyz2020 = True
End If



Text1 = ""
On Error Resume Next
If Dir(PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\Start.txt") <> "" Then
    Open PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\Start.txt" For Input As #1
    Line Input #1, Star
    Close #1
End If
On Error GoTo 0

'If Star = "" Then Star = Str(Now - DateSerial(0, 0, 40))

start = Now
If Star = "" Then Star = Now

Mnu_Open_Logg.Caption = "Edit This Logg - " + PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA + "\ClipBoard-" + Format$(start, "YYYY-MM-DD") + ".Txt"

'#IF ITS THE SAME DAY
CountR2$ = ""
If DateValue(Star) = DateValue(start) Then
    FR1 = FreeFile
    Open PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#ClipBoard.Txt" For Binary As #FR1
        RrS$ = Input$(LOF(FR1), FR1)
    Close #FR1
    Err.Clear
    On Error Resume Next
        Text1.Text = Text1.Text + RrS$
        If Err.Number > 0 Then
            Me.WindowState = vbNormal
            DoEvents
            MsgBox "Not Adding to CLip Board"
        End If
    On Error GoTo 0
    
    FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#Count.Txt"
    FR1 = FreeFile
    If Dir(FILENAME_IN_USE_CHECK) <> "" Then
        Open FILENAME_IN_USE_CHECK For Input As #FR1
        If Not (EOF(FR1)) Then Line Input #FR1, CountR2$
        Close #FR1
    End If
    CountR = Val(CountR2$)

Else
    tq = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#ClipBoard_Old.Txt"
    If Dir(tq) <> "" Then
        Kill tq
        If Dir(PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#ClipBoard.Txt") <> "" Then
            Name PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#ClipBoard.Txt" As PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#ClipBoard_Old.Txt"
        Else
            Open PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#ClipBoard.Txt" For Output As #FR1
            Close #FR1
        End If
    End If
    CountR = 0
End If

FR1 = FreeFile
Open PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\Start.txt" For Output As #FR1
    Print #FR1, Format$(start, "DD-MM-YYYY")
Close #FR1

'After load most import recent load a backlogg for viewing
FR1 = FreeFile
Open PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#ClipBoard_Old.Txt" For Binary As #FR1
    RrS$ = Input$(LOF(FR1), FR1)
Close #FR1
Err.Clear
On Error Resume Next
    Text1.Text = RrS$ + Text1.Text
    If Err.Number > 0 Then
        Me.WindowState = vbNormal
        DoEvents
    
        MsgBox "Not Adding to CLip Board"
    End If
On Error GoTo 0

Dim COMPARE1 As String, COMPARE2 As String
'PICTURE COMPARE
DUPE_IMAGE_AT_LOAD_FORM = False

If Clipboard.GetFormat(vbCFBitmap) = True Then
    Picture3.Picture = Clipboard.GetData(vbCFBitmap)
    SavePicture Picture3.Picture, App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBDataNoTArchive\HotKey-Shot Pic-TEST.jpg"
    'LAST_CLIPBOARD_FILE = TS1

    If Dir(App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBDataNoTArchive\HotKey-Shot Pic-TEST.jpg") <> "" Then
        FR1 = FreeFile
        Open App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBDataNoTArchive\HotKey-Shot Pic-TEST.jpg" For Binary As #FR1
            COMPARE1 = Space$(LOF(FR1))
            Get #FR1, 1, COMPARE1
        Close #FR1
        Kill App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBDataNoTArchive\HotKey-Shot Pic-TEST.jpg"
        'ANOTHER COMPARE IN MAIN ROUTINE
        PICXX$ = COMPARE1
    End If

    FILENAME1 = App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBDataNoTArchive\COMPARE DUPE.txt"
    
    'If Dir(App.Path + "\# DATA\" + GetComputerName +"_"+GETUSERNAME+ "\VBDataNoTArchive\HotKey-Shot Pic.jpg") <> "" Then
    If Dir(FILENAME1) <> "" Then
        
        FR1 = FreeFile
        Open FILENAME1 For Binary As #FR1
            COMPARE2 = Space$(LOF(FR1))
            Get #FR1, 1, COMPARE2
        Close #FR1
        
        If COMPARE1 = COMPARE2 Then
            Call Menu_clipboard_size(Len(COMPARE1))
                If FirstRun = True Then
                    'FirstRun = False
                    'Mnu_LAST_CLIP_TIME.Caption = Mnu_LAST_CLIP_TIME.Caption + " First Run"
                    Mnu_Run_Status.Caption = "1st Run"

                End If

            COMPARE1 = "": COMPARE2 = ""
            DUPE_IMAGE_AT_LOAD_FORM = True
            DUPE_CLIPPER_AT_LOAD_FORM = True
            
            Mnu_Clip_Status.Caption = "Dupe Image at Starter"
        Else
            Mnu_Clip_Status.Caption = "Image at Starter"
        
        
        End If
        'Mnu_Clip_Description.Caption = "Clip Format:- " + "Bitmap (.bmp file)"
        GetFormat_And_Display
        xyz2020 = True
    
    End If
End If

If xyz2020 = False Then
    GetFormat_And_Display
    Mnu_Clip_Status.Caption = "At Starter"
End If

'Form1.Show

Call zzLoad_Checks
    
'Text1.SelStart = 0
'Text1.SelLength = Len(Text1)
'Text1.Font.Size = 12
'Text1.SelColor = &HFF00&
    
Call SETUP_SOUND_FILE("")

'Call MNU_Norm_Click

'Me.WindowState = vbMaximized
'vbMaximized
'vbNORML
'vbMinimized



'Form1.WindowState = vbMinimized
If GetComputerName = "5-ASUS-P2520LA" Then
    Text1.Font.Size = 14
Else
    Text1.Font.Size = 14
End If
Text1.SelStart = Len(Text1.Text)

Label1.Top = 0
Label1.Left = 0
Label1.Caption = ""
Label1.Height = 80
Label1.Left = 0

'Call MNU_Norm_Click

'Load FrmJoy

'Exit Sub

Load frmClipTest

'WE CAN SET THE EXECUTE TO TRUE AFTER _ NOTHING WILL CLIPBOARD AT FIRST RUN
Call Timer1_Timer

'--------------------------------------
' SET TWICE HERE AND LATER IN THIS FORM
'--------------------------------------
MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = False

'ctlClipboard1_ClipboardChanged
EXECUTE_TIMER_ENABLED = True

'RESIZE_AT_LOAD = True

'Call ctlClipboard.StartClipboardViewer
'Call StartClipboardViewer
'Debug.Print mStarted

'Form2.Show

Timer3.Enabled = True
Timer_SCREEN_SHOT.Enabled = True

'If IsIDE = True Then
'    'RESIZEATLOAD = False
'    'Call MNU_Norm_Click
'    Me.WindowState = vbNormal
'End If

'--------------------------------------------------------
'SET THE MENU CAPTION
'--------------------------------------------------------
'SEND TO
Call MNU_MIRROR_SEND_TO_OPERATING_SYSTEM_SET_MENU_CAPTION
'--------------------------------------------------------
'SEND TO
Call SPECIAL_FOLDER_SEND_TO
'--------------------------------------------------------
VAR_FLAG_EXPLORER_LOAD_NOT = True
'START MENU
Call MNU_SPECIAL_FOLDER_PIN_TO_START_MENU_Click
'DESKTOP
Call MNU_SPECIAL_FOLDER_DESKTOP_USER_Click
Call MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC_Click
'--------------------------------------------------------
'CHECK HERE - WINMERGE
'Call SPECIAL_FOLDER_SENT_TO
Call SPECIAL_FOLDER_PIN_TO_START_MENU
'--------------------------------------------------------

MNU_EXPLORER_ME_VB.Caption = "EXPLORER ME_VB -- " + App.Path

'FOUND CPU TYPE AND GETCOMPUTERNAME FROM MNUE ABOVE
'SAVE ON THAT
'------------
'If GetComputerName = "5-ASUS-P2520LA" Then
'TEXT1.Font.Size =
'NEED EARLIER AS SET BEFORE

Timer_API_Test.Enabled = True

DONT_RESIZE_RUN_ONCE_OR_NORM = True
DONT_RESIZE_WHILE_SETUP = True
'-----------------------------------------
Call SETUP_DIMENSIONS_NORM
Call SETUP_DIMENSIONS_INNER_FORM
'-----------------------------------------
'RESIZE_AT_LOAD = False
DONT_RESIZE_WHILE_SETUP = False

If Timer_API_Test_Logick_Var2 = 0 Then
    VARTEXT = "THE TIME THE API CLIPBOARD FUNCTION LAST ACCESSED = NOT YET"
    MNU_TIME_API_FUNCTION_ACCESS.Caption = VARTEXT
End If

If IsIDE = True Then
    Me.WindowState = vbNormal
Else
    Me.WindowState = vbMinimized
End If

'STARTUP TIME
ADATE_APP_BEGIN_DATE2 = DateDiff("S", ADATE_APP_BEGIN_DATE, Now)

MNU_IDLE_ACTIVE.Caption = "IDLE ><>< ACTIVE"

'If IsIDE = True Then Mnu_VB.Enabled = False

DoEvents

'Call MNU_TEST_STOP_ALL_TIMER_Click

If Mnu_SoundOn.Checked = False Then
    Mnu_Audio_ON.Visible = True
    'MNU_AUDIO_WANT_ON.Visible = True
    MNU_AUDIO_WANT_ON.Caption = "AUDIO WANT ON_?"
Else
    Mnu_Audio_ON.Visible = False
    'MNU_AUDIO_WANT_ON.Visible = False
    MNU_AUDIO_WANT_ON.Caption = "AUDIO=ON"

    
End If

'MNU_404_CPC_PAGES.Visible = False

If IsIDE = True Then
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = True
Else
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = False
End If

'If MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = False Then
'    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Visible = True
'End If

Call MNU_NIRSOFT_SETUP

Call RESET_SETUP_SOUND_FILE("NOTSOUND")
'Call Mnu_API_Unload_Click
'Call Mnu_API_Reload_Click

MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = True

If MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = True Then
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Caption = "*--- SCROLL BACK DOWN TO BOTTOM ON TIMER.ENABLED=TRUE ---*"
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Caption = "*--- SCROLL BACK DOWN TO BOTTOM ON TIMER.ENABLED=TRUE ---*"
    MNU_SCROLL_BOTTOM_OFF__WANT_ON_ASK_PRESS.Caption = "SCROLL BOTTOM=ON"
    'MNU_SCROLL_BOTTOM_OFF__WANT_ON_ASK_PRESS.Visible = False
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Visible = False
Else
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Caption = "*--- SCROLL BACK DOWN TO BOTTOM ON TIMER.ENABLED=FALSE ---*"
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Caption = "*--- SCROLL BACK DOWN TO BOTTOM ON TIMER.ENABLED=FALSE ---*"
    MNU_SCROLL_BOTTOM_OFF__WANT_ON_ASK_PRESS.Caption = "SCROLL BOTTOM=NOT"
    'MNU_SCROLL_BOTTOM_OFF__WANT_ON_ASK_PRESS.Visible = True
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Visible = False

End If

'------------------------------------
'IDLE ACTIVE STATE
'------------------------------------
Timer_MOUSE_1_MINUTE.Interval = 4000

'Text1.ScrollBars = rtfNone
'Text1.ScrollBars = rtfVertical
'READ ONLY WANTED TO CONTROL LINE AND WORD WRAP
'HAS TO BE SET WITHOUT -- Text1.ScrollBars = rtfVertical

Call MNU_AUDIO_WANT_ON_Click

End Sub


Private Sub MNU_ABORT_SHUTDOWN_Click()
    
    'CSIDL_WINDOWS_SYSTEM32 = 37
    Shell GetSpecialfolder(37) + "\shutdown.exe /a", vbNormalFocus

End Sub

Private Sub Mnu_API_Reload_Click()

On Error Resume Next
Load frmClipTest

End Sub

Public Sub MNU_API_RESET_Click()

Beep
'-----------------------------
O_VAL_MINUTE_API = Now

'Call Form1.MNU_API_RESET_SUB
'-----------------------------
'Call MNU_API_RESET_SUB
Call RESET_SETUP_SOUND_FILE("")

Call Mnu_API_Unload_Click
Call Mnu_API_Reload_Click

End Sub

Private Sub MNU_CLIPPER_REMOVE_DOUBLE_VBCRLF_Click()

Form1.RTB_CLIPPER_FORMAT_CONVERTING.Text = Clipboard.GetText
RTB_CLIPPER_FORMAT_CONVERTING.Text = Replace(RTB_CLIPPER_FORMAT_CONVERTING.Text, vbCrLf + vbCrLf, vbCrLf)
EXECUTE_TIMER_ENABLED = False
Clipboard.Clear
EXECUTE_TIMER_ENABLED = True
Clipboard.SetText RTB_CLIPPER_FORMAT_CONVERTING.Text
Me.WindowState = vbMinimized
RTB_CLIPPER_FORMAT_CONVERTING.Text = ""
Beep
'http://roidsrim.blogspot.co.uk/2015/06/better-living-through-lobotomy-what-can.html

End Sub

Private Sub MNU_CLIPPER_WRAPPER_HTML_LINK_HREF_Click()

RTB_CLIPPER_FORMAT_CONVERTING.Text = Clipboard.GetText
'RTB_CLIPPER_FORMAT_CONVERTING.Text = Replace(RTB_CLIPPER_FORMAT_CONVERTING.Text, "http://", "<a href=""http://")

Dim XPOS_, XPOS2
Dim XPOS_21, XPOS_22, XPOS_23
Dim TEXT_HTML_TAG_1
Dim TEXT_HTML_TAG_2

Dim I_4, I_5


XPOS_ = 1
Do
    'DoEvents
    '01 OF 02 --------
    XPOS_ = InStr(XPOS_, UCase(RTB_CLIPPER_FORMAT_CONVERTING.Text), "HTTP")
    If XPOS_ > 0 Then
        MID_1 = Mid(UCase(RTB_CLIPPER_FORMAT_CONVERTING.Text), XPOS_ - 5, Len("REF=""HTTP"))
        MID_2 = "REF=""HTTP"
        If MID_1 <> MID_2 Then
            XPOS_YES = XPOS_
            XPOS_ = XPOS_ + 1
        Else
            XPOS_YES = -1
            XPOS_ = XPOS_ + 1
        End If
    End If
    
    '02 OF 02 --------
    If XPOS_YES > 0 Then
        'XPOS_CHECK = InStr(XPOS_, UCase(RTB_CLIPPER_FORMAT_CONVERTING.Text), "/"">")
        MID_1 = Mid(UCase(RTB_CLIPPER_FORMAT_CONVERTING.Text), XPOS_YES - 5, Len("/"">"))
        MID_2 = "/"">"
        If MID_1 <> MID_2 Then
            XPOS_YES = XPOS_
        Else
            XPOS_YES = -1
        End If
    End If
    
    
    'If XPOS_ = 0 Then XPOS_ = InStrRev(RTB_CLIPPER_FORMAT_CONVERTING.Text, "WWW")
    'If XPOS_ > 0 Then
    '    If Mid(UCase(RTB_CLIPPER_FORMAT_CONVERTING.Text), XPOS_ - 5, Len("REF=""HTTP")) = "REF=""HTTP" Then
    '        XPOS_ = XPOS_
    '    Else
    '        XPOS_ = 0
    '    End If
    'End If
    'If XPOS_ > 0 Then XPOS2 = InStr(RTB_CLIPPER_FORMAT_CONVERTING.Text, "WWW")
    
    '------------------------------------------------
    If XPOS_YES > 0 Then XPOS_21 = InStr(XPOS_YES, RTB_CLIPPER_FORMAT_CONVERTING.Text, " ")
    '------------------------------------------------
    If XPOS_YES > 0 Then XPOS_22 = InStr(XPOS_YES, RTB_CLIPPER_FORMAT_CONVERTING.Text, vbCrLf)
    '------------------------------------------------
    If XPOS_YES > 0 Then XPOS_23 = InStr(XPOS_YES, RTB_CLIPPER_FORMAT_CONVERTING.Text, "<")
    '------------------------------------------------
    
    Dim OKAY
    If XPOS_YES > 0 Then
        If XPOS_21 < XPOS_22 Then
            OKAY = XPOS_21
        Else
            OKAY = XPOS_22
            XPOS_21 = XPOS_22
        End If
        
        If XPOS_21 < XPOS_23 Then
            OKAY = XPOS_21
        Else
            OKAY = XPOS_23
        End If
    End If
    
    If XPOS_YES > 0 Then
        XPOS_INDEX_POINTER = XPOS_YES - 50: IDX_POINTER = 50
        If XPOS_YES - 10 <= 0 Then XPOS_INDEX_POINTER = 1: IDX_POINTER = XPOS_YES
        TEXT_HTML_TAG_0 = Mid(RTB_CLIPPER_FORMAT_CONVERTING.Text, XPOS_INDEX_POINTER - 1, IDX_POINTER)
        TEXT_HTML_TAG_1 = Mid(RTB_CLIPPER_FORMAT_CONVERTING.Text, XPOS_YES - 1, (OKAY - XPOS_YES) + 1)
        '----------------------------------------------------------------------
        'GOT EVERYTHING BUT NOT WANT IF NOT CHECKING FOR A SLASH AT END ALREADY
        'USUAL IS GET ADDER A SLASH / WHEN MAKE A HREF WRAPPER
        '----------------------------------------------------------------------
        
        If Mid(TEXT_HTML_TAG_1, Len(TEXT_HTML_TAG_1), 1) <> "/" Then
            TEXT_HTML_TAG_3 = "\"
        Else
            TEXT_HTML_TAG_3 = ""
        End If
        TEXT_HTML_TAG_2 = TEXT_HTML_TAG_0 + "<a href=""" + TEXT_HTML_TAG_1 + TEXT_HTML_TAG_3 + """>" + TEXT_HTML_TAG_1 + "</a>"
        Debug.Print TEXT_HTML_TAG_2
        I_4 = Len(RTB_CLIPPER_FORMAT_CONVERTING.Text)
        RTB_CLIPPER_FORMAT_CONVERTING.Text = Replace(RTB_CLIPPER_FORMAT_CONVERTING.Text, TEXT_HTML_TAG_0 + TEXT_HTML_TAG_1, TEXT_HTML_TAG_2)
        I_5 = Len(RTB_CLIPPER_FORMAT_CONVERTING.Text)
        DO_SOME = True
    End If
    
    XPOS_ = XPOS_ + Len(TEXT_HTML_TAG_2)
    XPOS_ = XPOS_ + 1
Loop Until XPOS_ = 0

If DO_SOME = False Then
    Beep
    Exit Sub
End If
RTB_CLIPPER_FORMAT_CONVERTING.Text = Replace(RTB_CLIPPER_FORMAT_CONVERTING.Text, vbCrLf + vbCrLf, vbCrLf)
EXECUTE_TIMER_ENABLED = False
Clipboard.Clear
EXECUTE_TIMER_ENABLED = True
Clipboard.SetText RTB_CLIPPER_FORMAT_CONVERTING.Text
Me.WindowState = vbMinimized
RTB_CLIPPER_FORMAT_CONVERTING.Text = ""
Beep

'<a href="http://roidsrim.blogspot.co.uk/">http://roidsrim.blogspot.co.uk</a><br />
End Sub

Private Sub MNU_FILE_LOCATOR_IMAGE_Click()


'SHELL" D:\0 00 Art Loggers\# APP AND SCREEN\7-ASUS-GL522VW\FILE LOCATOR __ SavedCriteria.srf

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")

EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE " + EXECUTE_PARAM
EXECUTE_FILE_NAME = "D:\0 00 Art Loggers\# APP AND SCREEN\7-ASUS-GL522VW\FILE LOCATOR __ SavedCriteria.srf"

WSHShell.Run """" + EXECUTE_FILE_NAME + """"

Set WSHShell = Nothing

Rem ---------------------------------------------------------
Rem EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE /SELECT, C:\"
Rem VBS CODE WILL NOT RUN FROM NOTEPAD++
Rem ---------------------------------------------------------


End Sub

Private Sub MNU_LAST_ART_PIC_SELECTOR_EXPLORER_Click()
'
Me.WindowState = vbMinimized
Shell "Explorer.exe /SELECT," + LAST_CLIPBOARD_FILE, vbMaximizedFocus
Beep

End Sub

Private Sub MNU_SLECTOR_WHOLE_LINE_MODE_Click()

If MNU_SLECTOR_WHOLE_LINE_MODE.Caption = FULL_LINE + "=FALSE" Then
    MNU_SLECTOR_WHOLE_LINE_MODE.Caption = FULL_LINE + "=TRUE"
    Exit Sub
End If
If MNU_SLECTOR_WHOLE_LINE_MODE.Caption = FULL_LINE + "=TRUE" Then
    MNU_SLECTOR_WHOLE_LINE_MODE.Caption = FULL_LINE + "=FALSE"
    Exit Sub
End If
End Sub


Private Sub Timer_API_RESET_INFO_Timer()

'CALL MNU_API_RESET_Click
Call MNU_API_RESET_SUB

End Sub

Public Sub MNU_API_RESET_SUB()
'BOOKMARK HERE
If O_VAL_MINUTE_API = 0 Then O_VAL_MINUTE_API = Now

'-----------------------------
'01 OF 03
'NUMERIC IT HAS ARRIVED AT OVER A MINUTE


'STATUS ARRIVED AT RESUME FROM IDLE
'----------------------------------
TAG_FLAG_VAR = False
If DateDiff("s", O_VAL_MINUTE_API, Now) > 60 Then
    If InStr(MNU_IDLE_ACTIVE.Caption, "IDLE") > 0 Then
    
        'IS RESUME FROM IDLE OR OVER LIMIT 1 MINUTE IDLE
        '-----------------------------------------------
    
        'STATE STATUS ABOVE CHANGED
        '--------------------------
        Beep
        TAG_FLAG_VAR = True
        Debug.Print Time$ + "  HERE 01 OF 02"

        O_VAL_MINUTE_API = Now
        Call zzCheckTimer_Timer
        Call RESET_SETUP_SOUND_FILE("NOTSOUND")
        Call Mnu_API_Unload_Click
        Call Mnu_API_Reload_Click
    End If
End If

'-----------------------------------------------
'02 OF 03
'SMALL NUMERIC AND THEN ENTER INTO IDLE

'02 OF 03
'---------------------------------------------------
'STATUS ARRIVED AT RESUME FROM IDLE
'STATUS ARRIVED AT ENTER INTO ACTIVE BY A FEW SECOND
'BOTH
'---------------------------------------------------
If TAG_FLAG_VAR = False Then
    If VAL_MINUTE_API = DateDiff("s", O_VAL_MINUTE_API, Now) < 50 Then
        If MNU_IDLE_ACTIVE.Caption <> oMNU_IDLE_ACTIVE Then
            Beep
            'Debug.Print "----"
            'Debug.Print Time$ + "  HERE 02 OF 03"
            O_VAL_MINUTE_API = Now
            Call zzCheckTimer_Timer
            Call RESET_SETUP_SOUND_FILE("NOTSOUND")
            Call Mnu_API_Unload_Click
            Call Mnu_API_Reload_Click
        End If
    End If
End If

oMNU_IDLE_ACTIVE = MNU_IDLE_ACTIVE.Caption

'---------------
'ASK WHAT GO3 IS
'---------------
'GO3 = 1
'---------------

'SORT THE DISPLAY
If VAL_MINUTE_API = DateDiff("s", O_VAL_MINUTE_API, Now) < 60 Then
    VAL_MINUTE_API = DateDiff("s", O_VAL_MINUTE_API, Now)
    Test_Min_Var = "s"
Else
    VAL_MINUTE_API = DateDiff("n", O_VAL_MINUTE_API, Now)
    Test_Min_Var = "m"
End If

OX_VAL_MINUTE_API = VAL_MINUTE_API

If IsIDE = True Then Test_Min_Var = Test_Min_Var + " -- " + Time$

MNU_API_RESET.Caption = "API=" + Trim(Str(VAL_MINUTE_API)) + Test_Min_Var

End Sub

Private Sub Mnu_API_Unload_Click()

Unload frmClipTest

End Sub


Private Sub Mnu_API_Unload_Reload_Click()
Unload frmClipTest
DoEvents
Load frmClipTest



End Sub



Private Sub Timer_API_Test_Timer()

Timer_API_Test.Enabled = False
Exit Sub
Timer_API_Test.Interval = 10000

'REMOVE THIS ONE AUG 08 2K SIXTEEN

'NOT MUCH USE
'BETTER IDEA
'DO A TIMER CLIPBOARD CHECK COMPARE TO LAST
'BUT NOT NEED WHILE COMPUTER IDLE
'WHEN RESUME FROM IDLE DO A RESET DRIVER API

'NEW TEST IF TIMER WITH DATE VAR IS UPDATING
'NOT GOOD - API STOP FUNCTION WHILE FORM STILL RUNNING
'SO NOT A UNLOAD PROBLEM

If Timer_API_Test_Logick_Var2 > 0 Then
    If Timer_API_Test_Logick_Var2_OLD <> Timer_API_Test_Logick_Var2 Then
        VARTEXT = "THE TIME THE API CLIPBOARD FUNCTION LAST ACCESSED = " + Format(Timer_API_Test_Logick_Var2, "DD-MM-YYYY HH:MM:SS")
        MNU_TIME_API_FUNCTION_ACCESS.Caption = VARTEXT
        Timer_API_Test_Logick_Var2_OLD = Timer_API_Test_Logick_Var2
    End If
End If

If Timer_API_Test_Logick_Var1 = Timer_API_Test_Logick_Var1_OLD Then
    Timer_API_Test_Logick_Var1_Missing_Count = Timer_API_Test_Logick_Var1_Missing_Count + 1
    If Timer_API_Test_Logick_Var1_Missing_Count = 20 Then
    
    Me.WindowState = vbNormal
    DoEvents
    Me.Refresh
    DoEvents

    Timer_API_OKAY_COLOUR.Enabled = False
    I = MsgBox("ClipBoard API Has Stopped and Gone Missing" + vbCrLf + "Use the Menu Option *INFO* to Diagnostic and Reload It" + vbCrLf + "This Can Happen If ChkDsk Unlocked All Handles to The Hard Drive and the ClipboardViewer.ocx Driver Couldn't Get Access", vbMsgBoxSetForeground)
        
    End If
End If


If Timer_API_Test_Logick_Var1 = Timer_API_Test_Logick_Var1_OLD Then Exit Sub
Timer_API_Test_Logick_Var1_OLD = Timer_API_Test_Logick_Var1

If Timer_API_OKAY_COLOUR.Enabled = False Then Timer_API_OKAY_COLOUR.Enabled = True

Timer_API_Test_Logick_Var1_Missing_Count = 0

Mnu_Missing_Link_API_Test.Caption = "Link Detector Check Clipboard API If Is Unloaded = " + Format(Timer_API_Test_Logick_Var1, "DD-MMM-YYYY HH:MM:SS")


ITEXT = "EXECUTE_TIMER_ENABLED = "
If EXECUTE_TIMER_ENABLED = True Then
    ITEXT = ITEXT + "TRUE = OKAY ALLOW API TO WORK"
Else: ITEXT = ITEXT + "FALSE = WRONG API NOT WORKING"
End If
MNU_EXECUTE_TIMER_ENABLED.Caption = ITEXT

If Timer_API_Test_Logick_Var1 = 0 Then Exit Sub

'MNU_INFO.Caption = "INFO " + Mid(Format(Timer_API_Test_Logick_Var1, "HH:MM:SS"), 8, 1)


End Sub






Private Sub MNU_AUDIO_01_MISSING_Click()
Beep
Call RESET_SETUP_SOUND_FILE("")
End Sub

Private Sub MNU_AUDIO_02_MISSING_Click()
Beep
Call RESET_SETUP_SOUND_FILE("")
End Sub

Private Sub Mnu_Audio_ON_Click()

Beep

''** /\/\ Audio Is Off /\/\ ** Hiit On Here for ON ***
Mnu_SoundOn.Checked = True
'Mnu_Audio_ON.Visible = False
Mnu_Audio_ON.Caption = "Audio is On"

Call RESET_SETUP_SOUND_FILE("")


End Sub

Private Sub MNU_Audio_Only_With_Text_and_Picture_Clip_Sound_Bug_Acer_Click()

MNU_Audio_Only_With_Text_and_Picture_Clip_Sound_Bug_Acer.Checked = Not MNU_Audio_Only_With_Text_and_Picture_Clip_Sound_Bug_Acer.Checked
'THIS BIT OF CODE TEST BUG SOUNBD CRASH WHEN CLIP
'REDUCE IT SOME

End Sub

Private Sub MNU_AUDIO_WANT_ON_Click()

Call Mnu_Audio_ON_Click

MNU_AUDIO_WANT_ON.Caption = "AUDIO=ON"

End Sub

Private Sub MNU_BRing_Front_Click()

Call FindWinPartFront(False) 'False = Display Result Count
MNU_BRing_Front.Caption = "Bring All Windows Front -- User Command @ " + Format(Now, "DD-MMM-YYYY HH:MM:SS")

End Sub


Private Sub MNU_CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY_Click()
Call CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY(True)
End Sub

Private Sub Mnu_Clip_Description_Click()
'Mnu_Clip_Description
'Mnu_Clip_Description Clip Format
End Sub

Private Sub Mnu_Clip_Status_Click()
'Mnu_Clip_Status
End Sub



Private Sub MNU_CLIPBOARD_API_PUBLIC_VAR_HOOK_Click()

'CAN'T SEEM TO GET A HOOK VAR PUBLIC IN THE FORM
'THAT WOULD VOID ITSELF IF FOLRM UNLOADED

Exit Sub

Dim HOOKSTAT

I = API_CLIPBOARD_HOOK

'If HOOKSTATold = HOOK_CLIPBOARD_API_lOADED Then Exit Sub
If HOOKSTATold = I Then Exit Sub

HOOKSTATold = I


If I = True Then
    HOOKSTAT = "True"
Else
    HOOKSTAT = "False"
End If




MNU_CLIPBOARD_API_PUBLIC_VAR_HOOK.Caption = "FORM CLIPBOAD API - PUBLIC VAR - SHOWS LOADED IS = " + HOOKSTAT

End Sub




Private Sub MNU_CLIPBOARD_EXPLORER_FILE_FOLDER_Click()

'1..MNU_CLIPBOARD_EXPLORER_FILE_FOLDER_Click
'2..MNU_REG_JUMP_Click
'3..MNU_URL_Browser_Click
'4..MNU_CPC_Click
    
    TEXT_PATH = Clipboard.GetText
    TEXT_PATH_ORI = TEXT_PATH
    TEXT_PATH_L = LCase(TEXT_PATH)
    
    XPOS = 0
    'FOLDER_FILE_PATH
    If XPOS = 0 Then If InStrRev(TEXT_PATH_L, ":\") > 0 Then XPOS = InStrRev(TEXT_PATH_L, ":\") - 1
    'NETWORK
    If XPOS = 0 Then XPOS = InStrRev(LCase(TEXT_PATH), "\\")
    
    TYPE_VAR = "PATH, URL OR REG_KEY"
    TYPE_VAR = "FILE, FOLDER AND NETWORK PATH"
    
    If XPOS = 0 Then
        I = MsgBox("LAST CLIPBOARD NOT CONTAIN A VERIFIABLE STRING " + TYPE_VAR + " TO LOAD" + vbCrLf + vbCrLf + TEXT_PATH + vbCrLf + vbCrLf + "----" + vbCrLf + vbCrLf + "WANT MINIMIZE ", vbYesNo, vbMsgBoxSetForeground)
        If I = vbYes Then Me.WindowState = vbMinimized
        Exit Sub
    End If
   
    If XPOS > 1 Then
        TEXT_PATH = Mid(TEXT_PATH, XPOS, InStr(XPOS, TEXT_PATH, vbCr) - XPOS)
    End If
    If XPOS = 1 And InStr(XPOS, URL_PATH, vbCr) > 0 Then
        TEXT_PATH = Mid(TEXT_PATH, XPOS, InStr(XPOS, TEXT_PATH, vbCr) - XPOS)
    End If
    
    '----------------------------------------
    'FILE SPEC
    '----------------------------------------
    GO3 = False
    FileSpec = TEXT_PATH
    If FS.FileExists(FileSpec) Then
        GO3 = True
        Shell "Explorer.exe /select, " + FileSpec, vbNormalFocus
    End If
    
    '----------------------------------------
    'FOLDER SPEC
    '----------------------------------------
    If FS.FolderExists(FileSpec) = True And GO3 = False Then
        GO3 = True
    End If

    If GO3 = False Then FileSpec_TMP = FindTreeLowerLevelWorkingExistPath(FileSpec)
    
    FileSpec = FileSpec_TMP
    If FS.FolderExists(FileSpec) = True And GO3 = False Then
        GO3 = True
    End If

    If GO3 = False Then
        MINFO = vbCrLf + String(50, "-") + vbCrLf + "SHOWING FIRST 300 CHAR OF STRING CLIPBOARD" + vbCrLf + String(50, "-")
        MINFO = MINFO + vbCrLf + vbCrLf + Mid(TEXT_PATH, 1, 300) + vbCrLf + vbCrLf
        If XPOS > 1 Then
            MINFO = MINFO + vbCrLf + vbCrLf + "SOME PATH INFO WAS SEARCH FOUND AND STILL NOT VERIFY" + vbCrLf + vbCrLf + TEXT_PATH + vbCrLf + vbCrLf
        End If
        
        I = MsgBox("PATH ON CLIPBOARD DOES NOT EXIST AS FILE OR FOLDER EVEN AFTER HACKNG PATH TREE DOWN" + MINFO + "WANT MINIMIZE", vbYesNo, vbMsgBoxSetForeground)
        If I = vbYes Then Me.WindowState = vbMinimized
    Else
        Shell "Explorer.exe /e, " + FileSpec, vbMaximized
        Me.WindowState = vbMinimized
    End If

End Sub

Private Sub MNU_CLIPBOARD_TEST_Click()

FlagTestClipBoardRoutine = True
MsgBox "A Flag Has Been Set So Next Clipboarded Object You Do Should Also See a Message Of Working", vbMsgBoxSetForeground

End Sub

Private Sub MNU_COMPARE_Click()
    Beep
    FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#COMPARE 1 DUPE CHECKER.Txt"
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_COMPARE_1 = FILENAME_IN_USE_CHECK
    FR1 = FreeFile
    On Error Resume Next
    Open FILENAME_IN_USE_CHECK For Output As #FR1
        Print #FR1, COMPARE_EXE_1;
    Close #FR1

    FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#COMPARE 2 DUPE CHECKER.Txt"
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_COMPARE_2 = FILENAME_IN_USE_CHECK
    FR1 = FreeFile
    On Error Resume Next
    Open FILENAME_IN_USE_CHECK For Output As #FR1
        Print #FR1, COMPARE_EXE_3;
    Close #FR1

    Beep
    Shell "C:\Program Files (x86)\WinMerge\WinMergeU.exe" + " """ + FILENAME_COMPARE_1 + """ """ + FILENAME_COMPARE_2 + """"
    Beep
    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_COPY_Click()

'----
'Copy rich text box's text - Xtreme Visual Basic Talk
'http://www.xtremevbtalk.com/general/15668-copy-rich-text-boxs-text.html
'----
    
    MNU_COPY.Caption = "<*** COPY --" + Str(Text1.SelLength) + " ***>"
    
    If Text1.SelLength = 0 Then Beep: Exit Sub
    
    Const WM_COPY = &H301
    EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        SendMessage Text1.hwnd, WM_COPY, 0&, 0& 'Copy
    EXECUTE_TIMER_ENABLED = True
    
    Beep
    
    On Error Resume Next
    Do
        Err.Clear
        AD$ = Clipboard.GetText
        Sleep 100

    Loop Until Err.Number = 0

    Me.WindowState = vbMinimized

    Call COMPARE_FOR_EXE

Exit Sub

'Call CopyFromRichTextBox(Text1.Text)

'MNU_LAST_GRAB_LOW_CASE_TOP_BAR
'AD$ = LCase(AD$)

    'EXECUTE_TIMER_ENABLED = False
    '    Clipboard.Clear
        'Clipboard.SetText AD$
        'Clipboard.SetText Text1.SelectedRtf, vbCFRTF
    'EXECUTE_TIMER_ENABLED = True
    
    'AD$ = Clipboard.GetText

Me.WindowState = vbMinimized
AD_DATE = 0

End Sub

'' Copy the RichTextBox's selected text into the Clipboard
'' Example: CopyFromRichTextBox (richTextBox1)
'
'Public Sub CopyFromRichTextBox(ByVal rtb As RichTextBox, _
'    Optional ByVal availableAfterEnd As Boolean = False)
'
'    Dim data As New DataObject
'
'    ' get the selected RTF text if there is a selection,
'    ' or the entire text is no text is selected
'    Dim rtfText, plainText As String
'
'    If rtb.SelectionLength > 0 Then
'        rtfText = rtb.SelectedRtf
'        plainText = rtb.SelectedText
'    Else
'        rtfText = rtb.Rtf
'        plainText = rtb.Text
'    End If
'
'    ' do the copy only if there is something to be copied
'    If rtfText.Length > 0 Then data.SetData(DataFormats.Rtf, rtfText)
'
'    If plainText.Length > 0 Then data.SetData(DataFormats.Text, plainText)
'
'    ' finally copy into the clipboard
'    Clipboard.SetDataObject(data, availableAfterEnd)
'End Sub

Private Sub MNU_COUNT_JUMP_BOTTOM_Click()
    Text1.Enabled = False
    'Text1.SelStart = 1
    
    SARSOSSAS = "---------------------" + vbCrLf + "Count ="
    IT_TIT1 = InStrRev(Text1.Text, SARSOSSAS, Len(Text1.Text))
    Text1.SelStart = Len(Text1.Text)
    Text1.SelStart = IT_TIT1 - 1
    Text1.SelLength = Len("---------------------")
    Text1.Enabled = True
End Sub

Private Sub MNU_COUNT_JUMP_Click()

End Sub

Private Sub MNU_COUNT_JUMP_DOWN_Click()
    SARSOSSAS = "---------------------" + vbCrLf + "Count ="
    IT_TIT1 = InStr(Text1.SelStart + 2, Text1.Text, SARSOSSAS)
    
    If IT_TIT1 = 0 Then
        'IT_TIT1 = Len(Text1.Text)
        Beep
        Exit Sub
        
    End If
    
    IT_TIT2 = IT_TIT1
    
    For R3 = 1 To 4
        IT_TIT4 = InStr(IT_TIT2 + 1, Text1.Text, vbCrLf)
        If IT_TIT4 <> 0 Then IT_TIT2 = IT_TIT4
        If IT_TIT4 = 0 Then Exit For
        If R3 = 3 Then IT_TIT3 = IT_TIT2
    
    Next R3
    If IT_TIT3 = 0 Then IT_TIT3 = Len(Text1.Text)
    
    Text1.Enabled = False
    'Text1.SelStart = 1
    'Text1.SelStart = Len(Text1.Text)
    'Text1.SelStart = Len(Text1.Text)
    
    Text1.SelStart = IT_TIT2
    'Text1.SelLength = IT_TIT3
    'Text1.SelStart = IT_TIT2
    
    Text1.SelStart = IT_TIT1 - 1
    Text1.SelLength = Len("---------------------")
    
    Text1.Enabled = True
    Text1.SetFocus

End Sub

Private Sub MNU_COUNT_JUMP_TOP_Click()
    Text1.Enabled = False
    Text1.SelStart = 2
    Text1.SelStart = 0
    Text1.SelLength = Len("---------------------")
    Text1.Enabled = True
End Sub

Private Sub MNU_COUNT_JUMP_UP_Click()
    
    SARSOSSAS = "---------------------" + vbCrLf + "Count ="
    IT_TIT1 = InStrRev(Text1.Text, SARSOSSAS, Text1.SelStart - 1)
    
    If Text1.SelStart = 0 Then
        Beep
        Exit Sub
    End If
    
    
    
    If IT_TIT1 = 0 Then
        'IT_TIT1 = 1
        Beep
        Exit Sub
    End If
    
    IT_TIT2 = IT_TIT1
    
    For R3 = 1 To 4
        IT_TIT4 = InStr(IT_TIT2 + 1, Text1.Text, vbCrLf)
        If IT_TIT4 <> 0 Then IT_TIT2 = IT_TIT4
        If IT_TIT4 = 0 Then Exit For
        If R3 = 3 Then IT_TIT3 = IT_TIT2
    
    Next R3
    If IT_TIT3 = 0 Then IT_TIT3 = Len(Text1.Text)
    
    Text1.Enabled = False
    'Text1.SelStart = 1
    'Text1.SelStart = Len(Text1.Text)
    Text1.SelStart = Len(Text1.Text)
    
    Text1.SelStart = IT_TIT1 - 1
    'Text1.SelLength = IT_TIT3
    'Text1.SelStart = IT_TIT2
    
    Text1.SelStart = IT_TIT1 - 1
    Text1.SelLength = Len("---------------------")
    
    Text1.Enabled = True
    Text1.SetFocus

End Sub

Private Sub MNU_CPC_Click()

'1..MNU_CLIPBOARD_EXPLORER_FILE_FOLDER_Click
'2..MNU_REG_JUMP_Click
'3..MNU_URL_Browser_Click
'4..MNU_CPC_Click
    
    URL_PATH = Clipboard.GetText
    
    'URL_PATH = "http://cpc.farnell.com/startech/st93007u2c/hub-7-port-usb3-0-2x2-4a-ports/dp/CS29423"
    
    XPOS = InStrRev(LCase(URL_PATH), "http")
    
    
    
    'TYPE_VAR = "PATH, URL OR REG_KEY"
    'TYPE_VAR = "FOR USE WITH CPC"

    'If XPOS = 0 Then
        'I = MsgBox("LAST CLIPBOARD NOT CONTAIN A VERIFIABLE STRING " + TYPE_VAR + " TO LOAD" + vbCrLf + vbCrLf + TEXT_PATH + vbCrLf + vbCrLf + "----" + vbCrLf + vbCrLf + "WANT MINIMIZE ", vbYesNo, vbMsgBoxSetForeground)
        'If I = vbYes Then Me.WindowState = vbMinimized
        'Exit Sub
    'End If
   
    If XPOS > 0 Then
        If InStr(XPOS, URL_PATH, vbCr) > 0 Then
            URL_PATH = Mid(URL_PATH, XPOS, InStr(XPOS, URL_PATH, vbCr) - XPOS)
        Else
            URL_PATH = Mid(URL_PATH, XPOS)
        End If
    
    
    End If
    
    If XPOS = 0 And InStrRev(LCase(URL_PATH), "http") > 0 Then
        '---------------------------------
        URL_PATH = AD$
        XPOS = InStrRev(LCase(URL_PATH), "http")
        If InStr(XPOS, URL_PATH, vbCr) > 0 Then
            URL_PATH = Mid(URL_PATH, XPOS, InStr(XPOS, URL_PATH, vbCr) - XPOS)
        Else
            URL_PATH = Mid(URL_PATH, XPOS)
        End If
    End If
    
    
    If InStr(LCase(URL_PATH), "http://cpc.farnell.com/") = 0 Then
        MsgBox "MUST BE A CPC URL PATH " + vbCrLf + "http://cpc.farnell.com/" + vbCrLf + "GIVE RESULT WAS" + URL_PATH, vbMsgBoxSetForeground
        Exit Sub
    End If
    'http://cpc.farnell.com/pro-power/ppc100/foam-cleaner-400ml-aerosol/dp/SA01881
    
    'URL_PATH = "http://cpc.farnell.com/webapp/wcs/stores/servlet/ProductDisplay?catalogId=15002&langId=69&urlRequestType=Base&partNumber=HK01161&storeId=10180"
    
    If InStr(LCase(URL_PATH), LCase("&storeId=")) > 0 Then
        URL_PATH = Mid(URL_PATH, 1, InStr(LCase(URL_PATH), LCase("&storeId=")) - 1)
    End If
    'http://cpc.farnell.com/webapp/wcs/stores/servlet/ProductDisplay?catalogId=15002&langId=69&urlRequestType=Base&partNumber=HK01161&storeId=10180    '
    
    Me.WindowState = vbMinimized
    
    'Shell "Explorer.exe /e," + Clipboard.GetText, vbNormalFocus
    vFile = URL_PATH 'Clipboard.GetText
    
    'MsgBox "READY TO LOAD 100 WEB PAGES" + VBCLRF + "PAUSE IS RIGHT CLICK MOUSE" + vbCrLf + "ESCAPE IS ABORT" + vbCrLf + "AND QUESTION CONFIRM", vbMsgBoxSetForeground
'    MsgBox "READY TO LOAD 100 WEB PAGES OF --" + vbCrLf + URL_PATH + vbCrLf + vbCrLf + "CURRENTLY PAGE NUMBERING OF 0# AN 1# AND 8# AND 9# ARE USED 40 TOTAL " + vbCrLf + "1 SECOND GAP INTERVAL WILL BE USED WITH SLEEP COMMAND" + vbCrLf + "YOU WONT STOP THIS ONCE GOING" + vbCrLf + "THEY WILL ALL BE BUFFERED IN", vbMsgBoxSetForeground
    
    ISIDE_VAR = IsIDE
    Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    Sleep 1000

    
    For RCPC = 0 To 99
    
        MNU_CPC_TOP_MENU.Caption = "CPC -- 00 80 90 -- " + Format(RCPC, "00")
    
        If Me.WindowState <> vbMinimized Then
            RESUME_GO_CPC = False
            Do
                Sleep 100
                Me.WindowState = vbMinimized
            
            Loop Until RESUME_GO_CPC = True Or Me.WindowState = vbMinimized
        End If
    
        IRCPC = Mid(Format(RCPC, "00"), 1, 1)
        If InStr("*089", IRCPC) > 0 Or 1 = 2 Then
            ShellExecute Me.hwnd, "open", vFile + Format(RCPC, "00"), vbNullString, vbNullString, conSwNormal
        
            Sleep 1000
            URL_COUNT = URL_COUNT + 1
        
            
        
            If URL_COUNT > 10 Then
                URL_COUNT = 0
                Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", vbNormalFocus
                Sleep 1000
            End If
        
            If 1 = 2 Then
                If RCPC = 0 Then OIRCPC = IRCPC
                If IRCPC <> OIRCPC Then
                    OIRCPC = IRCPC
                    'LOAD A NEW CHROME WINDOW EVERY 10 LOADED HOPE THEY LOAD INTO NEW CURRENT WINDOW
                    Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
                    Sleep 1000
                End If
            End If
        End If



        'DELAY BLOCK CODE NOT USED
        If 1 = 2 Then
            INOW = Int(Now) + Timer + 1
            Do
                
                I = 0
                For I = 0 To 255
                    GET_KEY = GetAsyncKeyState(I)
                '    If GET_KEY < -300 Then GET_KEY = I: Debug.Print GET_KEY: Exit For
                
                    'I <> 116 -- NOT F5 FOR TEST
                    If GET_KEY < -300 And (ISIDE_VAR = False And I <> 116) Then
                        GET_KEY = I
                        'Debug.Print GET_KEY
                        Exit For
                    End If
                Next
    
                
                DO_FLAG = False
                'PAGE UP
                'If GET_KEY = 33 Then DO_FLAG = True
                'ALT
                'If GET_KEY = 18 Then DO_FLAG = True
                'MOUSE WHEEL PUSH
                
                If GET_KEY = 16 Then DO_FLAG = True
                'MOUSE RIGHT CLICK
                If GET_KEY = 16 Then
                    DO_FLAG = True
                    I_ANSWER = MsgBox("MOUSE RIGHT CLICK PRESS" + vbCrLf + "REQUEST TO PAUSE LOAD 100 CPC WEB PAGES REQUESTED @ # " + vbCrLf + "# " + Format(RCPC, "00") + vbCrLf + " ENTER OKAY TO CONTINUE", vbMsgBoxSetForeground)
                End If
                
                
                'ESCAPE
                If GET_KEY = 27 Then
                    DO_FLAG = True
                    I_ANSWER = MsgBox("ESCAPE KEY PRESS" + vbCrLf + "REQUEST ABORT LOAD 100 CPC WEB PAGES REQUESTED @ # " + vbCrLf + "# " + Format(RCPC, "00") + vbCrLf + " IS THAT WHAT YOU WANT YES / NOT", vbYesNoCancel + vbMsgBoxSetForeground)
                    If IANSWER = vbYes Then Exit Sub
                End If
                
    
            COUNTLOOP = COUNTLOOP + 1
            Loop Until INOW < Int(Now) + Timer Or DO_FLAG = True
            
            
            Debug.Print COUNTLOOP

        End If

    Next RCPC




Me.WindowState = vbMinimized
RCPC = 0

End Sub

Private Sub MNU_CPC_TOP_MENU_Click()

If RCPC = 0 Then Call MNU_CPC_Click
'SUB ABOVE
RESUME_GO_CPC = True

MNU_404_CPC_PAGES.Visible = True
VAR_MNU_404_HITT_ONCE = True

End Sub


Private Sub MNU_404_CPC_PAGES_Click()

If CPC_404_CPC = "" Then CPC_404_CPC = MNU_404_CPC_PAGES.Caption
MNU_404_CPC_PAGES.Enabled = False
VAR_MNU_404_HITT_ONCE = True
Beep
Exit Sub

Call FindWinPart_404_REMOVE_TAB_CHROME(HUGE1, HUGE2) 'Display Result Count

If HUGE2 > 0 Or HUGE1 > 0 Then
    MsgBox "PAGE 404 REMOVED" + vbCrLf + "FOUND --" + Str(HUGE1) + vbCrLf + "REMOVED SUCCESS --" + Str(HUGE2), vbMsgBoxSetForeground
Else
    MsgBox "PAGE 404 REMOVED -- DONE -- NONE TO DO", vbMsgBoxSetForeground
End If


End Sub


Private Sub MNU_EBAY_BEGIN_FILTER_SAVE_EDITOR_Click()

    FILENAME_VAR = "C:\TEMP\EBAY FILTER RESULT.TXT"
    If Dir("C:\TEMP", vbDirectory) = "" Then
        MkDir ("C:\TEMP")
    End If
    'DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    
    EBAY_TEXT = Clipboard.GetText
    
    FR1 = FreeFile
    On Error Resume Next
'    Close FR1
'    Err.Clear
'    Reset
    Open FILENAME_VAR For Output As #FR1
        
        If Err.Number > 0 Then
            MsgBox FILENAME_VAR + vbCrLf + "WOULD NOT OPEN -- BE A PROBLEM" + vbCrLf + Err.Description, vbMsgBoxSetForeground
            Exit Sub
        End If
        
        Print #FR1, EBAY_TEXT;
        
        If Err.Number > 0 Then
            MsgBox FILENAME_VAR + vbCrLf + "WAS NOT SAVED -- BE A PROBLEM" + vbCrLf + Err.Description, vbMsgBoxSetForeground
            Exit Sub
        End If
    
    Close FR1
    
    On Error GoTo 0


    'SOME FILTER
    FIRST_LINE = True
    
    FR1 = FreeFile
    Open FILENAME_VAR + ".TMP" For Output As #FR1
    FR2 = FreeFile
    Open FILENAME_VAR For Input As #FR2
            
        If Not EOF(FR2) Then
        
            Do
                NOT_TO_GO = False
                
                Line Input #FR2, LINE_TEXT
            
                '-----------------------------------------------
                'REMOVE THE DUPE LINE THAT HAS DIFF DOUBLE SPACE
                '-----------------------------------------------
                If LINE_TEXT = OLINE_TEXT Then NOT_TO_GO = True
                OLINE_TEXT = Replace(LINE_TEXT, "  ", " ")
            
                If InStr("-" + LINE_TEXT, "This seller accepts PayPal") > 0 Then
                    NOT_TO_GO = True
                End If
                If InStr(LINE_TEXT, "Watch") = 1 Then
                    NOT_TO_GO = True
                End If
                If InStr(LINE_TEXT, "Item:") = 1 Then
                    NOT_TO_GO = True
                End If
                If InStr("-" + LINE_TEXT, "Free Postage") > 0 Then
                    NOT_TO_GO = True
                End If
                If InStr("-" + LINE_TEXT, "Watch") > 0 Then
                    NOT_TO_GO = True
                End If
                If InStr("-" + LINE_TEXT, "AdChoice") > 0 Then
                    BEGIN_TO_GO = True
                    NOT_TO_GO = True
                End If
                
                'WHEN ALMOST DONE
                If InStr("-" + LINE_TEXT, "Sponsored links") > 0 Then
                    NOT_TO_GO = True
                End If
                
                If InStr("-" + LINE_TEXT, "Tell us what you think") > 0 Then
                    NOT_TO_GO = True
                    NOT_TO_GO_UNTIL_END = True
                End If
                
                
                'BEGINING IS HERE -- 02 OF 02
                If InStr("-" + LINE_TEXT, "Items in search results") > 0 Then
                    
                    EXTRA_HEAD_TEXT_LINE = String$(80, "-") + vbCrLf + String$(80, "-") + vbCrLf + "BEGINING IS HERE -- 02 OF 02 --" + Format(Now, "DDD DD MMM YYYY HH:MM:SS") + vbCrLf + String$(80, "-") + vbCrLf + String$(80, "-") + vbCrLf + String$(80, "-") + vbCrLf + "COUNT -- 001 -- BLOCK OF TEXT BELOW" + vbCrLf + String$(80, "-") + vbCrLf + String$(80, "-") + vbCrLf
                
                End If
                
                'END AND START EACH LINE
                If InStr(LINE_TEXT, "Seller:") = 1 Then
                    COUNT_SALE = COUNT_SALE + 1
                    EXTRA_HEAD_TEXT_LINE = String$(80, "-") + vbCrLf + "COUNT -- " + Format(COUNT_SALE, "000") + " -- BLOCK OF TEXT ABOVE" + vbCrLf + String$(80, "-") + vbCrLf + String$(80, "-") + vbCrLf
                End If
                
                
                If NOT_TO_GO_UNTIL_END = True Then
                    NOT_TO_GO = True
                End If
                
                If BEGIN_TO_GO = False Then
                    NOT_TO_GO = True
                End If
                
                
                'OKAY TO HERE BUT NOT ANY DOUBLE REPEAT LINE SPACE
                If NOT_TO_GO = False Then
                    If Trim(OLINE_TEXT_VAR2) = Trim(LINE_TEXT) Then
                        NOT_TO_GO = True
                    End If
                    'NOT ANY DOUBLE REPEAT LINE SPACE
                    OLINE_TEXT_VAR2 = LINE_TEXT
                End If
                
                
                If NOT_TO_GO = False Then
                
                    'BEGINING IS HERE -- 01 OF 02 --
                    If FIRST_LINE = True Then
                        FIRST_LINE = False
                        EXTRA_HEAD_TEXT_LINE = String$(80, "-") + vbCrLf + String$(80, "-") + vbCrLf + "BEGINING IS HERE -- 01 OF 02 --" + Format(Now, "DDD DD MMM YYYY HH:MM:SS") + vbCrLf + String$(80, "-") + vbCrLf + String$(80, "-")

                    End If
                    
                    LINE_TEXT = Replace(LINE_TEXT, Chr(9), "")
                    LINE_TEXT = Replace(LINE_TEXT, "We're sorry,", "")
                    LINE_TEXT = Replace(LINE_TEXT, "  Follow this search", "")
                
                    If Trim(LINE_TEXT) <> "" Then
                    If EXTRA_HEAD_TEXT_LINE <> "" Then
                        EXTRA_HEAD_TEXT_LINE = vbCrLf + EXTRA_HEAD_TEXT_LINE
                    End If
                    Print #FR1, Trim(LINE_TEXT) + EXTRA_HEAD_TEXT_LINE
                    End If
                
                    EXTRA_HEAD_TEXT_LINE = ""
                
                End If
                
                
                
            Loop Until EOF(FR2)
            
        End If
        
    Close #FR1, FR2

    If BEGIN_TO_GO = False Then
        MsgBox FILENAME_VAR + vbCrLf + " FIRST HEAD LINE TEXT NOT FOUND -- RESULT WILL BE EMPTY, vbMsgBoxSetForeground"
        Exit Sub
    End If
    


    If Err.Number > 0 Then
        MsgBox FILENAME_VAR + vbCrLf + "CHANGES WERE NOT SAVED -- BE A PROBLEM" + vbCrLf + Err.Description, vbMsgBoxSetForeground
        Exit Sub
    End If

    Kill FILENAME_VAR
    Name FILENAME_VAR + ".TMP" As FILENAME_VAR
    
    If Err.Number > 0 Then
        MsgBox FILENAME_VAR + vbCrLf + "RENAME TEMP FILE -- BE A PROBLEM" + vbCrLf + Err.Description, vbMsgBoxSetForeground
        Exit Sub
    End If



    
    
    'LOAD EDITOR NOTEPAD++
    
    vFile = FILENAME_VAR
    ShellExecute Me.hwnd, "open", vFile, vbNullString, vbNullString, conSwNormal
    



End Sub

Private Sub MNU_EBAY_FILTER_2222WILDCARD2222_Click()
      'Begin VB.Menu MNU_EBAY_FILTER_WILDCARD
          'fault at load program code not allow menu
          '-----------------------------------------

    FILENAME_VAR = "C:\TEMP\EBAY FILTER RESULT.TXT"
    If Dir("C:\TEMP", vbDirectory) = "" Then
        MkDir ("C:\TEMP")
    End If
    'DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    
    'SOME FILTER
    
    Dim LINE_TEXT_ARRAY(200, 20)
    
    Reset
    FR1 = FreeFile
    Open FILENAME_VAR + ".TMP" For Output As #FR1
    FR2 = FreeFile
    Open FILENAME_VAR For Input As #FR2
            
        If Not EOF(FR2) Then
        
            Do
                
                Line Input #FR2, LINE_TEXT
                Debug.Print LINE_TEXT
                'SEARCH THE COUNT TEXT FOR A HITT
                If InStr("-" + LINE_TEXT, "COUNT -- ") > 0 Then
                    FOUND_BEGINING = True
                    LINE_COUNT_HITT1 = LINE_COUNT_HITT1 + 1
                    LINE_COUNT_HITT2 = 0
                End If
            
                If FOUND_BEGINING = True Then
                    LINE_COUNT_HITT2 = LINE_COUNT_HITT2 + 1
                    LINE_TEXT_ARRAY(LINE_COUNT_HITT1, LINE_COUNT_HITT2) = LINE_TEXT
                End If
                
                
                If FOUND_BEGINING = False Then
                    Print #FR1, LINE_TEXT
                End If
                
            Loop Until EOF(FR2)
            Close #FR1
            Reset
            
            FR1 = FreeFile
            Open FILENAME_VAR + ".TMP" For Append As #FR1
            
            For r = 1 To LINE_COUNT_HITT1
                NOT_TO_GO = False
                If InStr("-" + LINE_TEXT_ARRAY(r, 5), "*") > 0 Then NOT_TO_GO = True
                
                'TEST ALWAYS
                'NOT_TO_GO = False
                
                If NOT_TO_GO = False Then
                    For R1 = 1 To 20
                        If LINE_TEXT_ARRAY(r, R1) <> "" Then
                            If InStr("-" + LINE_TEXT_ARRAY(r, R1), "COUNT -- ") > 0 Then
                                If InStr("-" + LINE_TEXT_ARRAY(r, R1), "ABOVE") > 0 Then
                                    COUNT_SALE = COUNT_SALE + 1
                                    'CS_STR_VAR = "COUNT -- " + Format(COUNT_SALE, "000") + " -- "
'                                Else
'                                Stop
                                End If
                                
                                CS_STR_VAR = "COUNT -- " + Format(COUNT_SALE, "000") + " -- "
                                
                                LINE_TEXT_ARRAY(r, R1) = CS_STR_VAR + LINE_TEXT_ARRAY(r, R1)
                                
                                'Stop
                            End If
                            
                            
                            Print #FR1, LINE_TEXT_ARRAY(r, R1)
                        End If
                    Next
                End If
            Next
        End If
    Close #FR1, FR2


    If Err.Number > 0 Then
        MsgBox FILENAME_VAR + vbCrLf + "CHANGES WERE NOT SAVED -- BE A PROBLEM" + vbCrLf + Err.Description, vbMsgBoxSetForeground
        Exit Sub
    End If

    Kill FILENAME_VAR
    Name FILENAME_VAR + ".TMP" As FILENAME_VAR
    
    If Err.Number > 0 Then
        MsgBox FILENAME_VAR + vbCrLf + "RENAME TEMP FILE -- BE A PROBLEM" + vbCrLf + Err.Description, vbMsgBoxSetForeground
        Exit Sub
    End If
    
    
    'LOAD EDITOR NOTEPAD++
    
    vFile = FILENAME_VAR
    ShellExecute Me.hwnd, "open", vFile, vbNullString, vbNullString, conSwNormal
    




End Sub

Private Sub MNU_ENTER_LARGE_Click()
    
    ENTER_LARGE_IN_LOGGER = True
    Call Timer1_Timer
End Sub

Private Sub MNU_ERROR_INFO_Click()
'ANY ERROR
End Sub

Private Sub Mnu_Explorer_Form_Capture_Click()

'AUTO_FORM_IMAGE_CAPTURE

'called at - ---Timer_SCREEN_SHOT_Timer

Call LAST_IMAGE("EXPLORER", FOLDERNAME_AUTO)

'Shell "Explorer.exe " + FOLDERNAME2, vbNormalFocus

End Sub

Private Sub MNU_EXPLORER_ME_VB_Click()

Beep
Me.WindowState = vbMinimized

'------------------------------------------------------------
'THIS IS DONE -- AT FORM LOAD
'------------------------------------------------------------
'MNU_EXPLORER_ME_VB.Caption = "EXPLORER ME_VB -- " + App.Path
'------------------------------------------------------------
Shell "Explorer.exe " + App.Path, vbNormalFocus

End Sub

Private Sub Mnu_Explorer_Screen_Capture_Click()

'AUTO SCREEN IMAGE CAPTURE

'called at - ---Timer_SCREEN_SHOT_Timer

Call LAST_IMAGE("EXPLORER", FOLDERNAME_AUTO)

'Shell "Explorer.exe " + FOLDERNAME1, vbNormalFocus

End Sub

Sub LAST_IMAGE(VAR1, VAR2)

'XGO = False
'If GetComputerName = "1-ASUS-EEE" Then XGO = True
'If GetComputerName = "1-ASUS-X5DIJ" Then XGO = True


'GIVE UP ON THIS A BIT

VAR2 = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\AUTO_Form_Shot\"

MNU_SCANPATH_COUNTER.Visible = True


'ScanPath.SHOW
'DONT KEEP SCREEN SHOT ANY MORE ONLY APP SHOT

'LAST_FILE_DATE_PATH_HOT_KEY_SCREENSHOT = ""
LAST_FILE_DATE_PATH = ""
'Me.WindowState = vbMinimized

'XdATE2 = 0
'ScanPath.chk_LIST_VIEW_SHORT_5 = vbChecked

LAST_FILE_DATE_TIME = DateSerial(100, 1, 1)
'SCAN_PARTMASK = "# APP AND SCREEN\"+GETCOMPUTERNAME+"\Hot-Key-"

ScanPath.cboMask = "*.JPG"
ScanPath.chkSubFolders = vbUnchecked
ScanPath.TxtPath = VAR2

'Debug.Print VAR2

'THIS LOOK GOOD BUT TAKE TOO LONG DON'T KNOW HOW MICROSOFT DO IT
'QUICKER WITH FOLDER UNCHECKed BUT NOT AS QUiCK MICROSOFT
'TEST WORKING BUT NOT GOOD ENOUGH FOR NOW
'Call ScanPath.cmdScanDir_FAST_Click
'ScanPath.TxtPath = LAST_FILE_DATE_PATH

If Dir(ScanPath.TxtPath, vbDirectory) = "" Then MsgBox "NOT THAT FOLDER" + vbCrLf + ScanPath.TxtPath: Exit Sub

If FS.FileExists(VAR2) = True Then
    'ScanPath.TxtPath = Mid(VAR2, 1, InStrRev(VAR2, "\"))
    'Dir1.Path = ScanPath.TxtPath
End If

Dir1.Path = ScanPath.TxtPath
ScanPath.TxtPath = Dir1.List(Dir1.ListCount - 1)

LAST_FILE_DATE_TIME = DateSerial(100, 1, 1)
ScanPath.chkSubFolders = vbChecked
Call ScanPath.CMDScan_NO_LIST_FAST_Click
'SCAN_PARTMASK = ""

FileSpec = LAST_FILE_DATE_PATH
'FileSpec = LAST_FILE_DATE_PATH_HOT_KEY_SCREENSHOT

If FileSpec = "" Then MsgBox "NOT ANY OF THOSE FILES" + vbCrLf + vbCrLf + ScanPath.TxtPath + "\" + ScanPath.cboMask: Exit Sub


'Filespec1 = ScanPath.lblCount7
'Set F = FS.getfile((Filespec1))
'ADATE1 = F.datelastmodified
'
'ScanPath.lblCount7 = ""
'ScanPath.ListView1.ListItems.Clear
'
'ScanPath.TxtPath = "D:\0 00 Art Loggers\# APP AND SCREEN\"
'Call ScanPath.cmdScan_Click


'Filespec2 = ScanPath.lblCount7
'If Filespec2 <> "" Then
'    Set F = FS.getfile((Filespec2))
'    ADATE2 = F.datelastmodified
'    If ADATE1 > ADATE2 Then
'        FileSpec = Filespec1
'    Else
'        FileSpec = Filespec2
'
'    End If
'Else
'    FileSpec = Filespec1
'End If


'Me.WindowState = vbMinimized

If MNU_MESSAGE_BOXES.Checked = False Then
'    MsgBox "FOUND LATEST IMAGE Clipboarded - LOAD Explorer Minimized AS Well as IrFan Maximized To View" + vbCrLf + "FILES FOUND =" + str(tFileCount) + vbCrLf, vbMsgBoxSetForeground
End If







If VAR1 = "EXPLORER" Then
 '   Shell "Explorer.exe /select, " + FileSpec, vbNormalFocus
'    Exit Sub

    'VAR2 = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\AUTO_Form_Shot\*.JPG"
    VAR2 = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName '+ "\AUTO_Form_Shot\"
    
    On Error Resume Next
    PID = Shell("C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe " + """" + VAR2 + """", vbMaximizedFocus)
    If PID > 0 Then Exit Sub
    PID = Shell("C:\Program Files (x86)\Mythicsoft\FileLocator Pro\FileLocatorPro.exe " + """" + VAR2 + """", vbMaximizedFocus)
    If PID > 0 Then Exit Sub
    PID = Shell("C:\Program Files\Mythicsoft\FileLocator Lite\FileLocatorLite.exe " + """" + VAR2 + """", vbMaximizedFocus)
    If PID > 0 Then Exit Sub
    PID = Shell("C:\Program Files (x86)\Mythicsoft\FileLocator Lite\FileLocatorLite.exe " + """" + VAR2 + """", vbMaximizedFocus)
    If PID > 0 Then Exit Sub
    Shell "Explorer.exe /select, " + VAR2, vbNormalFocus
    
    
    
    'D:\#0 1 INSTALLATIONS\Installation programs\# 00 New Install Progs\# Installed Now\#01 Matthew's\Agent RanSack.exe
    'D:\#0 1 INSTALLATIONS\Installation programs\# 00 New Install Progs\# Installed Now\#00 Paid For\AGENT RANSACK\FILELOCATOR LITE\FileLocatorLite_828.exe
    'D:\#0 1 INSTALLATIONS\Installation programs\# 00 New Install Progs\# Installed Now\#00 Paid For\AGENT RANSACK\FILELOCATOR PRO\flpro_2654 -- 64Bit.exe

    Exit Sub

End If



If IRFANVIEW_PATH <> "" Then
    If VAR1 = "IVIEW" Then
        Shell IRFANVIEW_PATH + " """ + FileSpec + """ /fs /silent", vbMaximizedFocus
    End If
Else
    Me.WindowState = vbNormal
    MsgBox "IRFANVIEW_PATH VAR -- PATH NOT FOUND FOR FILE" + vbCrLf + "NOT INSTALED AT EXPECTED LOCATION " + IRFANVIEW_PATH3 + vbCrLf + "OR" + vbCrLf + IRFANVIEW_PATH2, vbMsgBoxSetForeground
End If

End Sub






Private Sub MNU_FORMAT_PLAIN_TEXT_TOP_BAR_Click()
Beep
    Call MNU_FORMAT_PLAIN_TEXT_Click

End Sub





Private Sub Mnu_IVIEW_Form_Capture_Click()

'called at - ---Timer_SCREEN_SHOT_Timer

Call LAST_IMAGE("IVIEW", FOLDERNAME_AUTO)

End Sub

Private Sub MNU_IVIEW_LAST_WEBCAM_PIC_Click()

Call LAST_IMAGE("IVIEW", "D:\0 00 ART LOGGERS - WEBCAM\WEBCAM\") '=EXPLORER

End Sub

Private Sub MNU_JUMP_ANY_SPECIAL_FOLDER_Click()
Call MNU_JUMP_ANY_SPECIAL_FOLDER_Click
Beep

Me.WindowState = vbMinimized
Load Form2_ANY_SPECIAL_FOLDER

End Sub

Private Sub mnu_Keyboard_Logger_Remove_Doubel_Char_Into_One_Click()

'HERE IS --
    
Dim EE As String

EE = AD$
   
'If EGG = 0 Then EE = LCase(EE)
EE = LCase(EE)

'Mid$(EE, 1, 1) = UCase(Mid$(EE, 1, 1))

For r = 65 To 65 + 25
    EE = Replace(EE, LCase(Chr(r)) + LCase(Chr(r)), LCase(Chr(r)))  ', , -1, vbBinaryCompare)
    EE = Replace(EE, "  ", " ")
'    EE = Replace(EE, vbCr + LCase(Chr(r)), vbCr + UCase(Chr(r)))
'    EE = Replace(EE, "-" + LCase(Chr(r)), "-" + UCase(Chr(r)))
'    EE = Replace(EE, "(" + LCase(Chr(r)), "-" + UCase(Chr(r)))
'    EE = Replace(EE, "." + LCase(Chr(r)), "." + UCase(Chr(r)))
'    EE = Replace(EE, "\" + LCase(Chr(r)), "\" + UCase(Chr(r)))
'    EE = Replace(EE, "_" + LCase(Chr(r)), "_" + UCase(Chr(r)))
Next

EE = Replace(EE, Chr(17), " ") ' CLIP LOGGER CHAR __ DC1
EE = Replace(EE, Chr(16), " ") ' CLIP LOGGER CHAR __ DLE
EE = Replace(EE, Chr(8), " ")  ' CLIP LOGGER CHAR __ BS BACK SPACE

Dim O_LEN_EE
Do
    O_LEN_EE = EE
    EE = Replace(EE, "  ", " ") 'DOUBLE SPACE TO ONE SpACE
    EE = Replace(EE, "((", "(")
    EE = Replace(EE, "&&", "&")
    EE = Replace(EE, "$$", "$")
    EE = Replace(EE, "..", ".")
    EE = Replace(EE, "%%", "%")
    EE = Replace(EE, "##", "#")
    EE = Replace(EE, "''", "'")
    EE = Replace(EE, "½½", "½")
    EE = Replace(EE, "¼¼", "¼")
    EE = Replace(EE, "¾¾", "¾")
    
    EE = Replace(EE, "{{", "{")
    EE = Replace(EE, "ºº", "'") 'KEY BETWEEN DON'T THINKER
    EE = Replace(EE, "($", " ")
    EE = Replace(EE, "(.", " ")
    
    EE = Replace(EE, "(&", " ")
    EE = Replace(EE, "( ", " ")
    EE = Replace(EE, "('", " ")
    EE = Replace(EE, "(&", " ")
    
    EE = Replace(EE, "'%", " ")
    EE = Replace(EE, "'(", " ")
    EE = Replace(EE, "½'%", " ")
    EE = Replace(EE, "½'", " ")
    EE = Replace(EE, "&#", " ")
    
    EE = Replace(EE, "$&", " ")
    EE = Replace(EE, "&(", " ")
    
    EE = Replace(EE, "à(", " ")
    EE = Replace(EE, "(%", " ")
    EE = Replace(EE, "%#", " ")
    EE = Replace(EE, "%#", " ")
    EE = Replace(EE, "%&%", " ")
    
    EE = Replace(EE, "%'", " ")
    EE = Replace(EE, "%.", " ")
    EE = Replace(EE, "%&", " ")
    
Loop Until EE = O_LEN_EE

AD$ = EE

Beep
'HERE IS -- PROCAPS
Call MNU_FORMAT_MINE_Click
    
    EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
    EXECUTE_TIMER_ENABLED = True
    EXECUTE_TIMER_ENABLED = False
        Clipboard.SetText AD$
    EXECUTE_TIMER_ENABLED = True

AD_DATE = 0

Me.WindowState = 1

'Text1.Text = EE

End Sub

Private Sub MNU_LAST_ART_PIC_IVIEW_Click()

Call LAST_IMAGE("IVIEW", "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\Hot-Key-App-Shots\") ' = I_VIEW

Exit Sub


'-------------------------

'ScanPath.SHOW

Me.WindowState = vbMinimized

'XdATE2 = 0
'ScanPath.chk_LIST_VIEW_SHORT_5 = vbChecked

LAST_FILE_DATE_TIME = DateSerial(100, 1, 1)
SCAN_PARTMASK = "# APP AND SCREEN\" + GetComputerName + "\CLIP_"
ScanPath.cboMask = "*.JPG"
ScanPath.chkSubFolders = vbChecked
ScanPath.TxtPath = "D:\0 00 Art Loggers\# APP AND SCREEN\"

Call ScanPath.CMDScan_NO_LIST_FAST_Click
SCAN_PARTMASK = ""

FileSpec = LAST_FILE_DATE_PATH
'FileSpec = LAST_FILE_DATE_PATH_HOT_KEY_SCREENSHOT

If FileSpec = "" Then MsgBox "NOT ANY OF THOSE FILES" + vbCrLf + ScanPath.TxtPath + "\" + ScanPath.cboMask: Exit Sub


'Filespec1 = ScanPath.lblCount7
'Set F = FS.getfile((Filespec1))
'ADATE1 = F.datelastmodified
'
'ScanPath.lblCount7 = ""
'ScanPath.ListView1.ListItems.Clear
'
'ScanPath.TxtPath = "D:\0 00 Art Loggers\# APP AND SCREEN\"
'Call ScanPath.cmdScan_Click


'Filespec2 = ScanPath.lblCount7
'If Filespec2 <> "" Then
'    Set F = FS.getfile((Filespec2))
'    ADATE2 = F.datelastmodified
'    If ADATE1 > ADATE2 Then
'        FileSpec = Filespec1
'    Else
'        FileSpec = Filespec2
'
'    End If
'Else
'    FileSpec = Filespec1
'End If


Me.WindowState = vbMinimized

If MNU_MESSAGE_BOXES.Checked = False Then
'    MsgBox "FOUND LATEST IMAGE Clipboarded - LOAD Explorer Minimized AS Well as IrFan Maximized To View" + vbCrLf + "FILES FOUND =" + str(tFileCount) + vbCrLf, vbMsgBoxSetForeground
End If

'Shell "Explorer.exe /select, " + FileSpec, vbMinimizedNoFocus

If IRFANVIEW_PATH <> "" Then
    If VAR1 = "IVIEW" Then
        Shell IRFANVIEW_PATH + " """ + FileSpec + """ /fs /silent", vbMaximizedFocus
    End If
Else
    Me.WindowState = vbNormal
    MsgBox "IRFANVIEW_PATH VAR -- PATH NOT FOUND FOR FILE" + vbCrLf + "NOT INSTALED AT EXPECTED LOCATION " + IRFANVIEW_PATH3 + vbCrLf + "OR" + vbCrLf + IRFANVIEW_PATH2, vbMsgBoxSetForeground
End If

End Sub


Private Sub MNU_MIRROR_SEND_TO_OPERATING_SYSTEM_SET_MENU_CAPTION()
    MNU_MIRROR_SEND_TO_OPERATING_SYSTEM.Caption = "MIRROR COPY OUR FAT32 SPECIAL SEND-TO FOLDER TO ANY OPERATING SYSTEM - SPECIAL FOLDER SEND TO"
End Sub

Private Sub Mnu_LAST_CLIP_TIME_Click()
'Mnu_LAST_CLIP_TIME
End Sub

Private Sub Mnu_Last_Clipboard_Timer_Click()
'Mnu_Last_Clipboard_Timer
End Sub

Private Sub MNU_LAST_GRAB_CAPS_TOP_BAR_Click()

    Call MNU_LAST_GRAB_CAPS_Click
    Beep
End Sub

Private Sub MNU_LAST_GRAB_LOW_CASE_TOP_BAR_Click()

'MNU_LAST_GRAB_LOW_CASE_TOP_BAR
AD$ = LCase(AD$)

    EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        Clipboard.SetText AD$
    EXECUTE_TIMER_ENABLED = True

Me.WindowState = vbMinimized
AD_DATE = 0
Beep

End Sub

Private Sub MNU_MIRROR_EXE_DRY_RUN_Click()
    
    SIMU_TEST = True
    'Call VB_PROJECT_CHECKDATE
    
    Exit Sub
    
    'Shell App.Path + "\Network_Sync_EXE_Clipboard.exe """ + App.Path + "\" + App.EXEName + ".EXE" + """", vbNormalFocus
    Shell "D:\VB6\VB-NT\00_Best_VB_01\RELOAD_NETWORK_SYNC_EXE\RELOAD_Network_Sync_EXE_VB_MIRROR.exe """ + App.Path + "\" + App.EXEName + ".EXE" + """", vbNormalFocus
    
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub

End Sub

Private Sub MNU_MIRROR_SEND_TO_OPERATING_SYSTEM_Click()

'--------------------------
'JOB # 01
'WORK TO FINISH ON THE SEND TO SPECIAL FOLDER
'2016 MAY 02 MON BANK HOLIDAY
'JOB FINISH 2016 MAY 09 MON
'LEFT TO DO WIN 10 FOLDER
'AND MAKE DO SUBFOLDER
'------------------------
'--------------------------
'ALSO JOB BEFORE STOP FLICKER OF MENU WHEN HOVER OVER TEXT BOX
'MENU ITEM WHERE UPDATE WITH MOUSE DETECT MOVE - STORE IN VAR VAULE
'--------------------------

FOLDER_SENDTO_SYSTEM = GetSpecialfolder(CSIDL_SENDTO) + "\"

'----------------
'FIND OUR FAT32 DRIVE BY SEARCH FOLDER SENDTO FOLDER
'----------------
'DEFAULT XP
Dim FOLDER_FIND As String
FOLDER_FIND = "\01 SendTo\"
'----------------
'SIX = WIN 7
'----------------
If GETWinNT_Ver = "WIN XP" Then FOLDER_FIND = "\01 SendTo 01- WIN XP"
If GETWinNT_Ver = "WIN 7" Then FOLDER_FIND = "\01 SendTo 02- WIN 7"
If GETWinNT_Ver = "WIN 7 - 64" Then FOLDER_FIND = "\01 SendTo 02- WIN 7"
If GETWinNT_Ver = "WIN 10" Then FOLDER_FIND = "\01 SendTo 03- WIN 10"
'----------------

'HEAVY TASK - SWITCH TO PROBLEM CAN HAPPEN
'If GETWinNT_Ver = "WIN 7" Then
'    If GetOsBitness = 64 Then FOLDER_FIND = "\01 SendTo 02- WIN 7 - 64"
'End If
''GetCpuBitness

'BETTER - BUT STILL USE FEW CPU CYCLES
If GETWinNT_Ver = "WIN 7" Then
    If GetOsArchitecture = 64 Then FOLDER_FIND = "\01 SendTo 02- WIN 7 - 64"
End If

FOLDER_SENDTO_FAT32_STORE = GET_DRIVES_FIND_FOLDER(FOLDER_FIND)

If FOLDER_SENDTO_FAT32_STORE = "" Then
    MsgBox "SEND TO SPECIAL SYSTEM FOLDER LOCATION STORE FAT32 FOLDER - NOT FOUND"
    Exit Sub
End If

ScanPath.TxtPath = FOLDER_SENDTO_FAT32_STORE

ScanPath.cboMask = "*.*"
ScanPath.chkSubFolders = vbUnchecked
ScanPath.ListView1.ListItems.Clear
Call ScanPath.cmdScan_Click

XCOUNT2 = ScanPath.ListView1.ListItems.Count
For WE = 1 To ScanPath.ListView1.ListItems.Count
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    
    If (FS.FileExists(A1 + B1) And FS.FileExists(FOLDER_SENDTO_SYSTEM + B1)) = False Then
        FS.CopyFile A1 + B1, FOLDER_SENDTO_SYSTEM + B1
        FILE_COPY_COUNT = FILE_COPY_COUNT + 1
    End If
Next

If FILE_COPY_COUNT > 0 Then
    MsgBox "FILE COPY COUNT TO SYSTEM SPECIAL FOLDER SEND_TO EQUAL =" + Str(FILE_COPY_COUNT)
Else
    MsgBox "NOT ANY FILE COPY COUNT" + vbCrLf + "ALL FILE MUST ALREADY EXIST AT DESTINATION OF SPECIAL SYSTEM FOLDER"
End If

'GetSpecialfolder(CSIDL_PERSONAL)
'GetSpecialfolder(CSIDL_PROGRAMS)

'Private Const CSIDL_PROGRAMS = &H2
'Private Const CSIDL_CONTROLS = &H3
'Private Const CSIDL_PRINTERS = &H4
'Private Const CSIDL_PERSONAL = &H5
'Private Const CSIDL_FAVORITES = &H6
'Private Const CSIDL_STARTUP = &H7
'Private Const CSIDL_RECENT = &H8
'Private Const CSIDL_SENDTO = &H9

End Sub

Sub MNU_NIRSOFT_SETUP()

ReDim I5(80, 2)

'---------------------------------------------------------------------
'USE SYSTEM OF UPPERCASE ONE CHARACTOR AND
'AFTER THE UPPER CASE -- WILL SPREAD THE WORD OUT
'BY GIVE EXTRA SPACE
'---------------------------------------------------------------------

I3_32 = "C:\PStart\Progs\0_Nirsoft_Package\NirSoft\"
I3_64 = "C:\PStart\Progs\0_Nirsoft_Package\NirSoft\x64\"
I3 = I3_32
I4 = 0
I4 = I4 + 1: I5(I4, 1) = I3 + "NirLauncher.exe"
I4 = I4 + 1: I5(I4, 1) = I3 + "----"
I4 = I4 + 1: I5(I4, 1) = I3 + "TurnedOnTimesView.exe"
I4 = I4 + 1: I5(I4, 1) = I3 + "----"
I4 = I4 + 1: I5(I4, 1) = I3 + "OpenedFilesView.exe"
I4 = I4 + 1: I5(I4, 1) = I3 + "RegScanner.exe"
I4 = I4 + 1: I5(I4, 1) = I3 + "SpecialFoldersView.exe"
I4 = I4 + 1: I5(I4, 1) = I3 + "ResourcesExtract"
I4 = I4 + 1: I5(I4, 1) = I3 + "RegScanner.exe"
I4 = I4 + 1: I5(I4, 1) = I3 + "----"
I4 = I4 + 1: I5(I4, 1) = I3 + "AdapterWatch.exe"
I4 = I4 + 1: I5(I4, 1) = I3 + "NetworkInterfacesView.exe"
I4 = I4 + 1: I5(I4, 1) = I3 + "WirelessNetworkWatcher.exe"
I4 = I4 + 1: I5(I4, 1) = I3 + "WirelessNetView.exe"
I4 = I4 + 1: I5(I4, 1) = I3 + "----"
I4 = I4 + 1: I5(I4, 1) = I3 + "ExecutedProgramsList"
I4 = I4 + 1: I5(I4, 1) = I3 + "LastActivityView"
I4 = I4 + 1: I5(I4, 1) = I3 + "WhatIsHang"
I4 = I4 + 1: I5(I4, 1) = I3 + "WinCrashReport"
I4 = I4 + 1: I5(I4, 1) = I3 + "WinLogOnView"
I4 = I4 + 1: I5(I4, 1) = I3 + "Wireless Network Watcher"
I4 = I4 + 1: I5(I4, 1) = I3 + "WirelessNetView"
I4 = I4 + 1: I5(I4, 1) = I3 + "----"
I4 = I4 + 1: I5(I4, 1) = I3 + "MetarWeather"
I4 = I4 + 1: I5(I4, 1) = I3 + "MyUnInst"
I4 = I4 + 1: I5(I4, 1) = I3 + ""
I4 = I4 + 1: I5(I4, 1) = I3 + ""
I4 = I4 + 1: I5(I4, 1) = I3 + ""



I_TOTAL = I4

For I4 = 1 To I_TOTAL
    If I5(I4, 1) <> "" Then
        If InStrRev(UCase(I5(I4, 1)), ".EXE") = 0 Then
            If InStr(I5(I4, 1), "----") = 0 Then
                I5(I4, 1) = I5(I4, 1) + ".exe"
            End If
        End If
    End If
Next

For I4 = 2 To I_TOTAL
    If I5(I4, 1) <> "" Then
        If InStrRev(UCase(I5(I4, 1)), ".EXE") > 0 Then
            PATH_TEST = Replace(I5(I4, 1), I3_32, I3_64)
            If Dir(PATH_TEST) <> "" Then
                I5(I4, 1) = PATH_TEST
            End If
        End If
    End If
Next

For I4 = 1 To I_TOTAL
    If I5(I4, 1) <> "" And I5(I4, 2) = "" And InStr(I5(I4, 1), "----") = 0 Then
        I7 = Mid(I5(I4, 1), InStrRev(I5(I4, 1), "\") + 1)
        I9 = ""
        For I8 = 1 To Len(I7) - 4
            If UCase(Mid(I7, I8, 1)) = Mid(I7, I8, 1) Then I9 = I9 + " "
            I9 = I9 + Mid(I7, I8, 1)
        Next I8
        
        I9 = Trim(I9 + Mid(I5(I4, 1), InStrRev(I5(I4, 1), ".")))
        
        I5(I4, 2) = I3 + " -- " + I9
    End If
Next

For I4 = 1 To I_TOTAL
    If I5(I4, 1) <> "" Then
        If I21 < Len(I5(I4, 1)) Then I21 = Len(I5(I4, 1))
    End If
Next

For I4 = 1 To I_TOTAL
    If InStr(I5(I4, 1), "----") > 0 Then
        I23 = String(I21, "-")
        I5(I4, 2) = I23
    End If
Next

I4 = 0
For Each Control In MNU_NIRSOFT
    I4 = I4 + 1
    If I5(I4, 1) <> "" Then
        Control.Caption = I5(I4, 2)
        Control.Visible = True
        If InStr(I5(I4, 1), I3_64) > 0 Then Control.Caption = Control.Caption + " ---- X64"
    Else
        Control.Visible = False
    End If
Next

End Sub

Private Sub MNU_MULTI_MENU_Click()
Beep
Shell "D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe", vbNormalFocus
Me.WindowState = vbMinimized

End Sub

Private Sub MNU_NIRSOFT_Click(Index As Integer)

'SEARCH HERE TO ADD ENTRY WANTING
'--------------------------------
'MNU_NIRSOFT_SETUP
'--------------------------------

'THE ARRAY LOADED HERE
'SUB MNU_NIRSOFT_SETUP()
Beep

EXEPATH = I5(Index, 1)

If Mid(EXEPATH, 1, 4) = "----" Then Exit Sub

If InStr(EXEPATH, "NirLauncher.exe") > 0 Then
    EXEPATH = Replace(LCase(EXEPATH), LCase("C:\PStart\Progs\0_Nirsoft_Package\NirSoft"), "C:\PStart\Progs\0_Nirsoft_Package")
    '-----------------------------------------
    'DROP DOWN A LEVEL FOR THE NirLauncher.exe
    'ALL THE PROGRAM CODE ARE INA SUB FOLDER
    '-----------------------------------------
End If

If Dir(EXEPATH) <> "" Then
    Me.WindowState = vbMinimized
    Shell EXEPATH, vbNormalFocus
    Beep
    
    Exit Sub
End If

MsgBox "FIND PROGRAM ATTEMPT" + vbCrLf + vbCrLf + EXEPATH + vbCrLf + vbCrLf + "NOT FOUND"

End Sub


Private Sub MNU_NOTEPAD___Click()
Beep
If Dir("C:\Program Files (x86)\Notepad++\notepad++.exe") <> "" Then
    Shell "C:\Program Files (x86)\Notepad++\notepad++.exe", vbNormalFocus
    Exit Sub
End If
If Dir("C:\Program Files\Notepad++\notepad++.exe") <> "" Then
    Shell "C:\Program Files\Notepad++\notepad++.exe", vbNormalFocus
    Exit Sub
End If

MsgBox "FIND PROGRAM ATTEMPT IN PROGRAM FILES AND Program Files (x86)" + vbCrLf + "NOT FOUND"

End Sub

Private Sub Mnu_Play_Sound_1_Click()
            
    MMControl1.Command = "prev"
    MMControl1.Command = "Play"
            
    If Mnu_SoundOn.Checked = False Then
        MsgBox "Play Sound in Program Running Set of Off", vbMsgBoxSetForeground
    End If
End Sub

Private Sub Mnu_Play_Sound_2_Click()
    MMControl2.Command = "prev"
    MMControl2.Command = "Play"
            
    If Mnu_SoundOn.Checked = False Then
        MsgBox "Play Sound in Program Running Set of Off", vbMsgBoxSetForeground
    End If
End Sub

Private Sub MNU_PROCAPS_TOP_BAR_Click()

Call MNU_PROCAPS_Click
Beep



End Sub

Private Sub Mnu_Project_Source_Code_Click()
    
    iMessageResultRECompile = False
    
    Call Timer3_Timer

End Sub

Private Sub MNU_PROPER_CAPS_TOP_BAR_Click()
Beep
Call MNU_FORMAT_MINE_Click
End Sub

Private Sub MNU_REFORMAT_ADD_A_DASH_Click()
Beep
Dim EE As String

EE = AD$

EE = Replace(EE, vbCrLf, " --" + vbCrLf)

AD$ = EE


    EXECUTE_TIMER_ENABLED = False
    Clipboard.Clear
    Clipboard.SetText AD$
    EXECUTE_TIMER_ENABLED = True



End Sub


Private Sub MNU_REFORMAT_REMOVE_THE_DASH_Click()
Beep
Dim EE As String

EE = AD$

EE = Replace(EE, " --" + vbCrLf, vbCrLf)

AD$ = EE


    EXECUTE_TIMER_ENABLED = False
    Clipboard.Clear
    Clipboard.SetText AD$
    EXECUTE_TIMER_ENABLED = True



End Sub




Private Sub MNU_REG_JUMP_Click()

'1..MNU_CLIPBOARD_EXPLORER_FILE_FOLDER_Click
'2..MNU_REG_JUMP_Click
'3..MNU_URL_Browser_Click
'4..MNU_CPC_Click

    Beep

    TEXT_PATH = Clipboard.GetText
    XPOS = 0
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HKEY_")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HCR\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HCU\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HKCU\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HKLM\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HKU\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HU\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HCC\"): REG_KEY = True
    
    TYPE_VAR = "PATH, URL OR REG_KEY"
    TYPE_VAR = "FOR USE WITH CPC"
    TYPE_VAR = "REG_KEY"
    
    If XPOS = 0 Then
        I = MsgBox("LAST CLIPBOARD NOT CONTAIN A VERIFIABLE STRING " + TYPE_VAR + " TO LOAD" + vbCrLf + vbCrLf + TEXT_PATH + vbCrLf + vbCrLf + "----" + vbCrLf + vbCrLf + "WANT MINIMIZE ", vbYesNo, vbMsgBoxSetForeground)
        If I = vbYes Then Me.WindowState = vbMinimized
        Exit Sub
    End If
   
    If XPOS > 1 Then
        TEXT_PATH = Mid(TEXT_PATH, XPOS, InStr(XPOS, TEXT_PATH, vbCr) - XPOS)
    End If
    If XPOS = 1 And InStr(XPOS, URL_PATH, vbCr) > 0 Then
        TEXT_PATH = Mid(TEXT_PATH, XPOS, InStr(XPOS, TEXT_PATH, vbCr) - XPOS)
    End If
    
    Me.WindowState = vbMinimized
    FileSpec = TEXT_PATH
    
    If Dir("C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\REGJUMP_CLIP_BOARD.BAT") = "" Then
        MsgBox "PROGRAM EXPECTING NOT FOUND AT THIS LOCATION TO REG JUMP" + vbCrLf + "C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\REGJUMP_CLIP_BOARD.BAT"
        Exit Sub
    End If
    
    Shell "C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\REGJUMP_CLIP_BOARD.BAT", vbNormalNoFocus
    'Shell "C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\regjump.exe -c", vbNormalFocus

    

End Sub

Private Sub Mnu_Reload_Audio_Click()

'DO THIS LATER
'FIND BUG WONT LOAD PROPER AFTER FIRST

Call SETUP_SOUND_FILE("")

End Sub

Private Sub MNU_RENAME_MP3_TAG_MUSIC_FOLDER_Click()

'Call FILE_OPERATIONS_HANDLER

Dim I As Long

Beep

'------------------------------
'read only problem in scan path
'must be set
'------------------------------
'not find all without
'------------------------------
'NOT THAT PROBLEM
'FIULE COUNT SHOULD BE 50K
'NOT 19K
'MISSING ITEM
'------------------------------




Do
    If ScanPath.ListView1.ListItems.Count > 0 Then
        MSGBOX_OPTION = vbYesNoCancel  'vbRetryABORTCancel
        If IsIDE = False Then MSGBOX_OPTION = vbYesNo 'vbRetryOnly + vbAbortONLY
        I = MsgBox("SCAN PATH -- IS OCCUPIED WITH ANOTHER TASK" + vbCrLf + vbCrLf + "RETRY", MSGBOX_OPTION, vbMsgBoxSetForeground)
        If I = vbCancel Then
            Stop
            Exit Sub
        End If
        Exit Sub
        If I = vbNo Then
            Exit Sub
        End If
    End If
Loop Until ScanPath.ListView1.ListItems.Count = 0


'-----------------------------------
'ONLY USE IF WANT EXTRA DATE INFO ADDED
'-----------------------------------
ScanPath.chkCopyMemory.Value = vbUnchecked
'-----------------------------------
'-----------------------------------
'ONLY EXTENSION NOT FILTER FILE NAME
'-----------------------------------
ScanPath.cboMask = "*.MP3"
ScanPath.chkSubFolders = vbChecked
ScanPath.ListView1.ListItems.Clear
'-----------------------------
'Case 4 --- sort of PATH then FILE
SORTI = 4
'---------------------------------------------------------
'USE OF CLEAR AND NOT TO DO ADD TO SCANPATH SCAN IS WORKER
'---------------------------------------------------------
ScanPath.ListView1.ListItems.Clear
'ScanPath.ListView1.ListItems.Count

MNU_F_O_M.Caption = "* FILE OPERATIONS MENU *" + " SCANNING BEGIN *"
Me.Refresh

ScanPath.TxtPath = "D:\0 00 MUSIC ---"
Call ScanPath.cmdScan_Click
I = ScanPath.ListView1.ListItems.Count
MNU_F_O_M.Caption = "* FILE OPERATIONS MENU * SCANNING" + Str(I)
Beep
Me.Refresh
GoTo JUMP22442
'MsgBox I '19012
'----------------
ScanPath.cboMask = "*.MP3"
ScanPath.TxtPath = "D:\0 00 MUSIC -Z0x"
Call ScanPath.cmdScan_Click
I = ScanPath.ListView1.ListItems.Count
MNU_F_O_M.Caption = "* FILE OPERATIONS MENU * SCANNING" + Str(I)
Beep
Me.Refresh
'MsgBox I
'----------------
ScanPath.TxtPath = "\\4-asus-gl522vw\4-asus-gl522vw_02_d-drive\0 00 MUSIC ---"
Call ScanPath.cmdScan_Click
I = ScanPath.ListView1.ListItems.Count
MNU_F_O_M.Caption = "* FILE OPERATIONS MENU * SCANNING" + Str(I)
Beep
Me.Refresh
'MsgBox I
'----------------
ScanPath.TxtPath = "\\4-asus-gl522vw\4-asus-gl522vw_02_d-drive\0 00 MUSIC -Z0x"
Call ScanPath.cmdScan_Click
I = ScanPath.ListView1.ListItems.Count
MNU_F_O_M.Caption = "* FILE OPERATIONS MENU * SCANNING" + Str(I)
Beep
Me.Refresh
'MsgBox I
'----------------


JUMP22442:

'----------------
If ScanPath.ListView1.ListItems.Count = 0 Then
    MSGBOX_OPTION = vbOKCancel
    If IsIDE = False Then MSGBOX_OPTION = vbOKOnly
    I = MsgBox("SCAN PATH -- RESULT RETURNED ** EMPTY ** NOT SUCCESSFUL " + vbCrLf + vbCrLf + "ScanPath.ListView1.ListItems.Count = 0" + vbCrLf + vbCrLf + "REQUEST TO SCAN WAS EMPTY VARIABLE" + vbCrLf + vbCrLf + "EXIT RUN", MSGBOX_OPTION, vbMsgBoxSetForeground)
    If I = vbCancel Then
        Stop
        Exit Sub
    End If
    Exit Sub
End If

'----
'---- GO
'----

'D:\0 00 MUSIC ---\04 My Music\01 Banging Tunes\New\New 2009\00 Had To Modify But Okay

'Do

X10 = False
Debug.Print "------------------------------------------------------------"
Debug.Print "------------------------------------------------------------"
Debug.Print "------------------------------------------------------------"
Debug.Print "------------------------------------------------------------"
Debug.Print "------------------------------------------------------------"

TEST_TEXT_FILE = "C:\TEMP\TEST_APPEND_MP3_FILE_SCANNER.TXT"
If Dir(TEST_TEXT_FILE) <> "" Then Kill TEST_TEXT_FILE

STRING_CHECK = "0#00"
For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    
    If A1 + B1 = "D:\0 00 MUSIC ---\04 My Music\01 Banging Tunes\New\New 2009 01 Auto DONE\2008 - Dj Jademan\0#00--------------------------^.mp3" Then
        Stop
    End If
    
    If InStr(A1, "D:\0 00 MUSIC ---\04 My Music\01 Banging Tunes\New\New 2009\00 Had To Modify But Okay") > 0 Then
        Stop
    End If
    
    If InStr(A1, "D:\0 00 MUSIC ---\04 My Music\01 Banging Tunes\New\New 2009") > 0 Then
        'Stop
    End If
    
    If InStr(A1, "D:\0 00 MUSIC ---\04 My Music\01 Banging Tunes\New\New 2009") > 0 Then
        'If X10 = False Then Stop
        X10 = True
    End If
    If InStr(A1, "D:\0 00 MUSIC ---\04 My Music\01 Banging Tunes\New\New 2009") = 0 Then
        'If X10 = False Then Stop
        X10 = False
    End If

    '-------------------------------------------------------------------------
    'CAN'T READ IT ALL IN ON 64 BIT
    'SOME FILE SCANNER ARE ALL ABLE TO
    'PROBLEM IN THIS AREA
    'TO DO WITH STUCK AT FOLDER AND NOT READ FILE NAME OR DIR ONE WAY OR OTHER
    '-------------------------------------------------------------------------
    If X10 = True Then
        Debug.Print A1 + B1
        FR1 = FreeFile
        Open TEST_TEXT_FILE For Append As FR1
            WRITE_SOME_OUTPUT_RESULT_HAPPEN = True
            Print #FR1, A1 + B1
        Close #FR1
    End If
    
    'GoTo JUMP2244
    
    'If Mid(B1, 1, 4) <> STRING_CHECK Then
    'End If

    If WE Mod 1000 = 0 Then
        MNU_F_O_M.Caption = "* FILE OPERATIONS MENU * PROCESS COUNT " + Str(WE) + " of" + Str(ScanPath.ListView1.ListItems.Count)
        Me.Refresh: DoEvents: Beep
    End If

    
    If Mid(B1, 1, 4) = STRING_CHECK Then
        COUNT_CHECK = COUNT_CHECK + 1
    End If



    If Mid(B1, 1, 4) = STRING_CHECK Then 'And 1 = Test + 1 Then
        For WE2 = 1 To ScanPath.ListView1.ListItems.Count
            If EXIT_TRUE = True Then Exit Sub
        'MNU_F_O_M.Caption = "* FILE OPERATIONS MENU * PROCESS COUNT " + Str(WE) + " -- " + Str(WE2)
            
            
            
            A12 = ScanPath.ListView1.ListItems.Item(WE2).SubItems(1)
            B12 = ScanPath.ListView1.ListItems.Item(WE2)
            C12 = ScanPath.ListView1.ListItems.Item(WE2).SubItems(2)
            If A1 = A12 Then
                If Mid(B12, 1, 4) = STRING_CHECK Then
                    If Len(B12) > Len(B1) Then
                        'Stop
                        
                        PATH_SHORT = GetShortName(A12 + C12)
                        If FS.FileExists(PATH_SHORT) = True Then
                            On Error Resume Next
                            Err.Clear
                            FS.DeleteFile PATH_SHORT
                            If Err.Number > 0 Then
                                MsgBox A12 + C12 + vbCrLf + vbCrLf + Err.Description
                            End If
                            Beep
                            On Error GoTo 0
                        End If
                        
                    End If
                End If
            End If
        Next
    End If
    



JUMP2244:

Next
'Loop Until 1 = 2



MsgBox "COUNT RESULT FOR ---- " + vbCrLf + vbCrLf + STRING_CHECK + vbCrLf + vbCrLf + Str(COUNT_CHECK)

If WRITE_SOME_OUTPUT_RESULT_HAPPEN = True Then
    MsgBox "LOAD NOTE PAD LOOK AT SOME WRITING RESULT", vbMsgBoxSetForeground
    Shell "NOTEPAD """ + TEST_TEXT_FILE + """", vbNormalFocus
End If

If ScanPath.cboMask <> "*.MP3" Then
    MsgBox " SCAN PATH EXTENSION MASK WAS CHANGED BEFORE END THIS SUB PROCEDURE TO" + vbCrLf + vbCrLf + ScanPath.cboMask
End If

Beep
MNU_F_O_M.Caption = "* FILE OPERATIONS MENU * DONE AT COUNT " + Str(ScanPath.ListView1.ListItems.Count)
ScanPath.ListView1.ListItems.Clear

End Sub

Private Sub Mnu_Reset_MMControl_Click()

'    MMControl1.Command = "Stop"
'    MMControl1.Command = "Close"
'    DoEvents
'    MMControl2.Command = "Stop"
'    MMControl2.Command = "Close"
'    DoEvents



'THIS GET RUN HERE AT RESET
'AND AT ONE CLICK SELECT


'-----------------------------
'MODIFY_SOUND_SELECTION = True
'Call SETUP_SOUND_FILE("")
'-----------------------------
Call RESET_SETUP_SOUND_FILE("")



Exit Sub


If Path_Of_Sound_File <> "" Then
    
'    MMControl1.Notify = False
'    MMControl1.Wait = True
    MMControl1.Command = "Stop"
'    MMControl1.fileName = ""

    MMControl1.Command = "Close"
    DoEvents
    MMControl1.Notify = True
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
    
    MMControl1.fileName = Path_Of_Sound_File
    
    MMControl1.Command = "Open"

'    Exit Sub
End If

If Path_Of_Sound_File = "" Then
    MsgBox "Sound File Path Variable Not Set", vbMsgBoxSetForeground
    'Exit Sub
End If

If Path_Of_Sound_File <> "" And Dir(Path_Of_Sound_File) = "" Then
    MsgBox "Sound File - Path and File Not Found" + vbCrLf + Path_Of_Sound_File, vbMsgBoxSetForeground
    'Exit Sub
End If



vPathSOUND2 = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\"
If Dir(vPathSOUND2, vbDirectory) = "" Then
    MkDirNested vPathSOUND2
End If

vFileSOUND2 = Dir(vPathSOUND2 + "*.WAV")
'vFileSOUND2

If vFileSOUND2 = "" Then
    MsgBox "Sound File *2* PATH AND OR FILENAME NOT SET " + vbCrLf + App.Path + "\Sound_Wav's--2\" + GetComputerName + "\*.WAV", vbMsgBoxSetForeground
    'Exit Sub
End If

If vFileSOUND2 <> "" Then
    MNU_SOUND_2.Caption = "SOUND OPTION - 2 ------ \" + vFileSOUND2
    
    MMControl2.Command = "Stop"
    MMControl2.Command = "Close"
    '---- Set properties needed by MCI to open.
    MMControl2.Notify = True
    MMControl2.Wait = True
    MMControl2.Shareable = False
    MMControl2.DeviceType = "WaveAudio"
    MMControl2.fileName = vPathSOUND2 + vFileSOUND2
    ' Open the MCI WaveAudio device.
    MMControl2.Command = "Open"
       
    'MMControl2.Command = "prev"
    
    'MMControl2.Command = "Play"
    
'    MNU_SOUND_2.Caption = "SOUND OPTION - 2 - " + vFileSOUND2


Else
    MNU_SOUND_2.Caption = "SOUND OPTION - 2 - WAV File Not Found - 1st Found Here Used " + vPathSOUND2 + "*.WAV"

End If



End Sub

Private Sub MNU_SCANPATH_COUNTER_Click()

    MsgBox "This Shows a Count When Scanning Files When Find Last Date Image To Show", vbMsgBoxSetForeground

End Sub













Private Sub MNU_SCROLL_BOTTOM_OFF__WANT_ON_ASK_PRESS_Click()
Beep
Call MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_Click

End Sub

Public Sub Mnu_Show_Error_Script_Frm_Msgbox_Click()
    On Error Resume Next
    '** Show Error Script **
    'DEBUG_HERE
    Beep
    FRM_MSGBOX.Show
End Sub

Private Sub MNU_SHOW_IMAGE_Click()
    
    'Beep
    'Call MNU_LAST_ART_PIC_IVIEW_Click
    'Me.WindowState = vbMinimized
    
    Beep
    
    Dim FS, F, f1, fc, S
    Dim STORE_DATE
    
    folderspec = "D:\0 00 Art Loggers\# APP AND SCREEN\7-ASUS-GL522VW\HOT_KEY_SCREEN_SHOT"
    
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set F = FS.GetFolder(folderspec)
    Set fc = F.SubFolders
    For Each f1 In fc
'        s = s & f1.Name
        If f1.DateLastModified > STORE_DATE Then STORE_NAME = f1.Name
    Next
    
    If IsIDE = True Then Stop
    '-------------------------------------------
    'WORK TO DO GATHER WHOLE FILE AND SUB-FOLDER
    'NOT JUST SUB FOLDER AT THE MOMENT
    '-------------------------------------------
    Exit Sub
    
    Beep
    If LAST_CLIPBOARD_FILE = "" Then Exit Sub
    If Dir(LAST_CLIPBOARD_FILE) = "" Then Exit Sub
    Shell "Explorer.exe /SELECT, """ + LAST_CLIPBOARD_FILE + """", vbMaximizedFocus
    Beep
End Sub

Private Sub MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC_Click()
    
    'GetSpecialfolder_DEBUG
    
    STRING_VAR = "DESKTOP PUBLIC -- "
    '------------
    FOLDER_SYSTEM = GetSpecialfolder(25) + "\"
    
    MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC.Caption = STRING_VAR + FOLDER_SYSTEM
    
    If VAR_FLAG_EXPLORER_LOAD_NOT = False Then
        Shell "Explorer.exe " + FOLDER_SYSTEM, vbNormalFocus
    End If
End Sub


'Private Sub MNU_SPECIAL_FOLDER_PIN_TO_START_MENU_Click()
'
'Call SPECIAL_FOLDER_PIN_TO_START_MENU
'
'Shell "Explorer.exe " + FOLDER_PIN_TO_START_MENU, vbNormalFocus
'
'End Sub

Private Sub MNU_SPECIAL_FOLDER_DESKTOP_USER_Click()
    
    STRING_VAR = "DESKTOP USER -- "
    '------------
    FOLDER_SYSTEM = GetSpecialfolder(CSIDL_DESKTOP) + "\"
    'REPLACE ----
    'FOLDER_SYSTEM = Replace(FOLDER_SYSTEM, "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs", "\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu")
    '------------
    'FOLDER_PIN_TO_START_MENU = FOLDER_SYSTEM
    'Debug.Print FOLDER_SYSTEM
    MNU_SPECIAL_FOLDER_DESKTOP_USER.Caption = STRING_VAR + FOLDER_SYSTEM
    
    If VAR_FLAG_EXPLORER_LOAD_NOT = False Then
        Shell "Explorer.exe " + FOLDER_SYSTEM, vbNormalFocus
    End If
End Sub

Private Sub MNU_SPECIAL_FOLDER_PIN_TO_START_MENU_Click()
    '------------
    FOLDER_SYSTEM = GetSpecialfolder(CSIDL_PROGRAMS) + "\"
    'REPLACE ----
    FOLDER_SYSTEM = Replace(FOLDER_SYSTEM, "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs", "\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu")
    '------------
    
    FOLDER_PIN_TO_START_MENU = FOLDER_SYSTEM
    'Debug.Print FOLDER_SYSTEM
    MNU_SPECIAL_FOLDER_PIN_TO_START_MENU.Caption = "START MENU -- PIN TO -- " + FOLDER_SYSTEM
    
    If VAR_FLAG_EXPLORER_LOAD_NOT = False Then
        Shell "Explorer.exe " + FOLDER_SYSTEM, vbNormalFocus
    End If
End Sub


Private Sub SPECIAL_FOLDER_PIN_TO_START_MENU()


'Private Const CSIDL_PROGRAMS = &H2
FOLDER_SYSTEM = GetSpecialfolder(CSIDL_PROGRAMS) + "\"
'C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\Start Menu\Programs

FOLDER_SYSTEM = Replace(FOLDER_SYSTEM, "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs", "\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu")
FOLDER_PIN_TO_START_MENU = FOLDER_SYSTEM
'Debug.Print FOLDER_SYSTEM
'C:\Users\MATT 01\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu\

MNU_SPECIAL_FOLDER_PIN_TO_START_MENU.Caption = "EXPLORER -- PIN TO START MENU -- " + FOLDER_SYSTEM

'MNU_PIN_TO_START_MENU
'C:\Users\MATT 01\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu
End Sub

       
        


Sub SPECIAL_FOLDER_SEND_TO()

FOLDER_SENDTO_SYSTEM = GetSpecialfolder(CSIDL_SENDTO) + "\"
'--------
'01 OF 02
'--------
MNU_SEND_TO_SYSTEM_FOLDER.Caption = "SEND TO -- " + FOLDER_SENDTO_SYSTEM

'----------------
'FIND OUR FAT32 DRIVE BY SEARCH FOLDER SENDTO FOLDER
'----------------
'DEFAULT XP
Dim FOLDER_FIND As String
FOLDER_FIND = "\01 SendTo\"
'----------------
'SIX = WIN 7
'----------------
If GETWinNT_Ver = "WIN XP" Then FOLDER_FIND = "\01 SendTo 01- WIN XP"
If GETWinNT_Ver = "WIN 7" Then FOLDER_FIND = "\01 SendTo 02- WIN 7"
If GETWinNT_Ver = "WIN 7 - 64" Then FOLDER_FIND = "\01 SendTo 02- WIN 7"
If GETWinNT_Ver = "WIN 10" Then FOLDER_FIND = "\01 SendTo 03- WIN 10"
'----------------

'HEAVY TASK - SWITCH TO PROBLEM CAN HAPPEN
'If GETWinNT_Ver = "WIN 7" Then
'    If GetOsBitness = 64 Then FOLDER_FIND = "\01 SendTo 02- WIN 7 - 64"
'End If
''GetCpuBitness

'BETTER - BUT STILL USE FEW CPU CYCLES
If GETWinNT_Ver = "WIN 7" Then
    If GetOsArchitecture = 64 Then FOLDER_FIND = "\01 SendTo 02- WIN 7 - 64"
End If

'----------------
FOLDER_SENDTO_FAT32_STORE = GET_DRIVES_FIND_FOLDER(FOLDER_FIND)
'--------
'02 OF 02
'--------
MNU_SEND_TO_FAT32_FOLDER.Caption = "SEND TO -- " + FOLDER_SENDTO_FAT32_STORE

End Sub


Private Sub MNU_SEND_TO_SYSTEM_FOLDER_Click()

Call SPECIAL_FOLDER_SEND_TO

Shell "Explorer.exe " + FOLDER_SENDTO_SYSTEM, vbNormalFocus

End Sub

Private Sub MNU_SEND_TO_FAT32_FOLDER_Click()

Call SPECIAL_FOLDER_SEND_TO

Shell "Explorer.exe " + FOLDER_SENDTO_FAT32_STORE, vbNormalFocus

End Sub


Private Sub MNU_TEST_STOP_ALL_TIMER_Click()


Exit Sub

'ENABLED RESTORE
If MNU_TEST_STOP_ALL_TIMER.Checked = True And 1 = 2 Then
    MNU_TEST_STOP_ALL_TIMER.Checked = False
    
    TTCOUNT = 0
    Dim Control
    'SET ALL TIMERS IN ALL FORMS ENABLED=FALSE
    On Error Resume Next
        For I = 0 To Forms.Count - 1
            For Each Control In Forms(I).Controls
                If InStr(UCase(Control.Name), "TIMER") > 0 Then
                    TTCOUNT = TTCOUNT + 1
                    If TIMER_CODE_TEST(TTCOUNT) = True Then
                        'Debug.Print Control.Name
                        Control.Enabled = True
                    End If
                End If
            Next
        Next I
    On Error GoTo 0
    
    Exit Sub
End If


Exit Sub

'ENABLED = FALSE GLOBAL
If MNU_TEST_STOP_ALL_TIMER.Checked = False And 1 = 1 Then
    MNU_TEST_STOP_ALL_TIMER.Checked = True
    
    For RT2 = 1 To 200
        Do
            TTCOUNT = 0
            CTRL_ENABLED = False
            'SET ALL TIMERS IN ALL FORMS ENABLED=FALSE
            On Error Resume Next
                For I = 0 To Forms.Count - 1
                    For Each Control In Forms(I).Controls
                        
                        If InStr(UCase(Control.Name), "TIMER") > 0 Then
                            'Control.Name
                            If IsIDE = True Then Debug.Print Control.Name
                            TTCOUNT = TTCOUNT + 1
                            If Control.Enabled = True Then CTRL_ENABLED = True
                            TIMER_CODE_TEST(TTCOUNT) = Control.Enabled
                            'Debug.Print Control.Name
                            Control.Enabled = False
                            If Control.Enabled = True Then CTRL_ENABLED = True
                        End If
                    Next
                Next I
            On Error GoTo 0
            DoEvents
        Loop Until CTRL_ENABLED = False
        Sleep 10
    Next RT2
    Exit Sub
End If


End Sub

Private Sub MNU_TEXT_HEX_2_TEXT_Click()

Dim EE As String
EE = AD$
EE = "30 20 5F 20 4E 20 69"
EE = Replace(EE, " ", "")
EE2 = ""
For r = 1 To Len(EE) Step 2
    EE2 = EE2 + Chr("&H" + (Mid(EE, r, 2)))
Next

AD$ = EE2
EXECUTE_TIMER_ENABLED = False
Clipboard.Clear
Clipboard.SetText AD$
EXECUTE_TIMER_ENABLED = True
Beep
End Sub

Private Sub MNU_TIME_API_FUNCTION_ACCESS_Click()
'VARTEXT = "TIME API SUB FUCTION LAST ACCESSED = " + Format(Timer_API_Test_Logick_Var2, "DD-MM-YYYY HH:MM:SS")
'MNU_TIME_API_FUNCTION_ACCESS.Caption = VARTEXT
End Sub

Private Sub MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_Click()

'Debug.Print MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Caption
'Debug.Print MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Caption
'TOGGLE NOT SCROLL BACK TO BOTTOM WHILE WORKING

MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = Not MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked

If MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = True Then
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = False
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Caption = "SCROLL BACK TO BOTTOM ON TIMER.ENABLED=FALSE"
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Caption = "SCROLL BACK TO BOTTOM ON TIMER.ENABLED=FALSE"
    MNU_SCROLL_BOTTOM_OFF__WANT_ON_ASK_PRESS.Caption = "SCROLL BOTTOM=ON"
Else
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = True
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Caption = "SCROLL BACK TO BOTTOM ON TIMER.ENABLED=TRUE"
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Caption = "SCROLL BACK TO BOTTOM ON TIMER.ENABLED=TRUE"
    MNU_SCROLL_BOTTOM_OFF__WANT_ON_ASK_PRESS.Caption = "SCROLL BOTTOM=NOT"
End If





End Sub

Private Sub MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS_Click()

Debug.Print MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Caption
'TOGGLE NOT SCROLL BACK TO BOTTOM WHILE WORKING
Debug.Print MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Caption
'*--- NOT SCROLL IN OPTION MENU---*

If MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = True Then
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = False
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Caption = "*--- SCROLL BACK DOWN TO BOTTOM ON TIMER.ENABLED=FALSE ---*"
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Caption = "*--- SCROLL BACK DOWN TO BOTTOM ON TIMER.ENABLED=FALSE ---*"
Else
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = True
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Caption = "*--- SCROLL BACK DOWN TO BOTTOM ON TIMER.ENABLED=TRUE ---*"
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Caption = "*--- SCROLL BACK DOWN TO BOTTOM ON TIMER.ENABLED=TRUE ---*"
End If



'MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Enabled = False


End Sub

Private Sub MNU_URL_SCREEN_SCRAPER_Click()


Call LAST_IMAGE("EXPLORER", "D:\0 00 Art Loggers\URL SCREEN SHOT\") '=EXPLORER

Exit Sub

'THIS CAN USE SAME CODE AS LAST SCREEN SHOT MENU


'Dim FDS
'
''FDS = "D:\0 00 Art Loggers\URL SCREEN SHOT\URL SCREENSHOT - 00 BBC WEATHER - FULL SIZE.JPG"
'FDS = "D:\0 00 Art Loggers\URL SCREEN SHOT\"
'TXRS = Dir(FDS + "*.JPG")
'Do
'    TXRS = Dir
'    If TXRS = "" Then Exit Do
'    FileSpec = FDS + TXRS
'Loop Until True = False
'
'If FileSpec = "" Then MsgBox "NOT ANY OF THOSE FILES" + vbCrLf + FDS: Exit Sub
'
''Shell "Explorer.exe /select, " + FDS, vbMinimizedNoFocus
'Shell "Explorer.exe " + FDS, vbNormalFocus

'THIS DONT WORK YET - HAS TO FIND NEWEST FILE IN SUB FOLDER
'If IRFANVIEW_PATH <> "" Then
'    If VAR1 = "IVIEW" Then
'        Shell IRFANVIEW_PATH + FileSpec + """ /fs /silent", vbMaximizedFocus
'    End If
'Else
'    Me.WindowState = vbNormal
'    MsgBox "IRFANVIEW_PATH VAR -- PATH NOT FOUND FOR FILE" + vbCrLf + "NOT INSTALED AT EXPECTED LOCATION " + IRFANVIEW_PATH3 + vbCrLf + "OR" + vbCrLf + IRFANVIEW_PATH2, vbMsgBoxSetForeground
'End If

Me.WindowState = vbMinimized
'-----------------------
LAST_FILE_DATE_PATH_HOT_KEY_SCREENSHOT = ""
'ScanPath.SHOW

'XdATE2 = 0
'ScanPath.chk_LIST_VIEW_SHORT_5 = vbChecked

LAST_FILE_DATE_TIME = DateSerial(100, 1, 1)
'SCAN_PARTMASK = "# APP AND SCREEN\"+GETCOMPUTERNAME+"\CLIP_"
ScanPath.cboMask = "*.JPG"
ScanPath.chkSubFolders = vbChecked
ScanPath.TxtPath = "D:\0 00 Art Loggers\URL SCREEN SHOT\"

Call ScanPath.CMDScan_NO_LIST_FAST_Click
'SCAN_PARTMASK = ""

'FileSpec = LAST_FILE_DATE_PATH
FileSpec = LAST_FILE_DATE_PATH_HOT_KEY_SCREENSHOT

If FileSpec = "" Then MsgBox "NOT ANY OF THOSE FILES" + vbCrLf + ScanPath.TxtPath + "\" + ScanPath.cboMask: Exit Sub

'Me.WindowState = vbMinimized

'If MNU_MESSAGE_BOXES.Checked = False Then
    'MsgBox "FOUND LATEST IMAGE Clipboarded - LOAD Explorer Minimized AS Well as IrFan Maximized To View" + vbCrLf + "FILES FOUND =" + str(tFileCount) + vbCrLf, vbMsgBoxSetForeground
'End If


Shell "Explorer.exe /select, " + FileSpec, vbMinimizedNoFocus

'If IRFANVIEW_PATH <> "" Then
'    If VAR1 = "IVIEW" Then
'        Shell IRFANVIEW_PATH + FileSpec + """ /fs /silent", vbMaximizedFocus
'    End If
'Else
'    Me.WindowState = vbNormal
'    MsgBox "IRFANVIEW_PATH VAR -- PATH NOT FOUND FOR FILE" + vbCrLf + "NOT INSTALED AT EXPECTED LOCATION " + IRFANVIEW_PATH3 + vbCrLf + "OR" + vbCrLf + IRFANVIEW_PATH2, vbMsgBoxSetForeground
'End If





End Sub




Private Sub MNU_VB_FOLDER_Click()
Beep
Call MNU_EXPLORER_ME_VB_Click

End Sub

Private Sub MNU_VB_FOLDER_VB_NT__SEND_TO_EXPLORER_Click()
Beep
Me.WindowState = vbMinimized

Shell "Explorer.exe D:\VB6\VB-NT\00_Send_To", vbNormalFocus

End Sub

Private Sub MNU_VB_FOLDER_VB_NT_EXPLORER_Click()
Beep
Me.WindowState = vbMinimized

Shell "Explorer.exe D:\VB6\VB-NT\", vbNormalFocus

End Sub

Private Sub Text1_SelChange()
    
    
    
    If Text1.SelStart > 1 Then
        TEXTBOX1_SELECTION_CHANGE = Text1.SelStart
    End If
    
'    Text1.Enabled = False
'    Text1.SelStart = 1
'    Text1.Enabled = True
'    Text1.SelStart = Len(Text1)

End Sub

Private Sub Timer_API_OKAY_COLOUR_Timer()

'If IsIDE = True Then KEYBOARD_OR_MOUSE_ACTIVE_3_MIN = Now + 1

If IsIDE = True Then
    If KEYBOARD_OR_MOUSE_ACTIVE_3_MIN > Now Then
        Timer_API_OKAY_COLOUR.Interval = 4000
    Else
        Timer_API_OKAY_COLOUR.Interval = 1
    End If
End If



'Path_Of_Sound_File

'Dim Xcol, Xcol2, MOdColTime As Double
'Call ColorCycle
'Xcol = Timer Mod 7
'
'If Xcol = 0 Then Xcol2 = vbRed
'If Xcol = 1 Then Xcol2 = vbGreen
'If Xcol = 2 Then Xcol2 = vbYellow
'If Xcol = 3 Then Xcol2 = vbBlue
'If Xcol = 4 Then Xcol2 = vbMagenta
'If Xcol = 5 Then Xcol2 = vbCyan
'If Xcol = 6 Then Xcol2 = vbWhite

WTrue1 = WTrue1 + TW1 'TW1
If WTrue1 > 255 Then TW1 = -6: WTrue1 = WTrue1 + TW1
If WTrue1 < 1 Then TW1 = 6: WTrue1 = WTrue1 + TW1

WTrue2 = HTrue2 + TW2
If WTrue2 > 255 Then TW2 = -7: WTrue2 = WTrue2 + TW2
If WTrue2 < 1 Then TW2 = 7: WTrue2 = WTrue2 + TW2

WTrue3 = WTrue3 + TW3
If WTrue3 > 255 Then TW3 = -8: WTrue3 = WTrue3 + TW3
If WTrue3 < 1 Then TW3 = 8: WTrue3 = WTrue3 + TW3
   
'Foreground
'Label4.BackColor = RGB(KWTrue, HWTrue, WTrue)   ' Set drawing color.
'Inverse
'Label4.ForeColor = RGB(255 - KWTrue, 255 - HWTrue, 255 - WTrue)

'frmPassLock.Label4.BackColor = RGB(kwtrue, hwtrue, wtrue)   ' Set drawing color.
'frmPassLock.Label4.ForeColor = RGB(255 - kwtrue, 255 - hwtrue, 255 - wtrue)

'Line1.BorderColor = RGB(WTrue1, WTrue2, WTrue3)   ' Set drawing color.

Label1.BackColor = RGB(WTrue1, WTrue2, WTrue3)

'Label1.ForeColor = RGB(WTrue1, WTrue2, WTrue3)

End Sub




Private Sub Timer_INFORAPID_MSGBOX_Timer()

Timer_INFORAPID_MSGBOX.Enabled = False

MsgBox "IN THE FOLLOWNG INFORAPID APPLICATION TO LOAD" + vbCrLf + "ENTER THE WILDCARD * IN ClipBoard-*.TXT --- DIY" + vbCrLf + "AND SUBDIR SCAN CHECK BOX --- DIY -- OKAY BECUASE I DON'T KNOW", vbMsgBoxSetForeground

End Sub

Private Sub Timer_Idle_Few_Second_Timer()

'----------------------
'also this sub for here
'----------------------

Dim ITXT

If TIME_LAST_CLIPBOARD > 0 Then
    
    ITXT = "Last Clipboard = " + ClipFormatDesc2 + " --"
    ITXT = ITXT + Str(DateDiff("n", TIME_LAST_CLIPBOARD, Now))
    Test_Min_Var = "m"
    If DateDiff("s", TIME_LAST_CLIPBOARD, Now) < 59 Then
        ITXT = ITXT + Str(DateDiff("s", TIME_LAST_CLIPBOARD, Now))
        Test_Min_Var = "s"
    End If
    LAST_CLIP_TIME_INFO = ITXT + Test_Min_Var
    Mnu_LAST_CLIP_TIME_Second.Caption = LAST_CLIP_TIME_INFO + " Away"

End If

If TIME_LAST_CLIPBOARD = 0 And TIME_LAST_CLIPBOARD_Timer1 > 0 Then
    
    ITXT = "Last Clipboard = " + ClipFormatDesc2 + " --"
    ITXT = ITXT + Str(DateDiff("n", TIME_LAST_CLIPBOARD_Timer1, Now)) '+ " Second"
    Test_Min_Var = "m"
    If DateDiff("s", TIME_LAST_CLIPBOARD, Now) < 59 Then
        ITXT = ITXT + Str(DateDiff("s", TIME_LAST_CLIPBOARD_Timer1, Now))
        Test_Min_Var = "s"
    End If
    LAST_CLIP_TIME_INFO = ITXT + Test_Min_Var
    
    Mnu_LAST_CLIP_TIME_Second.Caption = "Last Clipper" + Str(DateDiff("s", TIME_LAST_CLIPBOARD_Timer1, Now)) + " Start-up In Buffer"

End If

'----------------------
If TIME_LAST_CLIPBOARD_ERROR_MSG > 0 Then
    
    ITXT = "Show Error Script && Last Error = "
    ITXT = ITXT + Str(DateDiff("s", TIME_LAST_CLIPBOARD_ERROR_MSG, Now)) + " Second"
    LAST_CLIP_TIME_INFO = ITXT
    Mnu_Show_Error_Script_Frm_Msgbox.Caption = LAST_CLIP_TIME_INFO
Else
    'Mnu_Show_Error_Script_Frm_Msgbox.Visible = False
End If


'THE RESUME FROM IDLE CHECK
If Mouse_Keyboard_Active_Time = 0 Then Exit Sub

'If DateDiff("s", Mouse_Keyboard_Active_Time, Now) < 10 Then Exit Sub
'If Mouse_Keyboard_Active_Time_Before = 0 Then Exit Sub
    
If Mouse_Keyboard_Active_Time_Before > DateDiff("s", Mouse_Keyboard_Active_Time, Now) Then
            
    Timer_MOUSE_5_MINUTE.Enabled = True
    'Debug.Print Str(Mouse_Keyboard_Active_Time_Before) + " -- > " + Str(DateDiff("s", Mouse_Keyboard_Active_Time, Now))

End If

Mouse_Keyboard_Active_Time_Before = DateDiff("s", Mouse_Keyboard_Active_Time, Now)

End Sub





Private Sub Timer_MENU_HEIGHT_CHANGED_Timer()
    
    'Exit Sub
    
    If IsIDE = True Then
        If KEYBOARD_OR_MOUSE_ACTIVE_3_MIN > Now Then
        Timer_MENU_HEIGHT_CHANGED.Interval = 4000
        Else
            Timer_API_OKAY_COLOUR.Interval = 1
        End If
    End If
'    Else
'        Timer_MENU_HEIGHT_CHANGED.Interval = 4000

    
    '--------------------------------------------
    'RESIZE IN SHAPE QUICKER
    '--------------------------------------------
   
    'RE-ENABLED FROM REISE FORM
    If KEYBOARD_OR_MOUSE_ACTIVE_10_SEC < Now Then
        Timer_MENU_HEIGHT_CHANGED.Enabled = False
        Timer_MENU_HEIGHT_CHANGED.Interval = 10
    End If

    'RESIZE IN SHAPE QUICKER
    'Why this happen unless form resize changed
    'SAVE THIS TIMER FROM OVER DOING IT
    If Menu_Height <> OMenu_Height Then
        Call SETUP_DIMENSIONS_INNER_FORM
    End If

    Exit Sub
    '--------------------------------------------
    '--------------------------------------------


    '--------------------------------------------
    'USE HERE SWITCH LATER IMPLIMENT
    '--------------------------------------------
    VAR_USE_EVENT_LESS_QUICK_ISIDE_IS_TRUE = True
    '--------------------------------------------
    'TestLowTimer = True
    
    If KEYBOARD_OR_MOUSE_ACTIVE_3_MIN > Now Then
        If VAR_USE_EVENT_LESS_QUICK_ISIDE_IS_TRUE = True Then
            active1 = True
        End If
    Else
        active1 = False
    End If
    
    If IsIDE = True And active1 = True Then
        Timer_MENU_HEIGHT_CHANGED.Interval = 4000
    Else
        Timer_MENU_HEIGHT_CHANGED.Interval = 500
    End If

End Sub

Private Sub Timer_MOUSE_1_MINUTE_Timer()

Timer_MOUSE_1_MINUTE.Interval = 4000

'RE_ENABLED WITH IDLE ACTIVITY
'AND END AFTER
'WANT TIMER1 - 5 SEC MIN AFTER ACTIVITY END BEGIN
'WANT TIMER2 - 2 MIN AFTER ACTIVITY END
'WANT TIMER3 - 1 MIN INTERVAL WHEN IDLE

Timer_MOUSE_1_MINUTE.Enabled = False

'ONE STEP - TWO NEXT STEP - ONE GIANT LEAP FOR MANKIND
'Timer_MOUSE_2_MINUTE.Enabled = True

MNU_IDLE_ACTIVE.Caption = "IDLE STATE = IDLE"

'Exit Sub

Exit Sub

'NOT TO RUN HERE QUICK LIKE IT USUAL FEW SECOND AFTE IDLE
'NOT TO WHEN RECENT ATIVITY OF CLIPBOARD
If CLIPBOARD_ACTIVITY_MOMENT < Now Then

    'Call Mnu_Reset_MMControl_Click
    Call RESET_SETUP_SOUND_FILE("NOTSOUND")
    Call Mnu_API_Unload_Click
    Call Mnu_API_Reload_Click

End If

'HOPE NOT TOO QUICK AS CALL API SUB LESS QUICK THAN TIMER
'TRY
'100 MSEC ACTIVE ONCE
'HAS A MSGBOX SO WONT TIE UP REST OF CODE
'TEST AD$ VAR HAS BEEN USED BY THE SUB CALL OF API CLIPPER
'---------------------------------------------------------
'FALSE AND THEN ENABLE WITH INTERVAL WORK AS A RESET TIMER
'---------------------------------------------------------
Timer_MOUSE_4_MINUTE.Enabled = False
    Timer_MOUSE_4_MINUTE.Interval = 1000 ' 1 SECOND
Timer_MOUSE_4_MINUTE.Enabled = True



'CALLED FROM TIMER_MOUSE
'--------------------------------
'Major Check to Test API Call in Program

'Clipboard.Clear

'Exit Sub
'
'
'
'Dim ADTEST$
'
'If XArchiveXClipFormatDescription <> ClipFormatDescription Or TimerCheckIntegrityOfProgram = True Then
'
'    'Could Check if Format Has Swapped and Back to Text Again But Contents Don't Match Before Check
'    'Too Much Code Workaround
'
'    If Clipboard.GetFormat(vbCFText) = True And TimerCheckIntegrityOfProgram = True Then
'        TimerCheckIntegrityOfProgram = False
'        If Len(Clipboard.GetText) <= LimitClipSize And Timer_MOUSE.Enabled = False Then
'            ADTEST$ = Clipboard.GetText
'            'If ADTest$ <> AD$ And Len(ADTest$) <= LimitClipSize Then
'            If ADTEST$ <> AD$ Then
'                Me.WindowState = vbNormal
'                Do
'                    DoEvents
'                    Me.Refresh
'                    DoEvents
'                    Sleep 100
'                Loop Until Me.WindowState <> vbMinimized
'
'                FORM_STAY_UP_WITH_MSGBOX = True
'                itech = MsgBox(Format(Now, "DD-MM-YYYY -- HH:MM:SS ") + vbCrLf + "Problem Clipboard Not Logging Text Clipbard Format or Timer Spot Check Detect Clipboard Has Changed but Not Same as Archive Variable" + vbCrLf + "Put the Test Flag Run Option Menu On and Try - And Test Options Menu Unload and Reload the API Form is Next Workaround OPtion", vbMsgBoxSetForeground)
'                FORM_STAY_UP_WITH_MSGBOX = False
'
'            End If
'        End If
'
'
'    End If
'
''    If Clipboard.GetFormat(vbCFText) = True Or Clipboard.GetFormat(vbCFBitmap) = True Then
''        MsgBox "Something Wrong With The Program" + vbCrLf + "Clipboard API Routine Not Logging Clipbard" + vbclrf + "Format Chnaged But Hasn't Been a Logg Capture Entry From the TimeStamp" + vbCrLf + "Put the Test Flag Run Option Menu On and Try" + vbCrLf + "Unload and Reload the API Form is Next Workaround OPtion Report to Programmer", vbMsgBoxSetForeground
'
'
'    End If
'
'
'XArchiveXClipFormatDescription = ClipFormatDescription
'
'
''    If ADTest$ = Clipboard.GetText Then
''        If ADTest$ <> AD$ Then
''            Me.WindowState = vbNormal
''            DoEvents
''            MsgBox "Problem Clipboard Not Logging Clipbard Format or Timer Check Detect Clipboard Has Chnaged but Not Same as Archive Would Be" + vbCrLf + "Put the Test Flag Run Option Menu On and Try - Unload and Reload the API Form is Next Workaround OPtion", vbMsgBoxSetForeground
''
''        End If
''    End If
'
'
''If Clipboard.GetFormat(vbCFText) = True Then
''    If ADTest$ = Clipboard.GetText Then
''        If ADTest$ <> AD$ Then
''            Me.WindowState = vbNormal
''            DoEvents
''            MsgBox "Problem Clipboard Not Logging Clipbard Format or Timer Check Detect Clipboard Has Chnaged but Not Same as Archive Would Be" + vbCrLf + "Put the Test Flag Run Option Menu On and Try - Unload and Reload the API Form is Next Workaround OPtion", vbMsgBoxSetForeground
''
''        End If
''    End If
''
''End If
'
'
'
'
'
''End If
'




End Sub

Private Sub Timer_MOUSE_2_MINUTE_Timer()

Exit Sub

'------------------------
'2 MIN TIMER AFTER ACTIVE END
Timer_MOUSE_2_MINUTE.Enabled = False

'------------------------
'THE INTERVAL 1 MIN TIMER
Timer_MOUSE_3_MINUTE.Enabled = False
Timer_MOUSE_3_MINUTE.Interval = 59990
Timer_MOUSE_3_MINUTE.Enabled = True

'------------------------
MNU_IDLE_ACTIVE.Caption = "IDLE STATE = IDLE 1 MIN 1ST" 'TICK OVER FIRST BEGIN"

'If Mnu_Run_Status.Caption = "1st Run" Then
'    Mnu_Run_Status.Visible = False
'    Mnu_Clip_Status.Visible = False
'End If



'------------------------------
'HERE ANOTHER IF YOU WANT SOURCE FIND
'Call Mnu_Reset_MMControl_Click
'------------------------------
Call RESET_SETUP_SOUND_FILE("NOTSOUND")

Call Mnu_API_Unload_Click
Call Mnu_API_Reload_Click

End Sub

Private Sub Timer_OutClipChunck_Timer()




End Sub

Private Sub Timer_MOUSE_3_MINUTE_Timer()

    Exit Sub

    'ONCE A MINUTE WHEN IDLE
    
    MNU_IDLE_ACTIVE.Caption = "IDLE STATE = IDLE 1 MIN REPEATER"
'        If InStr(MNU_IDLE_ACTIVE.Caption, "1 MIN") Then GO3 = 1

    'Call Mnu_Reset_MMControl_Click
    'HERE OR DIY
    Call RESET_SETUP_SOUND_FILE("NOTSOUND")
    Call Mnu_API_Unload_Click
    Call Mnu_API_Reload_Click

End Sub

Private Sub Timer_MOUSE_5_MINUTE_Timer()

'1 Second -- ACTIVE ONCE

'WORK AROUND THIS ROUTINE GETTING HITT TO QUICK AT KEY PRESS
Time_Selector = DateDiff("s", Mouse_Keyboard_Active_Time, Now)
If Time_Selector < 8 Then
    Exit Sub
End If

'TEST SEE IF -- AD$ - VAR -- HAS BEEN USED BY THE SUB CALL OF API CLIPPER
Timer_MOUSE_5_MINUTE.Enabled = False

Dim XYZ2

'PROBLEM CLIPBOARD NOT GIVE ACCESS WHILE BUSY
On Error Resume Next
    XYZ2 = Clipboard.GetFormat(vbCFText) = True
    If Err.Number > 0 Then
        Timer_MOUSE_5_MINUTE.Enabled = True
        Exit Sub
    End If
On Error GoTo 0

If XYZ2 Then
    
    On Error Resume Next
        ADTEST$ = Clipboard.GetText
        If Err.Number > 0 Then
            Timer_MOUSE_5_MINUTE.Enabled = True
            Exit Sub
        End If
    On Error GoTo 0
    
    If Len(ADTEST$) <= LimitClipSize Then
        
        If ADTEST$ <> AD$ And ADTEST$ <> "" And ADTEST_BEFORE$ <> ADTEST$ Then
            
            ADTEST_BEFORE$ = ADTEST$
            
            Me.WindowState = vbNormal
            Do
                DoEvents
                Me.Refresh
                DoEvents
                If Me.WindowState = vbMinimized Then Sleep 100
            Loop Until Me.WindowState <> vbMinimized
            
            'FORM_STAY_UP_WITH_MSGBOX = True
            
            
            'Mouse_Keyboard_Active_Time
            'IS THE TIME OS THE MOMENT A KEY OR MOUSE PRESS
            'SO IF 0 THEN THIS GETTING HIT SOON AS
            'SHOURLD NOT BE 0 BUT A MOMENT AFTER
            
            Time_Selector = DateDiff("s", Mouse_Keyboard_Active_Time, Now)
            If Len(ADTEST$) > 400 Then
                AWOL_TEXT_VAR = "<--TEXT-->" + vbCrLf + Mid(ADTEST$, 1, 400) + vbcflf
                AWOL_TEXT_VAR = AWOL_TEXT_VAR + "</--TEXT-->" + vbCrLf
                AWOL_TEXT = "This is First 400 Byte of" + Str(Len(ADTEST$)) + " Text Showing"
            Else
                AWOL_TEXT_VAR = "<--TEXT-->" + vbCrLf + ADTEST$ + vbcflf
                AWOL_TEXT_VAR = AWOL_TEXT_VAR + "</--TEXT-->" + vbCrLf
                
                AWOL_TEXT = Trim(Str(Len(AWOL_TEXT_VAR))) + " Byte of Text Showing"
            End If
            
            MSGBOX2 = Format(Now, "DDD DD-MM-YYYY__HH:MM:SS ") + vbCrLf + Trim(Str(Time_Selector)) + " -- Second -- Resume From IDLE CHECK" + vbCrLf + "Problem Clipboard Not Logging TEXT BY COMPARE VARIABLE" + vbCrLf + "Put the Test Flag *ON* -- TEST by Get Yourself a COPY PASTE" + vbCrLf + "And THEN Test with -- Options Menu -- Reset and Reload the API Form" + vbCrLf + "is THE Next Workaround OPtion" + vbCrLf + "Or Better Solution Use Enter Large Text into Logger -- is --" + vbCrLf + "Format Text -- Menu --" + vbCrLf + AWOL_TEXT + vbCrLf + vbCrLf + AWOL_TEXT_VAR + vbCrLf + "ClipBoard Logger -- Resume From IDLE CHECK -- MOUSE_5_MINUTE_Timer"
            FRM_MSGBOX.Timer1.Enabled = True
            
            'FORM_STAY_UP_WITH_MSGBOX = False
                
        End If
    End If
End If

End Sub

Private Sub Timer_MOUSE_4_MINUTE_Timer()

'THIS ACTIVATE ON IDLE AFTER FEW SEC -- IS INDICATE IN MENU BAR
'100 MS ACTIVE ONCE
'TEST AD$ - VAR HAS BEEN USED BY THE SUB CALL OF API CLIPPER
'-----------------------------------
'---------------------------------------------------------
'FALSE AND THEN ENABLE WITH INTERVAL WORK AS A RESET TIMER
'---------------------------------------------------------
If Not (Timer_MOUSE_4_MINUTE.Interval = 59999 And Timer_MOUSE_4_MINUTE.Enabled = True) Then
    Timer_MOUSE_4_MINUTE.Enabled = False
        Timer_MOUSE_4_MINUTE.Interval = 59999 ' 1 MIN IN MICRO SECOND
    Timer_MOUSE_4_MINUTE.Enabled = True
End If

Dim XYZ2

'PROBLEM CLIPBOARD NOT GIVE ACCESS WHILE BUSY
On Error Resume Next
    XYZ2 = Clipboard.GetFormat(vbCFText) = True
    If Err.Number > 0 Then
        
        '---------------------------------------------------------
        'FALSE AND THEN ENABLE WITH INTERVAL WORK AS A RESET TIMER
        '---------------------------------------------------------
        Timer_MOUSE_4_MINUTE.Enabled = False
            Timer_MOUSE_4_MINUTE.Interval = 100 'QUICK
        Timer_MOUSE_4_MINUTE.Enabled = True
        Exit Sub
    End If
On Error GoTo 0


'THIS IS ERROR AT GET CLIPBOARD ON TIMER
'GETFORMAT MAYBE BETTER
'SO WE WAIT 5 SEC AND NOT RUN AGAIN UNTIL KEY OR MOUSE
If KEYBOARD_OR_MOUSE_ACTIVE > Now Then
    KEYBOARD_OR_MOUSE_ACTIVE_LATCH = False
    Exit Sub
End If
If KEYBOARD_OR_MOUSE_ACTIVE_LATCH = True Then Exit Sub
KEYBOARD_OR_MOUSE_ACTIVE_LATCH = True


If XYZ2 = True Then
            
    On Error Resume Next
        ADTEST$ = Clipboard.GetText
        If Err.Number > 0 Then
            '---------------------------------------------------------
            'FALSE AND THEN ENABLE WITH INTERVAL WORK AS A RESET TIMER
            '---------------------------------------------------------
            Timer_MOUSE_4_MINUTE.Enabled = False
                Timer_MOUSE_4_MINUTE.Interval = 100 'QUICK
            Timer_MOUSE_4_MINUTE.Enabled = True
            Exit Sub
        End If
    On Error GoTo 0
    
    If Len(ADTEST$) <= LimitClipSize Then
        
        If ADTEST$ <> AD$ And ADTEST$ <> "" And ADTEST_BEFORE$ <> ADTEST$ Then
            
            ADTEST_BEFORE$ = ADTEST$
            
            Me.WindowState = vbNormal
            Do
                DoEvents
                Me.Refresh
                DoEvents
                If Me.WindowState = vbMinimized Then Sleep 100
            Loop Until Me.WindowState <> vbMinimized
            
            'FORM_STAY_UP_WITH_MSGBOX = True
            Time_Selector = DateDiff("s", Mouse_Keyboard_Active_Time, Now)
            
            If Len(ADTEST$) > 400 Then
                AWOL_TEXT_VAR = "<--TEXT-->" + vbCrLf + Mid(ADTEST$, 1, 400) + vbcflf
                AWOL_TEXT_VAR = AWOL_TEXT_VAR + "</--TEXT-->" + vbCrLf
                AWOL_TEXT = "This is First 400 Byte of" + Str(Len(ADTEST$)) + " Text Showing"
            Else
                AWOL_TEXT_VAR = "<--TEXT-->" + vbCrLf + ADTEST$ + vbcflf
                AWOL_TEXT_VAR = AWOL_TEXT_VAR + "</--TEXT-->" + vbCrLf
                
                AWOL_TEXT = Trim(Str(Len(AWOL_TEXT_VAR))) + " Byte of Text Showing"
            End If
            
            MSGBOX2 = Format(Now, "DDD DD-MM-YYYY__HH:MM:SS ") + vbCrLf + Trim(Str(Time_Selector)) + " -- Second -- IDLE CHECK" + vbCrLf + "Problem Clipboard Not Logging TEXT BY COMPARE VARIABLE" + vbCrLf + "Put the Test Flag *ON* -- TEST by Get Yourself a COPY PASTE" + vbCrLf + "And THEN Test with -- Options Menu -- Reset and Reload the API Form" + vbCrLf + "is THE Next Workaround Option" + vbCrLf + "Or Better Solution Use Enter Large Text into Logger -- is --" + vbCrLf + "Format Text -- Menu --" + vbCrLf + AWOL_TEXT + vbCrLf + AWOL_TEXT_VAR + vbCrLf + "ClipBoard Logger -- Resume From IDLE CHECK -- MOUSE_4_MINUTE_Timer"
            FRM_MSGBOX.Timer1.Enabled = True

            'FORM_STAY_UP_WITH_MSGBOX = False
                
        End If
    End If
End If

End Sub

Private Sub TIMER_OutClipChunck_Txt_Timer()

    '10MS TIMER

    FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#OutClipChunck.Txt"
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FR1 = FreeFile
    On Error Resume Next
    Open FILENAME_IN_USE_CHECK For Output As #FR1
        
        If Err.Number > 0 Then
'            MESSENGER = FILENAME_IN_USE_CHECK + vbCrLf + "WAS NOT SAVED AT EXIT MAY BE A PROBLEM FOR RESTORE AND VIEW LAST ARCHIVE DATA INFO FROM TEXTBOX"
'            Me.WindowState = vbNormal
'            DoEvents
'            MsgBox MESSENGER, vbMsgBoxSetForeground
        Else
            Print #FR1, AD$;
            Close #FR1
            If Err.Number = 0 Then
                TIMER_OutClipChunck_Txt.Enabled = False
            End If
        End If
    On Error GoTo 0

End Sub

'Private Sub Timer_UNLOAD_FORM_Timer()
'
'Timer_UNLOAD_FORM
'MsgBox "CLIPBOARD LOGGER IS STILL LOADED NOT UNLOAD"
'
'End Sub

Private Sub Timer_WEATHER_URL_Timer()
    'ORGINAL TIMER NOT DEPENDANT ON ANY SINGLE THING
    Timer_WEATHER_URL.Enabled = False
    Exit Sub
    
    Dim FDS
    FDS = "D:\0 00 Art Loggers\URL SCREEN SHOT\BBC WEATHER\URL SCREENSHOT - BBC WEATHER.JPG"
    Call SUB_RENAME_JPG_URL(FDS)
    FDS = "D:\0 00 Art Loggers\URL SCREEN SHOT\METOFFICE 01 - Satellite & Radar\URL SCREENSHOT - METOFFICE 01 - Satellite & Radar.JPG"
    Call SUB_RENAME_JPG_URL(FDS)
    FDS = "D:\0 00 Art Loggers\URL SCREEN SHOT\METOFFICE 02 - Satellite & Visible\URL SCREENSHOT - METOFFICE 02 - Satellite & Visible.JPG"
    Call SUB_RENAME_JPG_URL(FDS)
    
    
End Sub

Sub TIMER_FORM_RESIZE_TIMER()
    'CALLED BY FORM_RESIZE

    'Exit Sub

    If MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = True Then

        Text1.Enabled = False
        Text1.SelStart = 1
        Text1.Enabled = True
        Text1.SelStart = Len(Text1)
        
        Timer_FORM_RESIZE.Enabled = False

    End If

End Sub


Private Sub Form_Resize()

'Exit Sub

If DONT_RESIZE_WHILE_SETUP = True Then Exit Sub

Timer_MENU_HEIGHT_CHANGED.Enabled = True

'MAYBE SOMETIMES WINDOW IS MIN WHILE TAKE A CHANGE
'AT CLIPBOARD
'THINK NOT
'If Me.WindowState = vbMinimized Then RRR = Now - TimeSerial(0, 0, 3)

HHH = Now + TimeSerial(0, 2, 0) ' vbMinimized WINDOW

'If RESIZE_AT_LOAD = True And Me.WindowState <> vbMinimized Then
If Me.WindowState <> vbMinimized And DONT_RESIZE_RUN_ONCE_OR_NORM = True Then
    DO_THIS_RUN_NOT_DO_TWICE = True
'    Me.WindowState = vbNormal
    DONT_RESIZE_WHILE_SETUP = True
    '-----------------------------------------
    Call SETUP_DIMENSIONS_NORM
    Call SETUP_DIMENSIONS_INNER_FORM
    '-----------------------------------------
    'RESIZE_AT_LOAD = False
    DONT_RESIZE_WHILE_SETUP = False
    DONT_RESIZE_RUN_ONCE_OR_NORM = False

End If

If Me.WindowState = vbMinimized Then
    'MNU_SCANPATH_COUNTER.Visible = False
'    Timer_API_OKAY_COLOUR.Enabled = False
Else
'    Timer_API_OKAY_COLOUR.Enabled = True
End If


If Me.WindowState <> vbMinimized Then
    I = SetForegroundWindow(Me.hwnd)
End If


'----------------------------------
'vbMinimized - VICTIMIZED
'VITMIN PILL
'VICTIMIZOR - ADVISOR
'----------------------------------
'----------------------------------
'IF FROM vbMinimized TO vbNormal
'IF FROM vbMinimized TO vbMaximized
'IF FROM vbNormal TO vbMaximized
'IF FROM vbMaximized TO vbNormal
'----------------------------------
'OUGHT TO THINK OF THIS A WHILE AGO
'----------------------------------

DO_THE_CODE = False
If Me.WindowState = vbNormal And O_RESIZE = vbMinimized Then DO_THE_CODE = True
If Me.WindowState = vbMaximized And O_RESIZE = vbMinimized Then DO_THE_CODE = True
If Me.WindowState = vbNormal And O_RESIZE = vbMaximized Then DO_THE_CODE = True
If Me.WindowState = vbMaximized And O_RESIZE = vbNormal Then DO_THE_CODE = True
If DO_THIS_RUN_NOT_DO_TWICE = True Then DO_THE_CODE = False

'ANNOYING THING ABOUT COPY PASTE PICK UP
'WANT TO USE KEYBOARD TO CURSOR OVER AND PICK UP A WORD IF CAN
'RATHER THAN MOUSE

If DO_THE_CODE = True Then
    '--------------------------------------
    'NOTHING TO LEARN ABOUT SET CENTER MODE
    'ALL DONE IN THE CENTER ROUTINE
    '--------------------------------------

'    Me.WindowState = vbNormal
    DONT_RESIZE_WHILE_SETUP = True
    '-----------------------------------------
    'Call SETUP_DIMENSIONS_NORM
    Call SETUP_DIMENSIONS_INNER_FORM
    '-----------------------------------------
    'RESIZE_AT_LOAD = False
    DONT_RESIZE_WHILE_SETUP = False
    DONT_RESIZE_RUN_ONCE_OR_NORM = False
End If

O_RESIZE = Me.WindowState



'Call SETUP_DIMENSIONS_INNER_FORM

Exit Sub


''label222.Caption = Val(Mnu_CB_Size02)
'Debug.Print "---"
'For Each Control In Controls
'
''        TNAME = CURR_OBJ
'        TNAME = Control.Name
'        If InStr(LCase(TNAME), "mnu") > 0 Then
'            TTop = Control.Top
'            TWidth = Control.Width
'            THeight = Control.Height
''            Debug.Print Str(TTop) + Str(THeight); " " + TNAME
'
'
'
'        End If
''                    Debug.Print Control.HNWD
'        TIndex = Control.TabIndex
'        TLeft = Control.Left
'
'
''        THIS GET THE TITLE BAR HEIGHT NOT THE MENU
''        Debug.Print TitleBarHeight
'
''Dim Position2 As rect
''
''GetClientRect Me.Mnu_CB_Size02.hwnd, Position2
'''Position.x = Position2.Left
''Debug.Print Position2.Top
''Debug.Print Position2.Bottom
'
'
''Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
''Private Const SM_CYCAPTION = 4
'
''Property Get TitleBarHeight() As Long
''    TitleBarHeight = GetSystemMetrics(SM_CYCAPTION)
''End Property
'
'Next Control


'If GetComputerName = "MATT-555ROIDS" Then
''    Text1.Height = 10000
''    Text1.Width = 18000
'End If
'If GetComputerName = "55-88-HAPPY" Then
''    Text1.Height = 7000
''    Text1.Width = 14000
'End If


'If Text1.Font.Size > 14 Then
'    Text1.SelStart = 0
'    Text1.SelLength = Len(Text1)
'    Text1.Font.Size = 12
'    Text1.SelStart = Len(Text1)
'End If

'Me.SetFocus

'XXMOUSEMOVE = 0
'Text1.SelStart = Len(Text1)
'


End Sub

Private Sub Mnu_Minimize_Click()

    Me.WindowState = vbMinimized
    'Call Form_Resize
End Sub

Public Sub Mnu_Max_Click()

    'GIVE ACCESS TO RUN
'    DONT_RESIZE_RUN_ONCE_OR_NORM = True
   
'    Me.WindowState = vbMaximized
    Call SETUP_DIMENSIONS_MAX
    
    'Call Form_Resize
End Sub


Private Sub MNU_Norm_Click()
   
    'GIVE ACCESS TO RUN
    'DONT_RESIZE_RUN_ONCE_OR_NORM = True
   
    Call SETUP_DIMENSIONS_NORM
   
End Sub

Private Sub Mnu_Center_Click()
    
    'If Me.WindowState = vbMaximized Then MsgBox "Not in Windowstate = vbMaximized - Try Windowstate - Normal"
    
    If Me.WindowState <> vbMinimized Then
        
        DONT_RESIZE_WHILE_SETUP = True
        Me.WindowState = vbNormal
    
    End If
    
    'Form1.Left = (Screen.Width - Me.Width) / 2
    'Form1.Top = (Screen.Height - Me.Height) / 2

    Call SETUP_DIMENSIONS_CENTER


End Sub


Sub SETUP_DIMENSIONS_INNER_FORM()

'DoEvents   ' Yield for other processing.
'Line Method Example

'Line1.BorderWidth = 40
'Border Line On the Edge

Label1.Top = 0
Label1.Width = Form1.Width - 70

Text1.Left = -8

Text1.Top = Label1.Height

'On Error Resume Next
'Mnu_Height.Caption = Menu_Height

'VER SIX IS WIN 7
'GOT THICKER BOARDERS OF FORMS WINDOW -- SCROLL BAR MISSING A BIT

'ALL WIN SETUP SEEM TO HAVE OWN SIZE

WIDTH_ADJUST = 70 ' FOR WIN XP
If GETWinNT_Ver = "WIN XP" Then WIDTH_ADJUST = 70
If GETWinNT_Ver = "WIN 7" Then WIDTH_ADJUST = 170
If GETWinNT_Ver = "WIN 10" Then WIDTH_ADJUST = 250

If GetComputerName = "1-ASUS-EEE" Then WIDTH_ADJUST = 70
If GetComputerName = "1-ASUS-X5DIJ" Then WIDTH_ADJUST = 70
If GetComputerName = "3-LINDA-PC" Then WIDTH_ADJUST = 170
If GetComputerName = "4-ASUS-GW522VW" Then WIDTH_ADJUST = 250
If GetComputerName = "5-ASUS-P2520LA" Then WIDTH_ADJUST = 220 '100

HEIGHT_ADJUST = 70 ' FOR WIN XP
If GETWinNT_Ver = "WIN XP" Then HEIGHT_ADJUST = 70
If GETWinNT_Ver = "WIN 7" Then HEIGHT_ADJUST = 70
If GETWinNT_Ver = "WIN 10" Then HEIGHT_ADJUST = 100

'Bigger adjust number smaller inner form
If GetComputerName = "1-ASUS-EEE" Then HEIGHT_ADJUST = 70
If GetComputerName = "1-ASUS-X5DIJ" Then HEIGHT_ADJUST = 70
If GetComputerName = "3-LINDA-PC" Then HEIGHT_ADJUST = 70
If GetComputerName = "4-ASUS-GW522VW" Then HEIGHT_ADJUST = 100
If GetComputerName = "5-ASUS-P2520LA" Then HEIGHT_ADJUST = 0

'HIGHER NUMBER SMALLER INNER BOX

Text1.Width = Form1.Width - WIDTH_ADJUST + 20

'MAKE FORM TALLER OR TEXT BOX
'FORM IS PRIOITY OVER TEXT BOX

XXDD = Form1.Height - (Menu_Height * Screen.TwipsPerPixelY) - 500 - Label1.Height
If XXDD > 0 Then Text1.Height = XXDD - HEIGHT_ADJUST

If Me.Top + Me.Left + Me.Width + Me.Height = OLTLWH Then Exit Sub
OLTLWH = Me.Top + Me.Left + Me.Width + Me.Height

Timer_FORM_RESIZE.Enabled = False
Timer_FORM_RESIZE.Interval = 40
Timer_FORM_RESIZE.Enabled = True



End Sub


Sub SETUP_DIMENSIONS_NORM()

   
    Dim RECTT1 As RECT
    Dim RECTT2 As RECT
    Dim RECTT4 As RECT
    
    On Error Resume Next
    
    If Me.WindowState <> vbMinimized Then
        
        DONT_RESIZE_WHILE_SETUP = True
        Me.WindowState = vbNormal
                
        Test1 = FindWindow("Shell_TrayWnd", vbNullString)
        T1 = GetWindowRect(Test1, RECTT1) ' BOTTOM BAR
        Test2 = FindWindow("MOM Class", vbNullString)
        T1 = GetWindowRect(Test2, RECTT2) ' TOP BAR
        'IRON BAR - DUMB BELLS
        TEST4 = FindWindowPart_BASEBAR("BaseBar")
        T1 = GetWindowRect(TEST4, RECTT4) ' LEFT BAR
        A = RECTT4.Top
        y1 = (RECTT1.Bottom - RECTT1.Top) * Screen.TwipsPerPixelY
        
        If RECTT2.Bottom - RECTT2.Top < 33 Then
        
            Y2 = (RECTT2.Bottom - RECTT2.Top) * Screen.TwipsPerPixelY
        
        Else
        
            Y2 = 0
        
        End If
        
        If TEST4 > 0 Then
            If RECTT4.Right = RECTT4.Left Then
                X4 = (RECTT4.Right) * Screen.TwipsPerPixelX
            Else
                X4 = (RECTT4.Right - RECTT4.Left) * Screen.TwipsPerPixelX
            End If
        End If
        
        SCREEN_WIDTH_SPACE = Screen.Width - X4
        SCREEN_HEIGHT_SPACE = Screen.Height - y1 - Y2
        
        Form1.Width = SCREEN_WIDTH_SPACE - 1200
        Form1.Height = SCREEN_HEIGHT_SPACE - 900
        'THIS THE FORM HEIGHT
        'THE BOX INSIDE IS ADJUST AFTER
        
        
        
        
        'Form1.Left = (Screen.Width - Me.Width) / 2
        Form1.Left = X4 + (SCREEN_WIDTH_SPACE - Me.Width) / 2
        Form1.Top = Y2 + (SCREEN_HEIGHT_SPACE - Me.Height) / 2
    
'        DoEvents
        
        Call SETUP_DIMENSIONS_INNER_FORM
    
    End If

End Sub


Sub SETUP_DIMENSIONS_CENTER()

   
    Dim RECTT1 As RECT
    Dim RECTT2 As RECT
    Dim RECTT4 As RECT
    
    On Error Resume Next
    
    If Me.WindowState <> vbMinimized Then
        
        DONT_RESIZE_WHILE_SETUP = True
        
        Test1 = FindWindow("Shell_TrayWnd", vbNullString)
        T1 = GetWindowRect(Test1, RECTT1) ' BOTTOM BAR
        Test2 = FindWindow("MOM Class", vbNullString)
        T1 = GetWindowRect(Test2, RECTT2) ' TOP BAR
        'IRON BAR - DUMB BELLS
        TEST4 = FindWindowPart_BASEBAR("BaseBar")
        T1 = GetWindowRect(TEST4, RECTT4) ' LEFT BAR
        A = RECTT4.Top
        y1 = (RECTT1.Bottom - RECTT1.Top) * Screen.TwipsPerPixelY
        Y2 = (RECTT2.Bottom - RECTT2.Top) * Screen.TwipsPerPixelY
        
        If TEST4 > 0 Then
            If RECTT4.Right = RECTT4.Left Then
                X4 = (RECTT4.Right) * Screen.TwipsPerPixelX
            Else
                X4 = (RECTT4.Right - RECTT4.Left) * Screen.TwipsPerPixelX
            End If
        End If
        
        SCREEN_WIDTH_SPACE = Screen.Width - X4
        SCREEN_HEIGHT_SPACE = Screen.Height - y1 - Y2
        
        'Form1.Left = (Screen.Width - Me.Width) / 2
        Form1.Left = X4 + (SCREEN_WIDTH_SPACE - Me.Width) / 2
        Form1.Top = Y2 + (SCREEN_HEIGHT_SPACE - Me.Height) / 2
        
        DoEvents
        
        Call SETUP_DIMENSIONS_INNER_FORM
    
    End If



End Sub

Sub SETUP_DIMENSIONS_MAX()

   
    Dim RECTT1 As RECT
    Dim RECTT2 As RECT
    Dim RECTT4 As RECT
    
    On Error Resume Next
    
    If Me.WindowState <> vbMinimized Then
        
        DONT_RESIZE_WHILE_SETUP = True
        Me.WindowState = vbMaximized
                
        Test1 = FindWindow("Shell_TrayWnd", vbNullString)
        T1 = GetWindowRect(Test1, RECTT1) ' BOTTOM BAR
        Test2 = FindWindow("MOM Class", vbNullString)
        T1 = GetWindowRect(Test2, RECTT2) ' TOP BAR
        'IRON BAR - DUMB BELLS
        TEST4 = FindWindowPart_BASEBAR("BaseBar")
        T1 = GetWindowRect(TEST4, RECTT4) ' LEFT BAR
        A = RECTT4.Top
        y1 = (RECTT1.Bottom - RECTT1.Top) * Screen.TwipsPerPixelY
        
        If RECTT2.Bottom - RECTT2.Top < 33 Then
        
            Y2 = (RECTT2.Bottom - RECTT2.Top) * Screen.TwipsPerPixelY
        
        Else
        
            Y2 = 0
        
        End If
        
        If TEST4 > 0 Then
            If RECTT4.Right = RECTT4.Left Then
                X4 = (RECTT4.Right) * Screen.TwipsPerPixelX
            Else
                X4 = (RECTT4.Right - RECTT4.Left) * Screen.TwipsPerPixelX
            End If
        End If
        
        SCREEN_WIDTH_SPACE = Screen.Width - X4
        SCREEN_HEIGHT_SPACE = Screen.Height - y1 - Y2
        
        Form1.Width = SCREEN_WIDTH_SPACE - 1200
        Form1.Height = SCREEN_HEIGHT_SPACE - 900
        'THIS THE FORM HEIGHT
        'THE BOX INSIDE IS ADJUST AFTER
        
        
        
        
        'Form1.Left = (Screen.Width - Me.Width) / 2
        Form1.Left = X4 + (SCREEN_WIDTH_SPACE - Me.Width) / 2
        Form1.Top = Y2 + (SCREEN_HEIGHT_SPACE - Me.Height) / 2
    
'        DoEvents
        
        Call SETUP_DIMENSIONS_INNER_FORM
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Unload FrmJoy
'Unload frmSender

'----------------------------
MMControl2.Command = "close"
MMControl1.Command = "close"
'----------------------------

If TIMER_OutClipChunck_Txt.Enabled = True Then
    FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#OutClipChunck.Txt"
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FR1 = FreeFile
    On Error Resume Next
    Open FILENAME_IN_USE_CHECK For Output As #FR1
        If Err.Number > 0 Then
            MESSENGER = FILENAME_IN_USE_CHECK + vbCrLf + "WAS NOT SAVED AT EXIT OR AT ITS 10MS INTERVAL WHEN CHANGE" + vbCrLf + "MAY BE A PROBLEM FOR RESTORE AND VIEW LAST ARCHIVE DATA INFO FROM TEXTBOX -- LOOK FOR A DOUBLE COPY BACKUP SOLUTION LATER" + vbCrLf + "YOUR DRIVE MUST BE LOCKED OR SOMETHING"
            Me.WindowState = vbNormal
            DoEvents
            MsgBox MESSENGER, vbMsgBoxSetForeground
        Else
            Print #FR1, AD$;
            Close #FR1
        End If
    On Error GoTo 0
End If


'----------------------
' ABORT USING THIS
' Call zzSave_Checks
'----------------------
' ONLY SAVE IF CHANGES
Call zzCheckTimer_Timer
'----------------------

If Me.WindowState <> 1 And EXIT_TRUE = False Then
    Me.WindowState = 1
    'test may have to put back form need reseting
    'Unload FrmJoy
    Cancel = True
    Exit Sub
End If

EXIT_TRUE = True
Dim Control
'SET ALL TIMERS IN ALL FORMS ENABLED=FALSE
On Error Resume Next
    For I = 0 To Forms.Count - 1
        For Each Control In Forms(I).Controls
            If InStr(UCase(Control.Name), "TIMER") > 0 Then
                'Debug.Print Control.Name
                Control.Enabled = False
            End If
        Next
    Next I
On Error GoTo 0
   
Cancel = False
EXIT_TRUE = True
Dim Form As Form
For Each Form In Forms
    EXIT_TRUE = True
    Unload Form
    Set Form = Nothing
Next Form

Cancel = False

'---------------------------------------------------------
On Error Resume Next
'ERROR BECAUSE -- NOT CONTROLED HERE IF DRIVE DON'T EXIST
'---------------------------------------------------------
EXIT_TRUE = True
UNLOAD_FORM_SAFE.Timer1.Enabled = True

'Cancel = True
End Sub

Sub SETUP_SOUND_FILE(VARSOUND)

'DO THIS LATER

'------------------------------------------------------
'SUBPATH IN USE THEN DO LATER
'------------------------------------------------------

Do
    If ScanPath.ListView1.ListItems.Count > 0 Then
        Sleep 1
    End If
    If EXIT_TRUE = True Then Exit Sub
Loop Until ScanPath.ListView1.ListItems.Count = 0

'------------------------------------------------------


'Call SETUP_SOUND_FILE("NOTSOUND")
'VARSOUND = "NOTSOUND"
Dim XCOUNT1, XCOUNT2, XCHECKED

For Each Control In MNU_SOUND_OPTION
    If Control.Caption <> "MNU_SOUND_OPTION" Then
        XCOUNT1 = XCOUNT1 + 1
    End If
Next

'FIND THE CHECKED ONE AND STORE
For Each Control In MNU_SOUND_OPTION
    If Control.Checked = True Then XCHECKED = Mid(Control.Caption, 25): Exit For
Next

For Each Control In MNU_SOUND_OPTION
    Control.Caption = "MNU_SOUND_OPTION"
    Control.Enabled = False
    Control.Checked = False
Next

'-----------------------------------
'ONLY USE IF WANT EXTRA DATE INFO ADDED
'-----------------------------------
ScanPath.chkCopyMemory.Value = vbUnchecked
'-----------------------------------
ScanPath.TxtPath = App.Path + "\Sound_Wav's\"
ScanPath.cboMask = "*.WAV"
ScanPath.chkSubFolders = vbUnchecked

ARCHIVE_Path_Of_Sound_File = Path_Of_Sound_File
Path_Of_Sound_File = ""

'WE WANT THEM SORTED ALPHABETICALY
Call ScanPath.cmdScan_Click
'LAST BREAK POINT WHILE WORK WAS SET HERE

ReDim Path_Of_Sound(ScanPath.ListView1.ListItems.Count)
XCOUNT2 = ScanPath.ListView1.ListItems.Count
For WE = 1 To ScanPath.ListView1.ListItems.Count
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    
    If WE > 9 Then MsgBox "Maximum of 9 Sounds in Menu - Delete Some": Exit For
    
    Path_Of_Sound(WE) = A1 + B1
    MNU_SOUND_OPTION(WE).Caption = "SOUND OPTION - 1 -" + Str(WE) + " OF" + Str(XCOUNT2) + " - \" + B1
    MNU_SOUND_OPTION(WE).Visible = True
    'MNU_SOUND_OPTION(i).Enabled = True
    
Next

ScanPath.ListView1.ListItems.Clear
ScanPath.Hide

'For i = 1 To UBound(Path_Of_Sound)
'    'MNU_SOUND_OPTION(i).Caption = "SOUND OPTION -" + Str(i) + " - " + Path_Of_Sound(i)
'    MNU_SOUND_OPTION(i).Visible = True
'    MNU_SOUND_OPTION(i).Enabled = True
'Next

'DISABLE WHAT WAS NOT USED BUT WAIT LATER FOR ENABLE
'DISABLE WHAT WAS NOT USED -------------------------
For Each Control In MNU_SOUND_OPTION
    If Control.Caption = "MNU_SOUND_OPTION" Then
        Control.Visible = False
        Control.Enabled = False
    End If
Next

'GOT TO RUN TWICE - THINK
'SEARCHES FOR THE TEXT CAPTIONS BEEN PUT IN
'--------
'DONT RUN AGAIN WHEN PROGRAM IS GOING SETTINGS ONLY SAVED ON EXIT
'AND YES DO RUN TWICE ON FIRST LOAD
If LATCH_RUN_ONCE = False Then Call zzLoad_Checks
'--------

'MIGHT BE FROM Call zzLoad_Checks
Dim Mnu_Sound_Is_There_One_Selected

For Each Control In MNU_SOUND_OPTION
    If Control.Checked = True Then Mnu_Sound_Is_There_One_Selected = True: Exit For
    
'    If Control.Caption = "MNU_SOUND_OPTION" Then
'        Control.Visible = False
'        Control.Enabled = False
'        Else
'        'Control.Visible = True
'        'Control.Enabled = True
'
'
'    End If

Next


Dim XXVAR
'MAKE SURE ONLT ONE IS SELECTED
'XXVAR = False
For Each Control In MNU_SOUND_OPTION
    
    'TOPSY TURVY CODE 2 NEST
    If Control.Checked = True And XXVAR = True Then
        Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
        Control.Checked = False
        Control.Enabled = True
    End If
    
    If Control.Checked = True Then
        Path_Of_Sound_File = App.Path + "\Sound_Wav's" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
        XXVAR = True
        
    End If

Next

'THIS NEVER HAPPENS
If XXVAR = True Then Mnu_Sound_Is_One_There_Selected = True


If Mnu_Sound_Is_One_There_Selected = True Then
    'TAKE THE CHECKED SELECTED ONE -- DUPLICATE ROUTINE OR
    For Each Control In MNU_SOUND_OPTION
        If Control.Checked = True Then
            Path_Of_Sound_File = App.Path + "\Sound_Wav's" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
            'PROBLEM
            Exit For
        End If
    Next
Else

'OPTION 1 OF IF
'--------------
'HAS ONE BEEN SELECTED BEFORE WHILE PROGRAM RUNNING - IF SO SEARCH IT OUT
    If XCHECKED <> "" Then
        For Each Control In MNU_SOUND_OPTION
            If InStr(Control.Caption, XCHECKED) > 0 Then
                Path_Of_Sound_File = App.Path + "\Sound_Wav's" + Mid(XCHECKED, 5)
                Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
                Control.Checked = True
                Control.Enabled = True
                Exit For
            End If
        Next
    End If

'OPTION 2 OF IF
'--------------
'PICK THE FIRST ONE - PICK ME UP - GIVE US A LIFT - DON'T MIND IF I DO - HELP MYSELF - INDULGE MYSELF
'TAKE THE FIRST ONE
'AND THEN MAKE SURE IT'S SELECTED
    If Path_Of_Sound_File = "" Then
        For Each Control In MNU_SOUND_OPTION
    '        Path_Of_Sound_File = App.Path + "\Sound_Wav's" + Mid(MNU_SOUND_OPTION(1).Caption, 25)
            If InStrRev(Control.Caption, "\") > 0 Then
                Path_Of_Sound_File = App.Path + "\Sound_Wav's" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
        
                MNU_SOUND_OPTION(1).Enabled = False ' DON'T DO EXTRA CLICK'S TO THE ROUTINE
                MNU_SOUND_OPTION(1).Checked = True
                MNU_SOUND_OPTION(1).Visible = True
                MNU_SOUND_OPTION(1).Enabled = True ' SAFE
                Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
                Control.Checked = True
                Control.Enabled = True
                
                Exit For
            End If
        Next
    End If
End If


If MNU_MESSAGE_BOXES.Checked = False Then
    If Path_Of_Sound_File = "" And ARCHIVE_Path_Of_Sound_File = "" Then
          Me.WindowState = vbNormal
        DoEvents

        MsgBox "YOU HAVEN'T ANY SOUND FILE'S - SELECT OPTION TO OPEN FOLDER AND PUT SOME WAV'S IN", vbMsgBoxSetForeground
    Else
        If Path_Of_Sound_File = "" Then
            Me.WindowState = vbNormal
            DoEvents
                
            MsgBox "ALL THE SOUND FILES HAVE DISAPPEARED SINCE LAST EXAMINE - YOU HAVEN'T ANY SOUND FILE'S - SELECT OPTION TO OPEN FOLDER AND PUT SOME WAV'S IN", vbMsgBoxSetForeground
        End If
    End If
End If


'LASTLY -- ENABLE THE VAILD
For Each Control In MNU_SOUND_OPTION
    If Control.Caption <> "MNU_SOUND_OPTION" Then
        Control.Enabled = True
        Control.Visible = True 'AND SAFE MESSURE
    End If
Next



'EXTRA NEW - REALLY - OR
If LATCH_RUN_ONCE = True Then
    If XCOUNT2 > XCOUNT1 Then
        Me.WindowState = vbNormal
        DoEvents
        MsgBox "YOU HAVE GOT EXTRA NEW SOUND FILES SINCE LAST EXAMINE - CHECK THE SOUND FOLDER MENU OPTION IF YOU LIKE."
    Else
        If XCOUNT2 < XCOUNT1 Then
          Me.WindowState = vbNormal
        DoEvents
   
           MsgBox "YOU HAVE LESS SOUND FILES SINCE LAST INSPECTION - CHECK THE SOUND FOLDER MENU OPTION IF YOU LIKE."
        End If
    End If
End If

'Debug.Print Path_Of_Sound_File

If Path_Of_Sound_File <> "" Then
    
    MMControl1.Command = "Stop"
    MMControl1.Command = "Close"
'    DoEvents
    MMControl1.Notify = True
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.fileName = Path_Of_Sound_File
    'Debug.Print Path_Of_Sound_File
    
    
    'Debug.Print Path_Of_Sound_File
    
    
    
    'Path_Of_Sound(1) = App.Path + "\Camera1a_2_8kHz.wav"
    'MMControl1.fileName = App.Path + "\Camera1a_2_8kHz.wav"
    'MMControl1.fileName = App.Path + "\Shot-12 Mix to Shot-18.wav"
    'MMControl1.fileName = App.Path + "\01 Woarble Tone 5 Mins.wav"
    
    MMControl1.Command = "Open"

'    DoEvents

End If


vPathSOUND2 = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\"
If Dir(vPathSOUND2, vbDirectory) = "" Then
    I = MkDirNested(vPathSOUND2)
    If I = False Then
        MsgBox "UABLE TO MKDIR NESTED" + vbCrLf + vPathSOUND2, vbMsgBoxSetForeground
    End If
End If


vFileSOUND2 = Dir(vPathSOUND2 + "*.WAV")
If vFileSOUND2 <> "" Then
    MNU_SOUND_2.Caption = "SOUND OPTION - 2 ------ \" + vFileSOUND2
    
    MMControl2.Command = "Stop"
    MMControl2.Command = "Close"
    '---- Set properties needed by MCI to open.
    MMControl2.Notify = True
    MMControl2.Wait = True
    MMControl2.Shareable = False
    MMControl2.DeviceType = "WaveAudio"
    MMControl2.fileName = vPathSOUND2 + vFileSOUND2
    ' Open the MCI WaveAudio device.
    MMControl2.Command = "Open"
       
    'MMControl2.Command = "prev"
    'MMControl2.Command = "Play"
Else
    MNU_SOUND_2.Caption = "SOUND OPTION - 2 - WAV File Not Found - 1st Found Here Used " + vPathSOUND2

End If

'    MNU_SOUND_2.Caption = "SOUND OPTION - 2 ------ WAV File Not Found - 1st Found Here Used -- " + vPathSOUND2


'Path_Of_Sound_File
    
    'D:\Wave's\Camera1a_2_8kHz.wav
    
     
    'If DUPE_CLIPPER_AT_LOAD_FORM = False Then
        
    If Mnu_SoundOn.Checked = True And VARSOUND <> "NOTSOUND" Then
'        If Clipboard.GetFormat(vbCFText) = False And Clipboard.GetFormat(vbCFBitmap) = False Then
            
            
            If DUPE_CLIPPER_AT_LOAD_FORM = False Then
            
            
                If Path_Of_Sound_File <> "" Then
                    If MNU_PLAY_SOUND_ON_LOAD.Checked = True Or MODIFY_SOUND_SELECTION = True Then
                        MMControl1.Command = "prev"
                        MMControl1.Command = "Play"
                    End If
                End If
            Else
            
                If MNU_PLAY_SOUND_ON_LOAD.Checked = True Or MODIFY_SOUND_SELECTION = True Then
                    
                    If vFileSOUND2 <> "" Then
                        MMControl2.Command = "prev"
                        MMControl2.Command = "Play"
                
                    End If
                End If
    
            End If
        
'        End If
    End If




ScanPath.ListView1.ListItems.Clear

'WORK TO DO HERE
LATCH_RUN_ONCE = True
MODIFY_SOUND_SELECTION = False
DUPE_CLIPPER_AT_LOAD_FORM = False

End Sub




Sub RESET_SETUP_SOUND_FILE(VARSOUND)
'MMControl2.fileName = vPathSOUND2 + vFileSOUND2
'Debug.Print Path_Of_Sound_File

'If VARSOUND <> "" Then Stop

On Error GoTo EXIT_END


'compare sort
'If Dir(Path_Of_Sound_File) = "" Then
'    Me.WindowState = vbNormal
'    MsgBox "SOUND FILE 01# IS MISSING", vbMsgBoxSetForeground
'End If

'If Dir(vPathSOUND2 + vFileSOUND2) = "" Then
'    Me.WindowState = vbNormal
'    MsgBox "SOUND FILE 02# IS MISSING", vbMsgBoxSetForeground
'End If










'Path_Of_Sound_File = App.Path + "\Sound_Wav's\"
vPathSOUND1 = App.Path + "\Sound_Wav's\"
vFileSOUND1 = Dir(vPathSOUND1 + "*.WAV")
Path_Of_Sound_File = vPathSOUND1 + vFileSOUND1

If Path_Of_Sound_File <> "" And Dir(Path_Of_Sound_File) <> "" Then
    
    MMControl1.Command = "Stop"
    MMControl1.Command = "Close"
    MMControl1.Notify = True
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.fileName = Path_Of_Sound_File
'    Debug.Print Path_Of_Sound_File
    
    'Path_Of_Sound(1) = App.Path + "\Camera1a_2_8kHz.wav"
    'MMControl1.fileName = App.Path + "\Camera1a_2_8kHz.wav"
    'MMControl1.fileName = App.Path + "\Shot-12 Mix to Shot-18.wav"
    'MMControl1.fileName = App.Path + "\01 Woarble Tone 5 Mins.wav"
    MMControl1.Command = "Open"
    MNU_AUDIO_01_MISSING.Visible = False

End If


vPathSOUND2 = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\"
vFileSOUND2 = Dir(vPathSOUND2 + "*.WAV")

If vFileSOUND2 <> "" And Dir(vPathSOUND2 + vFileSOUND2) <> "" Then
    MNU_SOUND_2.Caption = "SOUND OPTION - 2 ------ \" + vFileSOUND2
    
    MMControl2.Command = "Stop"
    MMControl2.Command = "Close"
    '---- Set properties needed by MCI to open.
    MMControl2.Notify = True
    MMControl2.Wait = True
    MMControl2.Shareable = False
    MMControl2.DeviceType = "WaveAudio"
    MMControl2.fileName = vPathSOUND2 + vFileSOUND2
    ' Open the MCI WaveAudio device.
    MMControl2.Command = "Open"
       
    'MMControl2.Command = "prev"
    'MMControl2.Command = "Play"
    
    MNU_AUDIO_01_MISSING.Visible = False
End If
     
'If DUPE_CLIPPER_AT_LOAD_FORM = False Then

'If Mnu_SoundOn.Checked = True And VARSOUND <> "NOTSOUND" Then

 If VARSOUND <> "NOTSOUND" Then
            
    If Path_Of_Sound_File <> "" Then
        'MMControl1.Command = "prev"
        MMControl1.Command = "Play"
    End If
    If vFileSOUND2 <> "" Then
        'MMControl2.Command = "prev"
        MMControl2.Command = "Play"
    End If

End If


If Dir(Path_Of_Sound_File) = "" Then
    Me.WindowState = vbNormal
    'DEBUG_HERE
    'MsgBox "SOUND FILE 01# IS MISSING", vbMsgBoxSetForeground
    MNU_AUDIO_01_MISSING.Visible = True
Else
    MNU_AUDIO_01_MISSING.Visible = False

End If

If Dir(vPathSOUND2 + vFileSOUND2) = "" Then
    Me.WindowState = vbNormal
    'DEBUG_HERE
    'MsgBox "SOUND FILE 02# IS MISSING", vbMsgBoxSetForeground
    MNU_AUDIO_02_MISSING.Visible = True
Else
    MNU_AUDIO_02_MISSING.Visible = False

End If



'LATCH_RUN_ONCE = True
'MODIFY_SOUND_SELECTION = False
'DUPE_CLIPPER_AT_LOAD_FORM = False

EXIT_END:

End Sub




Private Sub MNU_LAST_GRAB_ALL_CAPS_Click()
    'AD$
    
    If MNU_LAST_GRAB_ALL_CAPS.Checked = False Then
        MNU_LAST_GRAB_ALL_CAPS.Checked = True
        GoTo EXIT1
    End If


     If MNU_LAST_GRAB_ALL_CAPS.Checked = True Then
        MNU_LAST_GRAB_ALL_CAPS.Checked = False
    End If
    
    
EXIT1:

AD$ = UCase(AD$)

'Clipboard.Clear
'Clipboard.SetText AD$

    EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        Clipboard.SetText AD$
    EXECUTE_TIMER_ENABLED = True
AD_DATE = 0

Me.WindowState = 1

'solvability
          
End Sub



Private Sub LAST_GRAB_ALL_CAPS_002_Click()
    
    If MNU_LAST_GRAB_ALL_CAPS.Checked = False Then Exit Sub
    
    If Clipboard.GetFormat(vbCFText) = False Then Exit Sub
    
    AD$ = UCase(AD$)
    EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        Clipboard.SetText AD$
    EXECUTE_TIMER_ENABLED = True
    
    Me.WindowState = 1
    AD_DATE = 0
    
'solvability
    


End Sub

Private Sub MNU_ALL_CAPS_Click()

End Sub

Private Sub MNU_CB_SIZE_BYTE_Click()

RESET_RRR = True
Call Mnu_CB_Size02(TT, TT1)
    EXECUTE_TIMER_ENABLED = False
Clipboard.Clear
    EXECUTE_TIMER_ENABLED = True
RESET_RRR = True
Clipboard.SetText TT1

End Sub



Private Sub MNU_CB_SIZE_MEG_Click()

RESET_RRR = True
Call Mnu_CB_Size02(TT, TT1)

    EXECUTE_TIMER_ENABLED = False
Clipboard.Clear
    EXECUTE_TIMER_ENABLED = True

RESET_RRR = True
Clipboard.SetText TT

End Sub

Sub Mnu_CB_Size02(TT, TT1)
    
    'RELATED
    'Menu_clipboard_size
    
    RESET_RRR = True
    
    Dim TTX
    Clipboard.Clear
    TT = Replace(MNU_CB_SIZE_MEG.Caption, " Meg", "")
    TT = Replace(TT, " Image", "")
    TT = Replace(TT, " Text", "")
    TT1 = Replace(MNU_CB_SIZE_BYTE.Caption, " Byte", "")
    TT1 = Replace(TT1, " Image", "")
    TT1 = Replace(TT1, " Text", "")
    
    If InStr(MNU_CB_SIZE_MEG.Caption, "Image") > 0 Then
        cat = "ClipBoard = Image"
    Else
        cat = "Clipboard = Text"
    End If
    
    Mnu_LAST_CLIP_TIME.Caption = Format(Now, "MMMM DD-MMM-YYYY -- HH:MM:SS")
    TTX = TTX + MNU_CB_SIZE_MEG.Caption + vbCrLf + TT + vbCrLf
    TTX = TTX + MNU_CB_SIZE_BYTE.Caption + vbCrLf + TT1
    TTX = Replace(TTX, " Image", "")
    TTX = Replace(TTX, " Text", "")
    TTX = cat + vbCrLf + TTX
    
    
    EXECUTE_TIMER_ENABLED = False
Clipboard.Clear
    EXECUTE_TIMER_ENABLED = True
    Clipboard.SetText TTX

    Me.Refresh
    DoEvents


End Sub

Private Sub Mnu_Explorer_Sound_1_Click()

Me.WindowState = vbMinimized

Shell "Explorer.exe /select," + Path_Of_Sound_File, vbNormalFocus


End Sub

Private Sub Mnu_Explorer_Sound_2_Click()

Me.WindowState = vbMinimized

Shell "Explorer.exe /select," + vPathSOUND2 + vFileSOUND2, vbNormalFocus

End Sub


Private Sub Mnu_Find_New_Sounds_Click()

'THIS GET RUN HERE FIND NEW SOUND
'AND AT ONE CLICK
'AND AT MENU RESET AUDIO
'FIRST RUN FORM LOAD
'

'------------------------------
'RELOAD OR MODIFY THE SELECTION
'------------------------------

MODIFY_SOUND_SELECTION = True
Call SETUP_SOUND_FILE("")
'Beep
'MODIFY_SOUND_SELECTION = True

'                    MMControl1.Command = "prev"
'                    MMControl1.Command = "Play"
'                    MMControLl.Command = "prev"
'                    MMControl2.Command = "Play"
                    
                    
                    

End Sub

Private Sub MNU_FORMAT_MINE_Click()

Beep

Dim EE As String

EE = LCase(AD$)

Mid$(EE, 1, 1) = UCase(Mid$(EE, 1, 1))

ReDim Preserve DIE(100)

DIE(1) = " USB"
DIE(2) = " XP"
DIE(3) = " Windows XP"
DIE(4) = " I"
DIE(5) = " I'm"
DIE(6) = vbLf + "I'm"
DIE(7) = " I'll"
DIE(8) = vbLf + "I'll"
DIE(9) = vbLf + "I"
DIE(10) = " HDD"
DIE(11) = " BIOS"
DIE(12) = " NTFS"
DIE(13) = " UBUNTU"
DIE(14) = " HW"
DIE(15) = " CH"
DIE(16) = " HBCD"
DIE(17) = " LINUX"
DIE(18) = " XBOOT"
DIE(19) = " VGA"
DIE(20) = " ISO"
DIE(21) = " DigiWiz"
DIE(21) = " Grub4Dos"
DIE(22) = " GoodSync"
DIE(23) = " RoboForm"
DIE(24) = " LSD"
DIE(25) = " LED"
DIE(26) = " LCD"
DIE(27) = " NHS"

ReDim Preserve DIE(27)
Dim O_LEN_EE
'DO THE COMMON
Do
    O_LEN_EE = EE
    EE = Replace(LCase(EE), LCase("DON;T"), "Don't")
    EE = Replace(LCase(EE), LCase("Donºt"), "Don't")
Loop Until O_LEN_EE = EE


'IE SET ABLE TO GO EITHER WAY BUTTER LCASE WOULD HAVE TO BE SET AGAIN FROM ABOVE
'-------------------------------------------------------------------------------
'DO THE DIE SET
For r = 1 To UBound(DIE)
    Do
        O_LEN_EE = EE
        EE = Replace(EE, LCase(DIE(r)) + " ", DIE(r) + " ")
        EE = Replace(EE, LCase(DIE(r)) + vbCr, DIE(r) + vbCr)
        EE = Replace(EE, LCase(DIE(r)) + vbLf, DIE(r) + vbLf)
    Loop Until O_LEN_EE = EE
Next

'DO THE FIRST UPPER CASE
For r = 65 To 65 + 25
    Do
        O_LEN_EE = EE
        EE = Replace(EE, " " + LCase(Chr(r)), " " + UCase(Chr(r)))
        EE = Replace(EE, vbLf + LCase(Chr(r)), vbLf + UCase(Chr(r)))
        EE = Replace(EE, vbCr + LCase(Chr(r)), vbCr + UCase(Chr(r)))
        EE = Replace(EE, "-" + LCase(Chr(r)), "-" + UCase(Chr(r)))
        EE = Replace(EE, "(" + LCase(Chr(r)), "-" + UCase(Chr(r)))
    Loop Until O_LEN_EE = EE
Next



AD$ = EE
   
EXECUTE_TIMER_ENABLED = False
    Clipboard.Clear
    Clipboard.SetText AD$
EXECUTE_TIMER_ENABLED = True

Beep

AD_DATE = 0

Me.WindowState = 1

End Sub

Private Sub MNU_FORMAT_PLAIN_TEXT_Click()
    
    EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        Clipboard.SetText AD$
    EXECUTE_TIMER_ENABLED = True
    Me.WindowState = vbMinimized
    AD_DATE = 0
    Beep

End Sub


Private Sub MNU_CLIP_TIME_Click()

EXECUTE_TIMER_ENABLED = False
TIME_XYZ = Format(Now, "DDD DD MMMM YYYY HH:MM:SS") + "----------"
Clipboard.Clear
Clipboard.SetText TIME_XYZ
EXECUTE_TIMER_ENABLED = True

'MsgBox "THE TIME IS CLIPBOARD" + vbCrLf + TIME_XYZ
Beep

Me.WindowState = vbMinimized
AD_DATE = 0

End Sub


Private Sub MNU_FORMAT_PLAIN_TEXT_LARGE_CLIPBOARD_Click()
Dim TEMPVAR$
TEMPVAR$ = Clipboard.GetText

    EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        Clipboard.SetText TEMPVAR$
    EXECUTE_TIMER_ENABLED = True
    
'Clipboard.Clear
'Clipboard.SetText TEMPVAR$
TEMPVAR$ = ""
End Sub


Private Sub MNU_INFO_RAPID_MYUSER_Click()

'ii = " /nologo /WithSubdirectoriesYES /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\ /File\*.TXT"
'ii = " /nologo /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\"

ii = GetShortName(PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\")
ii = " /nologo /Dir" + ii

Shell "C:\Program Files\seRapid\seRapid.exe " + ii, vbNormalFocus
Me.WindowState = 1

End Sub

Private Sub MNU_LAST_GRAB_CAPS_Click()

'RRR = Now + TimeSerial(0, 0, 3)

AD$ = UCase(AD$)

    EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        Clipboard.SetText AD$
    EXECUTE_TIMER_ENABLED = True

Me.WindowState = vbMinimized
AD_DATE = 0
Beep

End Sub

Private Sub MNU_LAST_WEBCAM_PIC_Click()

Call LAST_IMAGE("EXPLORER", "D:\0 00 ART LOGGERS - WEBCAM\WEBCAM\") '=EXPLORER

Exit Sub
'ScanPath.SHOW

'XdATE2 = 0
'ScanPath.chk_LIST_VIEW_SHORT_5 = vbChecked
LAST_FILE_DATE_PATH = ""
LAST_FILE_DATE_TIME = DateSerial(100, 1, 1)

ScanPath.cboMask = "*.JPG"
ScanPath.chkSubFolders = vbChecked

ScanPath.TxtPath = "D:\0 00 Art Loggers - WEBCAM\"

ScanPath.chkCopyMemory.Value = vbChecked

Call ScanPath.CMDScan_NO_LIST_FAST_Click

FileSpec = LAST_FILE_DATE_PATH 'ScanPath.lblCount7

If FileSpec = "" Then MsgBox "NOT ANY OF THOSE FILES" + vbCrLf + ScanPath.TxtPath + "\" + ScanPath.cboMask: Exit Sub

'FileSpec = LAST_FILE_DATE_PATH_HOT_KEY_SCREENSHOT

'Set F = FS.getfile((Filespec1))
'ADATE1 = F.datelastmodified

'ScanPath.lblCount7 = ""
'ScanPath.ListView1.ListItems.Clear

'ScanPath.TxtPath = "D:\0 00 Art Loggers - WEBCAM\"
'Call ScanPath.cmdScan_Click


'Filespec2 = ScanPath.lblCount7
'If Filespec2 <> "" Then
'    Set F = FS.getfile((Filespec2))
'    ADATE2 = F.datelastmodified
'    If ADATE1 > ADATE2 Then
'        FileSpec = Filespec1
'    Else
'        FileSpec = Filespec2
'
'    End If
'Else
'    FileSpec = Filespec1
'End If

'Me.WindowState = vbMinimized
If MNU_MESSAGE_BOXES.Checked = False Then
    'MsgBox "FOUND LATEST IMAGE Clipboarded - LOAD Explorer Minimized AS Well as IrFan Maximized To View" + vbCrLf + "FILES FOUND =" + str(tFileCount) + vbCrLf, vbMsgBoxSetForeground
End If
'Me.WindowState = vbMinimized



Shell "Explorer.exe /select, " + FileSpec, vbMinimizedNoFocus

'If IRFANVIEW_PATH <> "" Then
'    If VAR1 = "IVIEW" Then
'        Shell IRFANVIEW_PATH + FileSpec + """ /fs /silent", vbMaximizedFocus
'    End If
'Else
'    Me.WindowState = vbNormal
'    MsgBox "IRFANVIEW_PATH VAR -- PATH NOT FOUND FOR FILE" + vbCrLf + "NOT INSTALED AT EXPECTED LOCATION " + IRFANVIEW_PATH3 + vbCrLf + "OR" + vbCrLf + IRFANVIEW_PATH2, vbMsgBoxSetForeground
'End If

'Me.WindowState = vbMinimized

End Sub

Private Sub MNU_MESSAGE_BOXES_Click()
    
    If MNU_MESSAGE_BOXES.Checked = True Then MNU_MESSAGE_BOXES.Checked = False: Exit Sub
    
    MNU_MESSAGE_BOXES.Checked = True

End Sub

Private Sub MNU_PLAY_SOUND_ON_LOAD_Click()

    'ONLY A SOUND TEST
    'IF A SOUND REQUIRED AS WHEN DUPE CLIPBOARD WILL MAKE DOUBLE SOUND

    If MNU_PLAY_SOUND_ON_LOAD.Checked = True Then MNU_PLAY_SOUND_ON_LOAD.Checked = False: Exit Sub
    
    MNU_PLAY_SOUND_ON_LOAD.Checked = True

End Sub


Private Sub MNU_PROCAPS_Click()
'    RRR = Now + TimeSerial(0, 0, 3)

    Call Mnu_Fix1st_Letters_Click
    
    Me.WindowState = vbMinimized
End Sub

Private Sub Mnu_Fix1st_Letters_Click()
'HERE IS -- PROCAPS
    
Dim EE As String

'EE = Text1.Text
EE = AD$
   
'If EGG = 0 Then EE = LCase(EE)
EE = LCase(EE)

Mid$(EE, 1, 1) = UCase(Mid$(EE, 1, 1))

For r = 65 To 65 + 25
    EE = Replace(EE, " " + LCase(Chr(r)), " " + UCase(Chr(r)))
    EE = Replace(EE, vbLf + LCase(Chr(r)), vbLf + UCase(Chr(r)))
    EE = Replace(EE, vbCr + LCase(Chr(r)), vbCr + UCase(Chr(r)))
    EE = Replace(EE, "-" + LCase(Chr(r)), "-" + UCase(Chr(r)))
    EE = Replace(EE, "(" + LCase(Chr(r)), "-" + UCase(Chr(r)))
    EE = Replace(EE, "." + LCase(Chr(r)), "." + UCase(Chr(r)))
    EE = Replace(EE, "\" + LCase(Chr(r)), "\" + UCase(Chr(r)))
    EE = Replace(EE, "_" + LCase(Chr(r)), "_" + UCase(Chr(r)))
    EE = Replace(EE, """" + LCase(Chr(r)), """" + UCase(Chr(r)))
Next

'HERE IS -- PROCAPS

AD$ = EE
    
    EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
    EXECUTE_TIMER_ENABLED = True
    EXECUTE_TIMER_ENABLED = False
        Clipboard.SetText AD$
    EXECUTE_TIMER_ENABLED = True

AD_DATE = 0

Me.WindowState = 1

'Text1.Text = EE

End Sub



Private Sub MNU_100_Click()
TimerSCROLL.Enabled = True
TimerSCROLL.Interval = 100
For Each Control In Controls
If InStr(Control.Name, "MNU_") > 0 And InStr(Control.Name, "00") > 0 Then
    Control.Checked = False
End If
Next
MNU_100.Checked = True
End Sub

Private Sub MNU_200_Click()
TimerSCROLL.Enabled = True
TimerSCROLL.Interval = 200
For Each Control In Controls
If InStr(Control.Name, "MNU_") > 0 And InStr(Control.Name, "00") > 0 Then
    Control.Checked = False
End If
Next
MNU_200.Checked = True

End Sub

Private Sub MNU_300_Click()
TimerSCROLL.Enabled = True
TimerSCROLL.Interval = 300
For Each Control In Controls
If InStr(Control.Name, "MNU_") > 0 And InStr(Control.Name, "00") > 0 Then
    Control.Checked = False
End If
Next

MNU_300.Checked = True

End Sub

Private Sub MNU_400_Click()
TimerSCROLL.Enabled = True
TimerSCROLL.Interval = 400
For Each Control In Controls
If InStr(Control.Name, "MNU_") > 0 And InStr(Control.Name, "00") > 0 Then
    Control.Checked = False
End If
Next
MNU_400.Checked = True

End Sub

Private Sub MNU_500_Click()
TimerSCROLL.Enabled = True
TimerSCROLL.Interval = 500
For Each Control In Controls
If InStr(Control.Name, "MNU_") > 0 And InStr(Control.Name, "00") > 0 Then
    Control.Checked = False
End If
Next
MNU_500.Checked = True

End Sub

Private Sub MNU_800_Click()
TimerSCROLL.Enabled = True
TimerSCROLL.Interval = 800
For Each Control In Controls
If InStr(Control.Name, "MNU_") > 0 And InStr(Control.Name, "00") > 0 Then
    Control.Checked = False
End If
Next
MNU_800.Checked = True

End Sub

Private Sub Mnu_Clear_Text_Click()
'Call Command3_Click
End Sub

Private Sub Mnu_ClearClipBoard_Click()
'Call Command2_Click
End Sub

Private Sub Mnu_ClipAll_Click()
'Call Command1_Click
End Sub

Private Sub Mnu_Test_Clip_Click()
'Call Command4_Click
End Sub

Private Sub Mnu_Edit_Sound_Click()

Dim TT As String

TT = Path_Of_Sound_File

TT = FindShortPath(TT)

Shell """C:\Program Files\Cool2000\cool2000.exe"" """ + TT + """", vbNormalFocus

End Sub

Private Sub Mnu_Exit_Click()

'MMControl1.Command = "close"
'
'FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA+"\#OutClipChunck.Txt"
'DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
'FR1 = FreeFile
'Open FILENAME_IN_USE_CHECK For Output As #FR1
'Print #FR1, AD$;
'Close #FR1
Beep

EXIT_TRUE = True

Unload Form1


End Sub

Private Sub MNU_INFO_RAPID_Click()


'ii = " /nologo /DirD:\#*MY*DOCS\01*My*Documents\03*NotePad*Files\00*TOP\  /FileClued*Up.txt"


ii = " /nologo /Dir" + App.Path + "\#*DATA\" + GetComputerName + "\  /File00_ClipBoard_Total--TRIM.txt"
ii = Replace(ii, "ClipBoard Logger", "Fast*Clipboard")
ii = Replace(ii, "\# Data", "\#*Data")

Shell "C:\Program Files\seRapid\seRapid.exe " + ii, vbNormalFocus

Me.WindowState = 1

End Sub

Private Sub MNU_INFRO_RAPID_ALL_Click()

ii = " /nologo /WithSubdirectoriesYES /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\ /File\*.TXT"
ii = " /nologo /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\"

Shell "C:\Program Files\seRapid\seRapid.exe " + ii, vbNormalFocus
Me.WindowState = 1


'ClipBoard-

End Sub

Private Sub MNU_INFRO_RAPID_ALL_SMALL_FILES_Click()


Timer_INFORAPID_MSGBOX.Enabled = True


ii = " /nologo /WithSubdirectoriesYES /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\ /FileClipBoard-**.TXT"
ii = " /nologo /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\ /FileClipBoard-*.TXT"

'ii = " /nologo /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\"

Shell "C:\Program Files\seRapid\seRapid.exe " + ii, vbNormalFocus
Me.WindowState = 1


'ClipBoard-


End Sub

Private Sub MNU_LAST_ART_PIC_Click()

Me.WindowState = vbMinimized

Call LAST_IMAGE("EXPLORER", "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\Hot-Key-App-Shots\") '=EXPLORER


End Sub

Private Sub Mnu_LoggExplorer_Click()

Me.WindowState = vbMinimized
Shell "Explorer.exe /e," + PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + """", vbNormalFocus


End Sub



Sub TOT()

End Sub


'---------------------------------------------------------------

'http://www.vbforums.com/showthread.php?673134-RESOLVED-Minimum-height-for-menu-bar-to-be-visible
'-------------- MENU SIZE DECLARE
'Private Type rect
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'Private Type MENUBARINFO
'  cbSize As Long
'  rcBar As rect
'  hMenu As Long
'  hwndMenu As Long
'  fBarFocused As Boolean
'  fFocused As Boolean
'End Type
'Private MenuInfo As MENUBARINFO
'Private Const OBJID_MENU As Long = &HFFFFFFFD
'Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
'Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hwnd As Long, _
'ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
'Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Private Function Menu_Height()
 
    MenuInfo.cbSize = Len(MenuInfo)
    
    If GetMenuBarInfo(Me.hwnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            'Menu_Height = CStr(.Top) + CStr(.Bottom)
            Menu_Height = CStr(.Bottom) - CStr(.Top)
        End With
       
    End If
   
End Function
'------------------------------------------

Private Sub Mnu_Open_Logg_Click()
'Shell "notepad " + PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA+"\ClipBoard-" + Format$(Start, "YYYY-MM-DD") + ".Txt", vbNormalFocus


'WON'T DO THIS
'Shell "EXPLORER " + PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA+"\ClipBoard-" + Format$(Start, "YYYY-MM-DD") + ".Txt", vbNormalFocus
'OR THIS
'Shell "CMD /C ""START /MAX """ + PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA+"\ClipBoard-" + Format$(Start, "YYYY-MM-DD") + ".Txt""""", vbNormalFocus

'ANSWER
vFile = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA + "\ClipBoard-" + Format$(start, "YYYY-MM-DD") + ".Txt"

Me.WindowState = vbMinimized

ShellExecute Me.hwnd, "open", vFile, vbNullString, vbNullString, conSwNormal

End Sub

Private Sub Mnu_Open_Recent_Click()
Me.WindowState = vbMinimized

'Shell "notepad " + App.Path + "\# DATA\"+GetComputerName + "\Data\00_ClipBoard_Total--TRIM.txt", vbNormalFocus
vFile = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total--TRIM.txt"
ShellExecute Me.hwnd, "open", vFile, vbNullString, vbNullString, conSwNormal

End Sub

Private Sub Mnu_Open_Recent_Hex_Click()
Me.WindowState = vbMinimized

'C:\Program Files\XVI32\XVI32.exe
Shell "C:\Program Files\XVI32\XVI32.exe """ + PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total--TRIM.txt""", vbNormalFocus


End Sub

Private Sub Mnu_Open_Total_Click()
Me.WindowState = vbMinimized

'Shell "notepad " + App.Path + "\# DATA\"+GetComputerName + "\Data\00_ClipBoard_Total.txt", vbNormalFocus

vFile = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total.txt"
ShellExecute Me.hwnd, "open", vFile, vbNullString, vbNullString, conSwNormal

End Sub




Private Sub MNU_SCREEN_SHOT_Click()

Beep

Me.WindowState = vbMinimized

ART_FOLDER = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName


Shell "Explorer.exe /e, " + """" + ART_FOLDER + """", vbMaximizedFocus

'If Dir("M:\0 00 Art Loggers\", vbDirectory) <> "" And 1 = 2 Then
    'Shell "Explorer.exe /e, M:\0 00 Art Loggers\", vbNormalFocus
'End If

Me.WindowState = vbMinimized


End Sub

Private Sub MNU_SCROLL_DOWN_Click()
TimerSCROLL.Enabled = Not TimerSCROLL.Enabled
End Sub

Private Sub MNU_SCROLL_OFF_Click()
TimerSCROLL.Enabled = False
End Sub

Private Sub Mnu_Shell_T_Click()

If Dir("D:\Utils\T.com") = "" Then
    MsgBox "D:\UTILS\T.COM" + vbCrLf + "PROGRAM DON'T EXIST" + vbCrLf + " I SET THIS MENU OPTION TO USE " + vbCrLf + "LIST Enhanced v2.4y1 (c) 1990-2005 Vernon D. Buerg - All rights reserved" + vbCrLf + "Matthew Lancaster, Single-User, s/n ######" + vbCrLf + "Unauthorized duplication prohibited." + vbCrLf + "A FREE VERSION YOU CAN GET - ITS A TEXT BASED VIEWER FROM DOS6.22 DAYS AND HANDLES LONG FILE NAMES" + vbCrLf + "I CALL IT RENAMED FROM LIST.COM TO T.COM -- T FOR TEXT -- IN DIRECTORY D:\UTILS\T.COM", vbMsgBoxSetForeground
    Exit Sub
End If

EE$ = GetShortName(PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total.txt")
Shell "D:\Utils\T.com " + EE$, vbNormalFocus
End Sub

Private Sub Mnu_ShellTRecent_Click()

End Sub

Private Sub Mnu_ShellT_Todays_Click()

If Dir("D:\Utils\T.com") = "" Then
    MsgBox "D:\UTILS\T.COM" + vbCrLf + "PROGRAM DON'T EXIST" + vbCrLf + " I SET THIS MENU OPTION TO USE " + vbCrLf + "LIST Enhanced v2.4y1 (c) 1990-2005 Vernon D. Buerg - All rights reserved" + vbCrLf + "Matthew Lancaster, Single-User, s/n ######" + vbCrLf + "Unauthorized duplication prohibited." + vbCrLf + "A FREE VERSION YOU CAN GET - ITS A TEXT BASED VIEWER FROM DOS6.22 DAYS AND HANDLES LONG FILE NAMES" + vbCrLf + "I CALL IT RENAMED FROM LIST.COM TO T.COM -- T FOR TEXT -- IN DIRECTORY D:\UTILS\T.COM", vbMsgBoxSetForeground
    Exit Sub
End If

vFile = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA + "\ClipBoard-" + Format$(start, "YYYY-MM-DD") + ".Txt"

'EE$ = GetShortName(App.Path + "\# DATA\"+GetComputerName + "\Data\#ClipBoard.Txt")

EE$ = GetShortName(vFile)
Shell "D:\Utils\T.com " + EE$, vbNormalFocus

End Sub


Private Sub MNU_SOUND_OPTION_Click(Index As Integer)



'THIS GET RUN HERE AT CLICK
'AND AT MENU RESET AUDIO
'FIRST RUN
'AND FIND NEW SOUND



'If MNU_SOUND_OPTION(Index).Checked = True Then Exit Sub
If MNU_SOUND_OPTION(Index).Enabled = False Then Exit Sub
If MNU_SOUND_OPTION(Index).Visible = False Then Exit Sub


For Each Control In MNU_SOUND_OPTION
        'Control.Enabled = False
        Control.Checked = False
'        Debug.Print Control.Index
        
        DoEvents
Next

MNU_SOUND_OPTION(Index).Checked = True


'THIS GET RUN HERE AT CLICK MENU
'AND AT RESET AUDIO



For Each Control In MNU_SOUND_OPTION
    If Control.Caption <> "MNU_SOUND_OPTION" Then
        Control.Enabled = True
    End If
Next

MODIFY_SOUND_SELECTION = True

Call SETUP_SOUND_FILE("")



'--------------
Exit Sub
'--------------



'For Each Control In Controls
'    If InStr(Control.Name, "MNU_SOUND_OPTION") > 0 Then
'        If Control.Caption = MNU_SOUND_OPTION(Index).Caption Then Exit Sub
'        Control.Enabled = False
'    End If
'Next

For Each Control In MNU_SOUND_OPTION
    
    'If InStr(Control.Name, "MNU_SOUND_OPTION") > 0 Then
        'If Control.Enabled = False Then Exit
        
        Control.Enabled = False
    'End If
Next

MNU_SOUND_OPTION(Index).Checked = True
For Each Control In MNU_SOUND_OPTION
    If Control.Caption <> "MNU_SOUND_OPTION" Then
        Control.Enabled = True
    End If
Next
'MNU_SOUND_OPTION(Index).Enabled = True



End Sub

Private Sub Mnu_SoundOn_Click()

If Mnu_SoundOn.Checked = True Then

    Mnu_Audio_ON.Caption = "** /\/\ Audio Is Off /\/\ ** Hiit On Here for ON ***"
    
    Mnu_SoundOn.Checked = False
    Exit Sub

End If

Mnu_SoundOn.Checked = True

End Sub

Private Sub MNU_SPACE_Click()
    
    Me.WindowState = vbMinimized
        EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        Clipboard.SetText " "
    EXECUTE_TIMER_ENABLED = True
'    Clipboard.Clear
'    Clipboard.SetText " "
    
    If MNU_MESSAGE_BOXES.Checked = True Then Exit Sub
    
    MsgBox "CLIPBOARDED YOU A SPACE CHARACTER"
    
    'Me.WindowState = 1

End Sub

Private Sub Mnu_StripLogg_Click()
'    Shell "notepad " + App.Path + "\# DATA\"+GetComputerName + "\Data\00_ClipBoard_Tot_Strip.txt", vbNormalFocus
    
    Me.WindowState = vbMinimized
    
    vFile = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Tot_Strip.txt"
    ShellExecute Me.hwnd, "open", vFile, vbNullString, vbNullString, conSwNormal
    
End Sub





Private Sub Mnu_TEXT_Open_Logg_Click()
If Dir("C:\Program Files\TextView\Textview.exe") = "" Then
    DATA5 = "C:\Program Files\TextView\Textview.exe" + vbCrLf
    DATA5 = DATA5 + "PROGRAM DON'T EXIST" + vbCrLf
    DATA5 = DATA5 + "I SET THIS MENU OPTION TO USE THAT" + vbCrLf
    DATA5 = DATA5 + "----------------------------------" + vbCrLf
    DATA5 = DATA5 + "Textview 6.0 - The Explorer-like Text File Viewer" + vbCrLf
    DATA5 = DATA5 + "Textview 6.0.12" + vbCrLf
    DATA5 = DATA5 + "(c) Florian Balmer 1996-2004" + vbCrLf
    DATA5 = DATA5 + "http://www.flos-freeware.ch" + vbCrLf
    MsgBox DATA5, vbMsgBoxSetForeground
    Exit Sub
End If

vFile = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA + "\ClipBoard-" + Format$(start, "YYYY-MM-DD") + ".Txt"

Shell "C:\Program Files\TextView\Textview.exe """ + vFile + """", vbMaximizedFocus

End Sub

Private Sub Mnu_TEXT_Open_Recent_Click()

If Dir("C:\Program Files\TextView\Textview.exe") = "" Then
    DATA5 = "C:\Program Files\TextView\Textview.exe" + vbCrLf
    DATA5 = DATA5 + "PROGRAM DON'T EXIST" + vbCrLf
    DATA5 = DATA5 + "I SET THIS MENU OPTION TO USE THAT" + vbCrLf
    DATA5 = DATA5 + "----------------------------------" + vbCrLf
    DATA5 = DATA5 + "Textview 6.0 - The Explorer-like Text File Viewer" + vbCrLf
    DATA5 = DATA5 + "Textview 6.0.12" + vbCrLf
    DATA5 = DATA5 + "(c) Florian Balmer 1996-2004" + vbCrLf
    DATA5 = DATA5 + "http://www.flos-freeware.ch" + vbCrLf
    MsgBox DATA5, vbMsgBoxSetForeground
    Exit Sub
End If

vFile = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total--TRIM.txt"
Shell "C:\Program Files\TextView\Textview.exe """ + vFile + """", vbMaximizedFocus

End Sub

Private Sub Mnu_TEXT_Open_Total_Click()
If Dir("C:\Program Files\TextView\Textview.exe") = "" Then
    DATA5 = "C:\Program Files\TextView\Textview.exe" + vbCrLf
    DATA5 = DATA5 + "PROGRAM DON'T EXIST" + vbCrLf
    DATA5 = DATA5 + "I SET THIS MENU OPTION TO USE THAT" + vbCrLf
    DATA5 = DATA5 + "----------------------------------" + vbCrLf
    DATA5 = DATA5 + "Textview 6.0 - The Explorer-like Text File Viewer" + vbCrLf
    DATA5 = DATA5 + "Textview 6.0.12" + vbCrLf
    DATA5 = DATA5 + "(c) Florian Balmer 1996-2004" + vbCrLf
    DATA5 = DATA5 + "http://www.flos-freeware.ch" + vbCrLf
    MsgBox DATA5, vbMsgBoxSetForeground
    Exit Sub
End If

vFile = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total.txt"
Shell "C:\Program Files\TextView\Textview.exe """ + vFile + """", vbMaximizedFocus

End Sub

Private Sub Mnu_TEXT_StripLogg_Click()

Dim Checked_Var
If Dir("C:\Program Files\TextView\Textview.exe") = "" Then
    Checked_Var = "C:\Program Files\TextView\Textview.exe"
End If

If Dir("C:\Program Files (x86)\TextView\Textview.exe") = "" And Checked_Var = "" Then
    Checked_Var = "C:\Program Files (x86)\TextView\Textview.exe"
End If
    
If Checked_Var = "" Then
    
    DATA5 = "C:\Program Files\TextView\Textview.exe" + vbCrLf
    DATA5 = DATA5 + "PROGRAM DON'T EXIST" + vbCrLf
    DATA5 = DATA5 + vbCrLf
    
    DATA5 = "C:\Program Files (x86)\TextView\Textview.exe" + vbCrLf
    DATA5 = DATA5 + "PROGRAM DON'T EXIST" + vbCrLf
    DATA5 = DATA5 + vbCrLf
    
    DATA5 = DATA5 + "I SET THIS MENU OPTION TO USE THAT" + vbCrLf
    DATA5 = DATA5 + "----------------------------------" + vbCrLf
    DATA5 = DATA5 + "Textview 6.0 - The Explorer-like Text File Viewer" + vbCrLf
    DATA5 = DATA5 + "Textview 6.0.12" + vbCrLf
    DATA5 = DATA5 + "(c) Florian Balmer 1996-2004" + vbCrLf
    DATA5 = DATA5 + "http://www.flos-freeware.ch" + vbCrLf
    
    MsgBox DATA5, vbMsgBoxSetForeground
    Exit Sub

End If

Me.WindowState = vbMinimized
vFile = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Tot_Strip.txt"
Shell "C:\Program Files\TextView\Textview.exe """ + vFile + """", vbMaximizedFocus

End Sub

Private Sub MNU_URL_Browser_Click()
    
'1..MNU_CLIPBOARD_EXPLORER_FILE_FOLDER_Click
'2..MNU_REG_JUMP_Click
'3..MNU_URL_Browser_Click
'4..MNU_CPC_Click
    
    TEXT_PATH = Clipboard.GetText

    XPOS = 0
    If XPOS = 0 Then XPOS = InStrRev(LCase(TEXT_PATH), "http")
    If XPOS = 0 Then XPOS = InStrRev(LCase(TEXT_PATH), "www")
        
    TYPE_VAR = "PATH, URL OR REG_KEY"
    TYPE_VAR = "FOR USE WITH CPC"
    TYPE_VAR = "URL LINK HTTP OR WWW"
    
    If XPOS = 0 Then
        I = MsgBox("LAST CLIPBOARD NOT CONTAIN A VERIFIABLE STRING " + TYPE_VAR + " TO LOAD" + vbCrLf + vbCrLf + TEXT_PATH + vbCrLf + vbCrLf + "----" + vbCrLf + vbCrLf + "WANT MINIMIZE ", vbYesNo, vbMsgBoxSetForeground)
        If I = vbYes Then Me.WindowState = vbMinimized
        Exit Sub
    End If
   
    If XPOS > 1 Then
        TEXT_PATH = Mid(TEXT_PATH, XPOS, InStr(XPOS, TEXT_PATH, vbCr) - XPOS)
    End If
    If XPOS = 1 And InStr(XPOS, URL_PATH, vbCr) > 0 Then
        TEXT_PATH = Mid(TEXT_PATH, XPOS, InStr(XPOS, TEXT_PATH, vbCr) - XPOS)
    End If
    
    Beep
    
    '------------------------------------
    'STRIP WHITE SPACE CHARACTOR FROM URL
    '------------------------------------
    TEXT_PATH = Replace(TEXT_PATH, " ", "")
    TEXT_PATH = Replace(TEXT_PATH, vbCrLf, "")
    TEXT_PATH = Replace(TEXT_PATH, vbCr, "")
    TEXT_PATH = Replace(TEXT_PATH, vbLf, "")
    vFile = TEXT_PATH
    ShellExecute Me.hwnd, "open", vFile, vbNullString, vbNullString, conSwNormal
    ', vbMaximized
    Me.WindowState = vbMinimized
End Sub

Private Sub Mnu_Html_Url_Click()

    '<a href="LoveFolder/index.php?dir=HTML/&file=quotes_an_stuff.html">Quotes and Stuff</a>
    
    URL_PATH = Clipboard.GetText
    XPOS = InStrRev(LCase(URL_PATH), "http")
    
    If XPOS = 0 Then
        MsgBox "LAST CLIPBOARD NOT A HTTP URL TO MAKE HTML", vbMsgBoxSetForeground
        Me.WindowState = vbMinimized
        Exit Sub
    End If
   
    URL_PATH = Mid(URL_PATH, XPOS, InStr(XPOS, URL_PATH, vbCr) - XPOS)
    
    AD$ = "<a href="""
    AD$ = AD$ + URL_PATH
    AD$ = AD$ + """>" + URL_PATH + "</a>"

    'DONT REQUIRE THESE REM MOUSE MOTION WONT
    'ENTER ON CLIPBOARD
    'EXECUTE_TIMER_ENABLED = False
    EXECUTE_TIMER_ENABLED = False
    Clipboard.Clear

    EXECUTE_TIMER_ENABLED = True
'    Clipboard.Clear
    Clipboard.SetText AD$

'        EXECUTE_TIMER_ENABLED = False
'        Clipboard.Clear
'        Clipboard.SetText AD$
'    EXECUTE_TIMER_ENABLED = True
    
    'EXECUTE_TIMER_ENABLED = True
    
    AD_DATE = 0
    
    Call Menu_clipboard_size(Len(AD$))
    
    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_VB_Click()

' -----------------------------------------------------------------
' IS SHORT CUT NOT SET ADMINISTATOR FOR THIS PARTICALUR CODE IN EXE
' MODE IT WILL NOT RUN LOAD ITSELF ON THE VB COMPILER
' BUT AND WILL RUN OTHER PROGRAM
' AS TALK OTHER PROGRAM ARE OKAY BUT THEN MAYBE SET AS ADMIN
' AMSWER GIVEN
' AUG 28 SUN
' -----------------------------------------------------------------
Beep

If 1 = 1 And IsIDE = True Then
    MNU_VB.Enabled = False
    MsgBox "NOT WHEN ISIDE = TRUE"
    Exit Sub
End If

Dim ReturnHwnd As Long
Dim I
'VB ONLY WANTS THE 1ST OF THE 2 HWND
'ReturnHwnd = FindWindowPartVB("ClipBoard Logger - Microsoft Visual Basic[")

'DONT NEED ABOVE USE THIS
x1 = FindWindow("wndclass_desked_gsk", vbNullString)
'------------------------------------------------
'FIND THE WINDOW PRIZE TREASURE CHEST BUST BLOSSOM
'TRAIN SPOTTER
'------------------------------------------------
'-----------------------------------------------------------
'CHECK CLASS NAME IS VB IDE WINDOW WE WANT OR IF ANOTHER NOT
'-----------------------------------------------------------
X2 = GetWindowTitle(x1)
If InStr(X2, " - Microsoft Visual Basic [") > 0 Then
    'RUN A WINDOW SPY
    WIN_SPY_NAME = "ClipBoard Logger"
    If InStr(X2, WIN_SPY_NAME) > 0 Then
    
        MsgBox "DON'T RUN VB IDE - LOADED"
        I = GetWindowState(x1)
        If I = vbMinimized Then
            SHOWVAR = SW_SHOWDEFAULT
            ShowWindow ReturnHwnd, SHOWVAR
        End If
        
        EXIT_TRUE = True
        Unload Me
        Exit Sub
    End If
End If

'BETTER LINE NEXT DON'T USE VB MENU IF ISIDE

'------------------------------------------------
'TEMP WORK AROUND
'OVER DRIVE
'OVER RIDE
'------------------------------------------------
'FINDWINDOW PART PROBLEM IN EXE AND WHEN IN WIN 7
'------------------------------------------------
'WIN 7 PROBLEM MUST USE EXTRA LAST LEFT SQUARE BRACKET OF SERACH END ABOVE
'------------------------------------------------
If ReturnHwnd > 0 Then
    If MsgBox("ReturnHwnd" + Str(ReturnHwnd) + " -- " + "MeHwnd" + Str(Me.hwnd) + vbCrLf + "VB Code " + vbCrLf + WIN_SPY_NAME + vbCrLf + " already Running - Sure Want to Run VB IDE", vbYesNo) = vbYes Then
        WANT_TO_RUN_ANYWAY = True
    End If
End If


If ReturnHwnd > 0 Then
    
    I = GetWindowState(ReturnHwnd1)
    
    If I = vbMinimized Then
        
'        SHOWVAR = SW_RESTORE
'        SHOWVAR = SW_SHOW
'        SHOWVAR = SW_SHOWNA
'        SHOWVAR = SW_SHOWDEFAULT
'        SHOWVAR = SW_SHOWNORMAL
'        SHOWVAR = SW_SHOWMAXIMIZED
        SHOWVAR = SW_SHOWDEFAULT
        
        ShowWindow ReturnHwnd, SHOWVAR
        
        'GUESS MAYBE
        'SetForegroundWindow (ReturnHwnd)

        DoEvents
        
    End If
    
    'MADE REDUNDANT CODE BY CONDICTION HERE AND ABOVE
    If WANT_TO_RUN_ANYWAY = False Then
        MsgBox "EXIT AS FOUND WINODW OF VB AND PUT TO SHOW FOCUS"
        EXIT_TRUE = True
        Unload Me
        Exit Sub
    End If

End If


If Dir("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
End If


If Dir("C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
    'MsgBox """C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbMsgBoxSetForeground
    
    'Reset
    'Close
    EXIT_TRUE = True

    Dim Control
    'SET ALL TIMERS IN ALL FORMS ENABLED=FALSE
    On Error Resume Next
        For I = 0 To Forms.Count - 1
            For Each Control In Forms(I).Controls
                If InStr(UCase(Control.Name), "TIMER") > 0 Then
                    'Debug.Print Control.Name
                    Control.Enabled = False
                End If
            Next
        Next I
    On Error GoTo 0
       
    Dim Form As Form
    For Each Form In Forms
        If Form.Name <> Me.Name Then
            EXIT_TRUE = True
            Unload Form
            Set Form = Nothing
        End If
    Next Form
    
    'I = ExecCmdWAIT("""C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""")
    
    Shell """C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
    
    EXIT_TRUE = True
    
    Unload Me

End If


End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
RRR = Now + TimeSerial(0, 0, 3) 'STOPS THE ENTRY IN LOG 'SCROLL IF ABOVE 0 OR LOWER > TIME

If IsIDE = TRUR Then If KeyCode = 27 Then Beep: End


Exit Sub


Dim SI As SCROLLINFO




FlatSB_SetScrollPos Text1.hwnd, SB_VERT, 60, False
SI.cbSize = Len(SI)
SI.fMask = SIF_ALL
FlatSB_GetScrollInfo Text1.hwnd, SB_VERT, SI
SI.nPos = SI.nPos - 10
SI.nPos = 50
FlatSB_SetScrollInfo Text1.hwnd, SB_VERT, SI, True

'Text1.AutoRedraw = True


End Sub

Sub TEXT_SELECTOR_VALUE()

    If O_Text1_SelLength = Text1.SelLength Then Exit Sub

    MNU_COPY.Caption = "<*** COPY --" + Str(Text1.SelLength) + " ***>"
    O_Text1_SelLength = Text1.SelLength

End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Call TEXT_SELECTOR_VALUE
    
    If Button <> 1 Then Exit Sub ' Right Click
    RRR = Now + TimeSerial(0, 0, 3) 'STOPS THE ENTRY IN LOG 'SCROLL IF ABOVE 0 OR LOWER > TIME
    

    HHH = Now + TimeSerial(0, 2, 0) ' vbMinimized WINDOW
    
    XXMOUSEMOVE = Now + TimeSerial(0, 2, 0) ' RESET THE SCROLL TO BOTTOM OF TEXT BOX
        
End Sub
Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Call TEXT_SELECTOR_VALUE

    If MNU_SLECTOR_WHOLE_LINE_MODE.Caption = FULL_LINE + "=TRUE" Then
        AI_1 = Text1.SelStart
        AI_2 = InStrRev(Text1.Text, vbCrLf, AI_1)
        AI_3 = InStr(AI_2 + 1, Text1.Text, vbCrLf)
        Text1.SelStart = AI_2 + 1
        Text1.SelLength = AI_3 - AI_2 - 1
        'Text1.SE
        
        If Text1.SelLength > 2 Then Call CLIPBOARD_TEXT_OF_SELECTED_AUTO
    End If

End Sub


Sub CLIPBOARD_TEXT_OF_SELECTED_AUTO()
    '----
    'Copy rich text box's text - Xtreme Visual Basic Talk
    'http://www.xtremevbtalk.com/general/15668-copy-rich-text-boxs-text.html
    '----
    
    MNU_COPY.Caption = "<*** COPY --" + Str(Text1.SelLength) + " ***>"
    
    If Text1.SelLength = 0 Then Beep: Exit Sub
    
    Const WM_COPY = &H301
    EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        SendMessage Text1.hwnd, WM_COPY, 0&, 0& 'Copy
    EXECUTE_TIMER_ENABLED = True
    
    Beep
    
    On Error Resume Next
    Do
        Err.Clear
        AD$ = Clipboard.GetText
        Sleep 100
        'If Timer Mod 5 = 0 Then DoEvents
    Loop Until Err.Number = 0

    Call COMPARE_FOR_EXE
End Sub


Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Exit Sub
    Call TEXT_SELECTOR_VALUE

    RRR = Now + TimeSerial(0, 0, 3) 'STOPS THE ENTRY IN LOG 'SCROLL IF ABOVE 0 OR LOWER > TIME
    HHH = Now + TimeSerial(0, 2, 0) ' vbMinimized WINDOW
    
    XXMOUSEMOVE = Now + TimeSerial(0, 2, 0) ' RESET THE SCROLL TO BOTTOM OF TEXT BOX

    
    
    
    TRIPPING = " Byte's Selected"
    If MNU_SELECTED.Caption <> Trim(Str(Text1.SelLength)) + TRIPPING Then
        MNU_SELECTED.Caption = Trim(Str(Text1.SelLength)) + TRIPPING
    End If
    
    
    'MNU_SELECTED


End Sub


Sub GetFormat_And_Display()
'
'If FirstRun = False Then
'    Mnu_Run_Status.Visible = False
'    Mnu_Clip_Status.Visible = False
'End If

On Error GoTo EXIT_SUB

I = False
ClipFormatDescription = ""

If I = False Then I = Clipboard.GetFormat(vbCFText)
'"Clip Format:- "+"Text (.txt file)"
If I = True And ClipFormatDescription = "" Then
    ClipFormatDescription = "Text (.txt file)"
    ClipFormatDesc2 = "Text"
End If
' Check for Text Before Rich Text Better Option - But Check if You Can Grab Rich Text into Plain Text

' Two Most Important First
If I = False Then I = Clipboard.GetFormat(vbCFBitmap)
If I = True And ClipFormatDescription = "" Then
    ClipFormatDescription = "Bitmap (.bmp file)"
    ClipFormatDesc2 = "Image"
End If

If I = False Then I = Clipboard.GetFormat(vbCFRTF)
If I = True And ClipFormatDescription = "" Then
    ClipFormatDescription = "Rich Text Format (.rtf file)"
    ClipFormatDesc2 = "Rich Text Format File"
End If
If I = False Then I = Clipboard.GetFormat(vbCFLink)
If I = True And ClipFormatDescription = "" Then
    ClipFormatDescription = "DDE Conversation information"
    ClipFormatDesc2 = "DDE Conversation info"
End If

If I = False Then I = Clipboard.GetFormat(vbCFMetafile)
If I = True Then
    ClipFormatDescription = "Metafile (.wmf file)"
    ClipFormatDesc2 = "Metafile .WMF"
End If
If I = False Then I = Clipboard.GetFormat(vbCFEMetafile)
If I = True And ClipFormatDescription = "" Then
    ClipFormatDescription = "Device-independent bitmap - E Type"
    ClipFormatDesc2 = "Device-independent bitmap - E Type"
End If
If I = False Then I = Clipboard.GetFormat(vbCFDIB)
If I = True And ClipFormatDescription = "" Then
    ClipFormatDescription = "Device-independent bitmap"
    ClipFormatDesc2 = "Device-independent bitmap"
End If
If I = False Then I = Clipboard.GetFormat(vbCFPalette)
If I = True And ClipFormatDescription = "" Then
    ClipFormatDescription = "Color palette"
    ClipFormatDesc2 = "Color Palette"
End If
If I = False Then I = Clipboard.GetFormat(vbCFFiles)
If I = True And ClipFormatDescription = "" Then
    ClipFormatDescription = "File list from Windows Explorer"
    ClipFormatDesc2 = "File list from Windows Explorer"
End If
If I = False Then I = Clipboard.GetFormat(-15694) 'Pasted Objects on VB IDE Form Designer
If I = True And ClipFormatDescription = "" Then
    ClipFormatDescription = "VB IDE (Form Designer) Object -15694"
    ClipFormatDesc2 = "VB IDE (Form Designer) Object -15694"
End If

'Check the Most Common 1st

'And Then Check All - Don't Let Any Escape

iindex = -32000
If I = False Then
    For r = -30000 To 32000
        If I = False Then
            I = Clipboard.GetFormat(r)
            iindex = r
        Else
            Exit For
        End If
    Next
End If

If I = True And iindex <> -32000 And ClipFormatDescription = "" Then
    ClipFormatDescription = "Format Unknown #" + Str(iindex) + " of -- Between -30000 and 32000"
    ClipFormatDesc2 = "Format Unknown #" + Str(iindex) + " of -- Between -30000 and 32000"
End If
'If i = False Then ClipFormatDescription = "Null Void Empty Clipboard"
If I = False Then
    ClipFormatDescription = "Empty Clipboard"
    ClipFormatDesc2 = "Empty Clipboard"
End If

If Clipboard.GetFormat(vbCFText) Then
    If Len(Clipboard.GetText) = 0 Then
        ClipFormatDescription = "Text And Len(0) Empty Clipboard"
        ClipFormatDesc2 = "Text And Len(0) Empty Clipboard"
    End If
End If

'Mnu_Clip_Description.Caption = "Clip Format:- " + ClipFormatDescription
Mnu_Clip_Description.Caption = "Clip Format:- " + ClipFormatDesc2
If TIME_LAST_CLIPBOARD = 0 Then
    Mnu_Last_Clipboard_Timer.Caption = "Last Clipboard = " + ClipFormatDesc2 + " @ Start Buffer Process Different Than Before"
Else
    Mnu_Last_Clipboard_Timer.Caption = "Last Clipboard = " + ClipFormatDesc2 + " -- 0 Secs"
End If


If ClipFormatDescription = "Text And Len(0) Empty Clipboard" Then
    Debug.Print "ClipFormatDescription = ""Text And Len(0) Empty Clipboard"
    'If IsIDE = True Then Stop
End If


'--------- ON ERROR CALL
EXIT_SUB:

End Sub


Sub COMPARE_FOR_EXE()

On Error GoTo end_error_trap

If Clipboard.GetFormat(vbCFText) = True Then
    
    If Clipboard.GetText <> COMPARE_EXE_1 Then
        
'        Debug.Print Len(Clipboard.GetText)
'        Debug.Print Len(COMPARE_EXE_3)
'        Debug.Print Len(COMPARE_EXE_1)
        
        COMPARE_EXE_3 = COMPARE_EXE_1
    
        COMPARE_EXE_1 = Clipboard.GetText
    End If
    
    'If COMPARE_EXE_2 = COMPARE_EXE_1 Then
    '    COMPARE_EXE_2 = COMPARE_EXE_3
    'End If
    'COMPARE3 = ""
    
    MNU_COMPARE.Caption = "COMPARE" + Str(Len(COMPARE_EXE_1)) + " " + Str(Len(COMPARE_EXE_3))

    Else
    COMPARE3 = ""
    
    MNU_COMPARE.Caption = "COMPARE" + Str(Len(COMPARE_EXE_1)) + " " + Str(Len(COMPARE_EXE_2))

End If

'If COMPARE_EXE_1 = "" And AD$ <> "" Then
'    COMPARE_EXE_2 = COMPARE_EXE_1
'    COMPARE_EXE_1 = AD$
'    If COMPARE_EXE_2 = COMPARE_EXE_1 Then
'
'    End If
'    MNU_COMPARE.Caption = "COMPARE" + Str(Len(COMPARE_EXE_1)) + " " + Str(Len(COMPARE_EXE_2))
'
'End If

'Problem retry clipboard
Exit Sub
end_error_trap:

Sleep 200
Resume




End Sub


Public Sub Timer1_Timer()

'Exit Sub

'DateTimerCheckIntegrityOfProgram = Now
TIME_LAST_CLIPBOARD_Timer1 = Now

Dim GG$

Dim GetForegroundWindow_XXRT, XXR5, XXR7
Dim ASH2$
Dim Ash$

TEXT_TO_ENTER_IN_LOGGER = True

VAR_INFO_LOGG_CLIP = ""

TIMER1_LAST_DATE = Now

If FlagTestClipBoardRoutine = True Then
    FlagTestClipBoardRoutine = False
    MsgBox "Clipboarded Trigger Flag Has Found to Be Working Upto Here The Entry to Routine", vbMsgBoxSetForeground
End If

ADTEST_BEFORE$ = ""

'--
'EXECUTE_TIMER_ENABLED = True


On Error GoTo ErrorTrap

'STOPS THE ENTRY IN LOG 'SCROLL IF ABOVE 0 OR LOWER > TIME
'DON'T ACT ON WHAT IS COPIED FROM CLIPBOARD TEXT BOX
'RUN IF TIME > NOW -- OR = MY FORM -- AND ALSO ONLY IF PICTURE IMAGE = FALSE

'CHANGE TO FIRST IS AND


'----------------------------------------
'COMPARE TO CALL PROGRAM WINMERGE COMPARE
'----------------------------------------
Call COMPARE_FOR_EXE


If RESET_RRR = True Then
    RRR = 0: RESET_RRR = False
End If

If ENTER_LARGE_IN_LOGGER = True Then
    RRR = 0
End If

If Me.WindowState = vbMinimized Then
    RRR = 0
End If

If Mnu_Run_Status.Caption = "1st Run" Then
    Mnu_Run_Status.Visible = False
    Mnu_Clip_Status.Visible = False
End If


'-------------------------
'-------------------------
Call GetFormat_And_Display
'-------------------------
'-------------------------

If Clipboard.GetFormat(vbCFText) = True Then
    Call CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY(False)
End If



'WHEN GRAB TEXT FROM OWN WINDOW ClipBoard Logger
'TREAT IT AS A SAME COPY AND NOT WANTED AND SOUND EFFECT AS SAME #2
'RRR = + TIME WHEN SET
'THIS DONT SEEM TO ENTER AND RUN AND IS CALLED LATER
'YES WHEN WE DON'T WANT IT IT IS RUN

'If (GetForegroundWindow = Me.hwnd) = True Then Stop
'If RRR > Now Then Stop
On Error Resume Next
    CLIP_TEXT_TRUE = Clipboard.GetFormat(vbCFText) = True
If Err.Number > 0 Then MsgBox "ERROR 1ST READ CLIPBOARD IN MAIN TIMER1 : RESUME NEXT - FAULT OPPERATION"

On Error GoTo ErrorTrap

If (RRR > Now And GetForegroundWindow = Me.hwnd) = True And CLIP_TEXT_TRUE Then
    
    OO = Clipboard.GetFormat(vbCFText)
    If OO = True Then
        AD$ = Clipboard.GetText
        AD_DATE = Now
        Call Menu_clipboard_size(Len(AD$))
        

        
        
        '-----------------------------
        'MAYBE WE DON'T WANT THIS HERE
        'OR GET TWO CHUNK WHEN LOAD PROGRAM NEXT TIME
        'USE LATER WHEN WANT
        '-------------------
        TIMER_OutClipChunck_Txt.Enabled = True
        
        If FirstRun = True Then
            'FirstRun = False
            'Mnu_LAST_CLIP_TIME.Caption = Mnu_LAST_CLIP_TIME.Caption '+ " First Run"
            Mnu_Run_Status.Caption = "1st Run"
        End If

    End If
    
    RRR = 0
    
'-----------------------------------------------------
    
    '-------------------------------------------------------------
    'NOT HERE WE WILL IF
    'DUPE BUT NOT AS TAKE TEXT OFF SCREEN FORM
    'MAYBE
    'BUT ADD
    'DO THE MAKE ALL CAPS RUN
    '------------------------
    '-------------------------------------------------------------
    If Mnu_SoundOn.Checked = True Then
        MMControl2.Command = "prev"
        MMControl2.Command = "Play"
    End If

'-----------------------------------------------------
    
    VAR_INFO_LOGG_CLIP = "Clip in Own TextBox"
    TEXT_TO_ENTER_IN_LOGGER = False
    MNU_ERROR_INFO.Caption = VAR_INFO_LOGG_CLIP
    'MNU_ERROR_INFO.Visible = True
    
    Exit Sub
    
    If MNU_LAST_GRAB_ALL_CAPS.Checked = True Then
        Call LAST_GRAB_ALL_CAPS_002_Click
    
    End If
    'Exit Sub

    
'-----------------------------------------------------

End If


OO = False
If OO = False Then OO = Clipboard.GetFormat(vbCFText)
If OO = False Then OO = Clipboard.GetFormat(vbCFBitmap)

If OO = False Then
'-----------------------------------------------------
        '----------------------------------------------------
        'CHANGE TO NORM #1 SOUND WHEN OTHER TYPE OF CLIPBOARD
        'BUT ALREADY GOT THAT LATER
        '----------------------------------------------------
        If Mnu_SoundOn.Checked = True Then
            'MMControl2.Command = "prev"
            'MMControl2.Command = "Play"
            
            'MMControl2.Command = "prev"
            'MMControl2.Command = "Play"
        
        End If
'-----------------------------------------------------
    TEXT_TO_ENTER_IN_LOGGER = False
    'NOT A TEXT OR PICTURE
    
    'Exit Sub
'-----------------------------------------------------
End If




'If OO = False Then OO = Clipboard.GetFormat(vbCFDIB)
'If OO = False Then OO = Clipboard.GetFormat(vbCFEMetafile)
'If OO = False Then OO = Clipboard.GetFormat(vbCFFiles)
'If OO = False Then OO = Clipboard.GetFormat(vbCFLink)
'If OO = False Then OO = Clipboard.GetFormat(vbCFMetafile)
'If OO = False Then OO = Clipboard.GetFormat(vbCFPalette)
'If OO = False Then OO = Clipboard.GetFormat(vbCFRTF)
'If OO = False Then OO = Clipboard.GetFormat(-15694) 'Pasted Objects on form designer
'
'If OO = False Then
'    For R = -30000 To 32000
'        If OO = False Then
'            OO = Clipboard.GetFormat(R)
'        Else
'            Exit For
'        End If
'    Next
'End If

'If OO = False Then
'    If TRemGG$ <> "" Then
'        'ctlClipboard1_ClipboardChanged
'        EXECUTE_TIMER_ENABLED = False
'
'        Clipboard.Clear
'        Clipboard.SetText TRemGG$
'
'        EXECUTE_TIMER_ENABLED = True
'
'        Exit Sub
'    End If
'End If

'VBICON = False


'VBICON CHECKING
If Clipboard.GetFormat(vbCFBitmap) = True Then
    
    Picture1.Picture = Clipboard.GetData(vbCFBitmap)
    
    Xx1 = Picture1.Picture.Width
    yy1 = Picture1.Picture.Height
    'hh1 = Picture1.Picture.hPal
    hh2 = Picture1.Picture.Type
    
    'If Xx1 = 423 And yy1 = 423 And HH1 = 0 And hh2 = 1 Then 'Seems To Have Changed
    'If Xx1 = 847 And yy1 = 847 And HH1 = 0 And hh2 = 1 Then
    
    If Xx1 > 0 And yy1 > 0 And hh2 = 1 Then
        If Xx1 < 900 And yy1 < 900 Then
            
            'GOT TO SAVE IT -- IF SMALL
            SavePicture Picture1.Picture, App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBIcon.bmp"
            
            'OPEN IT AGAIN - TO GET DATA
            FR1 = FreeFile
            Open App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBIcon.bmp" For Binary As #FR1
                Pic2$ = Space$(LOF(FR1))
                Get #FR1, 1, Pic2$
                XLEN = Len(Pic2$)
            Close #1
            
            'IF IT IS NOT THE SAME AS OUR RECORD
            'AND A VBICON PICTURE
            
            'SOMETIME THE VBICON CHANGE
            'BUT MOST THE TIME ONLY REQUIRE CHECK AGAINST ONE OR MAYBE LATER JOIN TWO OR THREE TO COMPARE
            
            'XLEN ---- 4000 TO 7000 FOR WIN 10
            If Pic1$ <> Pic2$ And Mid$(Pic2$, 1, 3) = "BM6" And XLEN < 7000 Then 'And XLEN = 3126 Then
                If Dir(App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBIcon4.bmp") <> "" Then
                    Kill App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBIcon4.bmp"
                End If
                
                FR1 = FreeFile
                Open App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBIcon4.bmp" For Binary As #FR1
                    Pic1$ = Pic2$
                    Put #FR1, 1, Pic1$
                Close #FR1
                
                VBICON_VAR = True
                
                VB_ICON_TEXT_VAR = "VB IDE ICON CHANGED PROPERTY"
                
                MSGBOX2 = "INFO - " + VB_ICON_TEXT_VAR
                FRM_MSGBOX.Timer1.Enabled = True
                
                VB_ICON_CHANGED = True
                VAR_INFO_LOGG_CLIP = VB_ICON_TEXT_VAR
            End If
            
            '--------------------
            'AND TO CHECK IF WANT TEXT PUT BACK WHEN VBICON TRUE
            '---------------------------------------------------
            If VBICON_VAR = True Then
                
                Do
                    If FindWindow("VBSplash", vbNullString) = 0 Then Exit Do
                    'If FindWindow(vbNullString, "New Project") > 0 Then Exit Do
                    Sleep 100
                Loop Until 1 = 2
                
                '--------------------------------------------
                'IT IS A VBICON AND WE WANT OUR TEXT PUT BACK
                '--------------------------------------------
                If AD$ <> "" Then
                    
                    'ctlClipboard1_ClipboardChanged
                    EXECUTE_TIMER_ENABLED = False
                    Clipboard.Clear
                    Clipboard.SetText AD$
                    EXECUTE_TIMER_ENABLED = True
                    
                    
                    Call Menu_clipboard_size(Len(AD$))
                    If FirstRun = True Then
                        Mnu_Run_Status = "1st Run"
                    End If

                    'NOT A FUSS ABOUT SOUND IF VBICON AND RESTORE TEXT
                    If Mnu_SoundOn.Checked = True Then
            '            MMControl1.Command = "prev"
            '            MMControl1.Command = "Play"
                    End If

'-----------------------------------------------------
                    
                    VAR_INFO_LOGG_CLIP = "VB_ICON HAPPEN AND OUR TEXT WAS PUT BACK"
                    If VB_ICON_CHANGED = True Then
                        VAR_INFO_LOGG_CLIP = "VB_ICON HAPPEN AND IT CHANGED - AND OUR TEXT WAS PUT BACK"
                    End If
                    
                    TEXT_TO_ENTER_IN_LOGGER = False
                    'VBICON CHECKING
                    'Exit Sub
'-----------------------------------------------------
                End If
            End If
        End If
    End If
End If



'----
'Call GetFormat_And_Display


'GET SIZE INFO
If Clipboard.GetFormat(vbCFText) = True Then
    GG$ = Clipboard.GetText
    Call Menu_clipboard_size(Len(GG$))
End If

'GET FIRST RUN STATUS
If Clipboard.GetFormat(vbCFText) = True Then
    If FirstRun = True Then
        Mnu_Run_Status.Caption = "1st Run"
    End If
    FirstRun = False
End If

'WHAT DO IF NOTHING
If Clipboard.GetFormat(vbCFText) = True Then
    '-------------------------------------------------------
    'THE GG$ GET_CLIPBOARD WAS DONE BEFORE IN THIS PROCEDURE
    '-------------------------------------------------------
    If GG$ = "" Then
        '-----------------------------------------------------
        VAR_INFO_LOGG_CLIP = "CLIPBOARD SIZE = NOTHING"
        TEXT_TO_ENTER_IN_LOGGER = False
        'Exit Sub
        '-----------------------------------------------------
    End If
End If

'-------------------------------------------
'PICTURE IMAGE IS A DUPE  - CHECK TRUE FALSE
'-------------------------------------------
If Clipboard.GetFormat(vbCFBitmap) = True Then
    'DOUBT IF HERE VER RUNS
'    If VBICON = True Then
'        'Exit Sub
'    End If

    If FirstRun = True Then
        FirstRun = False

        Mnu_Run_Status.Caption = "1st Run"
    
    End If

    '----------------------
    If DUPE_IMAGE_AT_LOAD_FORM = True Then
        DUPE_IMAGE_AT_LOAD_FORM = False
    
        '-----------------------------------------------------
        'GRAB PICTURE SOUND IF SAME AT FIRST RUN
    
        'WHY THESE
        '---------
        If Mnu_SoundOn.Checked = True Then
            MMControl2.Command = "prev"
            MMControl2.Command = "Play"
        End If

        '-----------------------------------------------------
        VAR_INFO_LOGG_CLIP = "CLIPBOARD PICTURE IMAGE - DUPE"
        TEXT_TO_ENTER_IN_LOGGER = False
        'Exit Sub
        '-----------------------------------------------------
        DUPE_IMAGE_AT_FORM_LOAD_VAR2 = True
    End If
End If
    
    
If Clipboard.GetFormat(vbCFBitmap) = True Then
    If DUPE_IMAGE_AT_FORM_LOAD_VAR2 = False Then
            
        '-------------------------
        '-------------------------
        '-------------------------
        'HERE GET THE PICTURE INFO
        '-------------------------
        '-------------------------
        '-------------------------
        Call PictureClip
        
        VAR_INFO_LOGG_CLIP = "CLIPBOARD PICTURE IMAGE"
        TEXT_TO_ENTER_IN_LOGGER = False
        
    '-----------------------------------------------------
        'Exit Sub
    '-----------------------------------------------------
    End If
End If

'---------------------------------------------
'ENTER LARGE TEXT IN LOGGER
'---------------------------------------------
XI1 = TEXT_TO_ENTER_IN_LOGGER
XI2 = Clipboard.GetFormat(vbCFText)
If XI1 = True And XI2 = True Then

    '-------------------------------------------
    'SET EARLIER BECUASE OTHER EVENT WANT TO SEE
    '-------------------------------------------
    LimitClipSize = LimitClipSize
    '-------------------------------------------
    
    MNU_ENTER_LARGE.Caption = "ENTER LARGE TEXT IN LOGGER OR REPEAT ARE BLOCKED - Hardcoded ""BluetoothLogView"" Is Allow -- The Limit" + Str(Int(LimitClipSize / 1000)) + " K" + " and Current =" + Str(Len(GG$))
    If Len(GG$) > LimitClipSize Then
        MNU_ENTER_LARGE.Caption = MNU_ENTER_LARGE.Caption + " = Answer Not Okay"
        MNU_CB_SIZE_BYTE_OVERSIZE.Visible = True
        MNU_CB_SIZE_BYTE_OVERSIZE.Caption = "CLIP OVERSIZE LIMIT -" + Str(Val(VarSize)) + ">" + Trim(Str(Int(LimitClipSize / 1000))) + " K"
    Else
        MNU_ENTER_LARGE.Caption = MNU_ENTER_LARGE.Caption + " = Answer Okay"
    End If
    
    If ENTER_LARGE_IN_LOGGER = True Then
        If Len(GG$) > LimitClipSize Then
            VAR_INFO_LOGG_CLIP = "LARGE TEXT IN LOGGER - ACCEPTED IS CUSTOMER USER REQUESTED"
        End If
    End If
    
    If FindWindow(vbNullString, "BluetoothLogView") = GetForegroundWindow Then
        'If Len(GG$) > LimitClipSize = True Then
            VAR_INFO_LOGG_CLIP = "LARGE TEXT IN LOGGER - ACCEPTED IS BluetoothLogView"
        'End If
        ENTER_LARGE_IN_LOGGER = True
    End If
    
    If ENTER_LARGE_IN_LOGGER = False Then
        If Len(GG$) > LimitClipSize Then
            VAR_INFO_LOGG_CLIP = "BLOCKED LARGE TEXT IN LOGGER"
            GG$ = "": GGSize = True: Exit Sub
        Else
            GGSize = False
        End If
        'CLEAR THE GG$ BECAUSE IF IT'S BIG WON'T MATTER TO
        'RELOAD AGAIN AND CHECK IF IT'S DUPE
    End If

    '----------------------------------------------------
    'RESET THIS IT WAS CALLED BY THE MENU PULL DOWN CLICK
    '----------------------------------------------------
    ENTER_LARGE_IN_LOGGER = False

End If



'--------------------------------------------------------
'TEXT IS A DUPE TRUE FALSE
'--------------------------------------------------------
'DUPE AS LAST TIME THAT HAS BEEN LOGGED SEE FURTHER AHEAD
'--------------------------------------------------------
If Clipboard.GetFormat(vbCFText) = True Then
    XI2 = Clipboard.GetFormat(vbCFText)
    If AD$ = GG$ And XI2 = True Then
    '-----------------------------------------------------
        If AD$ <> "" Then
            '------------------
            'PLAY TUNE FOR DUPE
            If Mnu_SoundOn.Checked = True Then
                '----------------------------------------
                'IN WIN 10 IS IF SOUND NOT ABLE TO PLAY HERE
                'AND THEN CAUSE CRASH HOLD FREEZE
                '----------------------------------------
                
                If MMControl2.fileName = "" Then 'vPathSOUND2 + vFileSOUND2
                    Call RESET_SETUP_SOUND_FILE(VARSOUND)
                End If

                
                MMControl2.Command = "prev"
                MMControl2.Command = "Play"
            End If
        End If
    
    '-----------------------------------------------------
        VAR_INFO_LOGG_CLIP = "Dupe Text__Blocked"
        TEXT_TO_ENTER_IN_LOGGER = False
    '    Exit Sub
    '-----------------------------------------------------
    End If
End If
'----------------------------------
'----------------------------------
'----------------------------------





'----------------------------------
'----------------------------------
'----------------------------------
'BEGIN LOGG TEXT IN CLIPBOAD_LOGGER
'----------------------------------

'---------------------------------------------
'IF NOT WANT TEXT TO ENTER IN CLIPBOARD LOGGER
'AND INFO PRESENT AND THEN IT MUST BE
'SOMETHING ELSE
'---------------------------------------------

'------------------------------------------------------------------------------
'BETTER PLACE FOR THIS MORE NEAR FRONT AND OR ALSO END -- CHECK EXIT SUB - LINE
'MORE TO USE THIS LATER AND MORE USE IS IN FRONT
'TIDY WITH SUB ROUTINE
'------------------------------------------------------------------------------
'MNU_ERROR_INFO.Visible = False
If VAR_INFO_LOGG_CLIP <> "" Then
    'MNU_ERROR_INFO.Caption = Format(Now, "DDD HH:MM:SS") + " -- " + VAR_INFO_LOGG_CLIP
    '------------------------------------------
    MNU_ERROR_INFO.Caption = VAR_INFO_LOGG_CLIP
    'MNU_ERROR_INFO.Visible = True
    '------------------------------------------
End If

PICTURE_IMG = False

If TEXT_TO_ENTER_IN_LOGGER = False Then
    If INFO_TO_ENTER_IN_LOGGER = True Then
    
        '------------------------------------
        'ALL THE POSIABLE INFO TAGGS 02 OF 03
        '------------------------------------
        
        '------------------------------------------------
        'COULD BE MORE INFO TOLD BLUETOOTHLOGGER ENTERED MORE
        '------------------------------------------------
        
        If VAR_INFO_LOGG_CLIP = "CLIPBOARD SIZE = NOTHING" Then
            INFO_TO_ENTER_IN_LOGGER = False
        End If
        
        If VB_ICON_TEXT_VAR = "VB IDE ICON CHANGED PROPERTY" Then
            INFO_TO_ENTER_IN_LOGGER = False
        End If
        If VAR_INFO_LOGG_CLIP = "VB_ICON HAPPEN AND OUR TEXT WAS PUT BACK" Then
        End If
        
        If VAR_INFO_LOGG_CLIP = "Clipboard Taken From Within TextBox" Then
            INFO_TO_ENTER_IN_LOGGER = False
        End If
    
        If VAR_INFO_LOGG_CLIP = "Dupe Text__Blocked" Then
            INFO_TO_ENTER_IN_LOGGER = False
        End If
        
        If VAR_INFO_LOGG_CLIP = "VB_ICON HAPPEN AND OUR TEXT WAS PUT BACK" Then
            INFO_TO_ENTER_IN_LOGGER = False
        End If
        If VAR_INFO_LOGG_CLIP = "VB_ICON HAPPEN AND IT CHANGED - AND OUR TEXT WAS PUT BACK" Then
            INFO_TO_ENTER_IN_LOGGER = False
        End If
        
        '------------------------------------
        'ALL THE OTHER POSIABLE INFO 02 OF 03
        '------------------------------------
        If VAR_INFO_LOGG_CLIP = "BLOCKED LARGE TEXT IN LOGGER" Then
        End If
        
        If VAR_INFO_LOGG_CLIP = "CLIPBOARD PICTURE IMAGE - DUPE" Then
        End If
        
        If VAR_INFO_LOGG_CLIP = "LARGE TEXT IN LOGGER - ACCEPTED OKAY IS CUSTOMER REQUESTED" Then
        End If
    
        If VAR_INFO_LOGG_CLIP = "LARGE TEXT IN LOGGER - ACCEPTED OKAY IS BluetoothLogView" Then
        End If
    
        '------------------------------------
        'ALL THE OTHER POSIABLE INFO 03 OF 03
        '------------------------------------
    
    End If
    
    '-------------------------------------------------
    'PART 2 CHECK FOR MORE INFO THAT NOT ALREADY TAKEN
    'BY TEXT AND PICTURE
    '-------------------------------------------------
    
    
    
    XIO = False
    If Clipboard.GetFormat(vbCFBitmap) = True Then XIO = True: PICTURE_IMG = True
    If Clipboard.GetFormat(vbCFText) = True Then XIO = True

    If XIO = False Then
        VAR_INFO_LOGG_CLIP = "CLIP INFO --- " + ClipFormatDescription
        INFO_TO_ENTER_IN_LOGGER = True
    End If


    If VAR_INFO_LOGG_CLIP = "CLIPBOARD PICTURE" Then
        INFO_TO_ENTER_IN_LOGGER = True
        
    End If


    'this is correct when clipboard is a empty
    '-----------------------------------------------
    If INFO_TO_ENTER_IN_LOGGER = False Then Exit Sub
    '-----------------------------------------------

End If






'-------------------------------
'NORMAL TEXT ENTRY -- TRUE FALSE
'-----------------
If Clipboard.GetFormat(vbCFText) = True Then

    If GG$ = "" Then
        'PICTURE_OR_TEXT = True
        'INFO_TO_ENTER_IN_LOGGER = True
        Exit Sub
    
    Else
    
        INFO_TO_ENTER_IN_LOGGER = False
        CLIPBOARD_INFO_ENTER_TEXTBOX = GG$
        PICTURE_OR_TEXT = True
    End If
End If


If ClipFormatDescription = "Text And Len(0) Empty Clipboard" Then Stop


'wndclass_desked_gsk


'--------
'GOING IN
'--------
'--------------------------------
'GET SOME VALUES - REQUIRED
'--------------------------------

On Error GoTo 0
Ash$ = GetActiveWindowTitle(True)

GetForegroundWindow_XXRT = GetForegroundWindow
XXR7 = GetForegroundWindow_XXRT
Do
    XXR5 = GetForegroundWindow_XXRT
    GetForegroundWindow_XXRT = GetParent(GetForegroundWindow_XXRT)
Loop Until GetForegroundWindow_XXRT = 0

If XXR5 <> XXR7 Then
    GetForegroundWindow_XXRT = XXR5
    ASH2$ = ""
    ASH2$ = GetWindowTitle(GetForegroundWindow_XXRT)
Else
    GetForegroundWindow_XXRT = 0
End If

CountR = CountR + 1

frmSender.URLLogging

Dim UrlWork
'VAR_VBCRLF = vbLf
VAR_VBCRLF = vbCrLf

If URL <> "" Then
    UrlWork = UrlWork + "In Internet Explorer -- " + VAR_VBCRLF
    UrlWork = UrlWork + "WinTitle = " + WinTitle + VAR_VBCRLF
    UrlWork = UrlWork + "URL = " + URL + VAR_VBCRLF
End If

RR$ = "---------------------" + VAR_VBCRLF
RR$ = RR$ + "Count = " + Format$(CountR, "000") + " -- "
RR$ = RR$ + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + VAR_VBCRLF
RR$ = RR$ + "---------------------" + VAR_VBCRLF

If UrlWork <> "" Then
    RR$ = RR$ + UrlWork
    RR$ = RR$ + "---------------------" + VAR_VBCRLF
End If

If ASH2$ <> "" And ASH2$ <> Ash$ Then
    RR$ = RR$ + "Form Parent = " + VAR_VBCRLF + ASH2$ + VAR_VBCRLF
    RR$ = RR$ + "---------------------" + VAR_VBCRLF
End If

If ASH2$ <> "" And ASH2$ <> Ash$ Then
    RR$ = RR$ + "Form Parent = " + VAR_VBCRLF + ASH2$ + VAR_VBCRLF
    RR$ = RR$ + "---------------------" + VAR_VBCRLF
End If
    
RR$ = RR$ + "Form FindWindow ---" + VAR_VBCRLF + Ash$ + VAR_VBCRLF

'--------------------------------------------------
'THE VAR OF CLIPBOARD IS GG$
'--------------------------------------------------
'BUT WE WANT OTHER CIRCUMSTANCE OF INFO LOGGED ALSO
'--------------------------------------------------------
'AND WE WANT OTHER CIRCUMSTANCE OF INFO OF PICTURE LOGGED
'---------------------------------------------------------------
'AND WE WANT OTHER CIRCUMSTANCE OF INFO OF DRAG DROP INFO LOGGED
'---------------------------------------------------------------

'----------------------------------------------------

If INFO_TO_ENTER_IN_LOGGER = True Then
    
    CLIPBOARD_INFO_ENTER_TEXTBOX = VAR_INFO_LOGG_CLIP
    
    'IF A PICTRUE WE WANT THIS
    If VAR_INFO_LOGG_CLIP = "CLIPBOARD PICTURE" Then
        If PICTURE_INFO_TEXT <> "" Then
            CLIPBOARD_INFO_ENTER_TEXTBOX = PICTURE_INFO_TEXT
            PICTURE_OR_TEXT = True
        End If
    End If
End If
'----------------------------------------------------




'--------------------
'DECIDE ABOUT GO IN FOR PICTURE
'IF GOOD PICTURE WILL BE SET INTO VAR TO USE TEXT BOX IN THE NEXT LINE IN THE IF
'------------------------------
'SOMEHOW FIGURE OUT IN A MINUTE
'------------------------------


If PICTURE_OR_TEXT = False Then
    CLIPBOARD_INFO_ENTER_TEXTBOX = VAR_INFO_LOGG_CLIP
End If


If Clipboard.GetFormat(vbCFText) = True Then
    If GG$ = "" Then PICTURE_OR_TEXT = False
End If


If PICTURE_OR_TEXT = False Then
    If INFO_TO_ENTER_IN_LOGGER = False Then
        Exit Sub
    End If
End If

'ONE WORKAROUND - NOT WANT ANY VB ENTRY ICON
'-------------------------------------------

'CurProcHwnd= LAST SET = Sub PictureClip()
'CurProcHwnd = GetForegroundWindow

'If GetWindowClass(CurProcHwnd) = "wndclass_desked_gsk" Then

If VB_ICON_CHANGED = True Then Exit Sub



'------------------------
'GOING IN FINAL --------- FOR TEXT OR PICTURE IF GOOD
'------------------------

RR$ = RR$ + "---------------------" + VAR_VBCRLF
RR$ = RR$ + CLIPBOARD_INFO_ENTER_TEXTBOX + VAR_VBCRLF
RR$ = RR$ + "---------------------" + VAR_VBCRLF

'--------------
'AND THEN CLEAR
'--------------
CLIPBOARD_INFO_ENTER_TEXTBOX = ""

TRemGG$ = GG$

OFH = Text1.SelStart

Err.Clear
On Error Resume Next
Text1.Text = Text1.Text + RR$

If Err.Number > 0 Then
    Me.WindowState = vbNormal
    DoEvents
    '-----------------------------------------------------
    MSGBOX2 = "ERROR Not Adding to CLip Board TEXT1.TEXT RTB Textbox" + vbCrLf + Err.Description
    '-----------------------------------------------------
    FRM_MSGBOX.Timer1.Enabled = True
    '-----------------------------------------------------
    Exit Sub
    '-----------------------------------------------------
End If

'ADD CLIPBOARD ON UNIQUE SOUND PLAYED OTR DUPE ELSE WHERE
'AND PICTURE IS HANDLE ELSE WHERE
'AND HOPE TEST DUPE FOR THAT WORK
'BUT TEXT ON PICTURE NEVER BE DUPE

If PICTURE_IMG = False Then
    If Mnu_SoundOn.Checked = True Then
        MMControl1.Command = "prev"
        MMControl1.Command = "Play"
    End If
End If

On Error GoTo 0

If MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = True Then
    
    'TEXTBOX1_SELECTION_CHANGE = Text1.SelStart
    
    Text1.Enabled = False
    Text1.SelStart = 1
    Text1.Enabled = True
    Text1.SelStart = TEXTBOX1_SELECTION_CHANGE

End If

If MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = False Then
    
    'MAIN TEXT BOX IS TEXT_BOX NOT RTB
    'THINK
    'RTB IS CALLED TEXT1
    
    Text1.SelStart = Len(Text1.Text)
    
    'WT-HECK -- IN ANOTHER MODE THE TEXT SCROLL WONT UPDATE STAY AT LAST
    'If Form1.Width <> 2400 Then
    '    Text1.SelStart = OFH
    'End If
    
    '----------------------------------------------------
    'DONT MOVE CURSOR IN TEXBOX IF TIMER SET AT NEW PASTE
    '----------------------------------------------------
    If XXMOUSEMOVE < Now Then
        Text1.SelStart = Len(Text1.Text)
    End If

End If

On Error GoTo ErrorTrap

'---------------------------
'OUTPUT
'---------------------------
FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#Count.Txt"
'DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)

FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Output As #FR1
    Print #FR1, CountR
Close #FR1

'---------------------------
'APPEND
'---------------------------
FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA + "\ClipBoard-" + Format$(start, "YYYY-MM-DD") + ".Txt"
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, RR$;
Close #FR1

FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\#ClipBoard.Txt"
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, RR$;
Close #FR1

FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total.txt"
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, RR$;
Close #FR1

FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Total--TRIM.txt"
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, RR$;
Close #FR1
    
FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\00_ClipBoard_Tot_Strip.txt"
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, GG$
Close #FR1

FILENAME_IN_USE_CHECK = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\00_ClipBoard_Total--TRIM-" + GetComputerName + "-" + GetUserName + ".txt"
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, RR$;
Close #FR1




AD$ = GG$
AD_DATE = Now

'---------------------------------------
'IF MENU SET TO UPPER CASE ALL CLIPBOARD
'---------------------------------------
Call LAST_GRAB_ALL_CAPS_002_Click
'---------------------------------------

'-----------------------------
Exit Sub
'-----------------------------

'-----------------------------
ErrorTrap:
    'SLEEP 10
    DoEvents
    Resume Next
Exit Sub
'-----------------------------

End Sub

Public Function GetActiveWindowTitle(ByVal ReturnParent As Boolean) As String
   Dim I As Long
   Dim j As Long
   I = GetForegroundWindow
   ReturnParent = False
   If ReturnParent Then
      Do While I <> 0
         j = I
         I = GetParent(I)
      Loop
   I = j
   End If
   GetActiveWindowTitle = GetWindowTitle(I)
End Function

'Function GetWindowTitle(ByVal hwnd As Long) As String
'   Dim L As Long
'   Dim S As String
'   L = (hwnd)
'   S = Space(GetWindowTextLength(L))
'   GetWindowText hwnd, S, L + 1
'   'Prob here sudden start go wrong but look i is handle so how can text lenght be so
'   GetWindowTitle = S 'Left$(s, l)
'End Function

Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hwnd)
   S = Space(l + 1)
   GetWindowText hwnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function

Sub SET_FOLDER_CLIPBOARD_LOGGER()

    QQ2$ = "Hot-Key-App-Shots\"
    ART1$ = "D:\0 00 Art Loggers\"
    ART2$ = ART1$ + "# APP AND SCREEN\" + GetComputerName + "\"
    'QQ4$ = "Hot-Key-Screen-Shots\"
    
    FF$ = "HotKey-Capture_" + Format$(Now, "YYYY-MM-DD-DDD") + "-App"
    On Error Resume Next
    FF_22 = "HOT_KEY_SCREEN_SHOT\"
    FF_24 = "AUTO_FORM_SCREEN_SHOT\"
    
    PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY = ART2$ + FF_22 + QQ2$
    PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO___ = ART2$ + FF_24 + QQ2$
    
    If (Not DirExist(PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY)) Then
        Err.Clear
        MkDirNested PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY
        If Err.Number <> 0 Then
            MSGBOX2 = "ERROR Clipboard Images Not Getting Saved" + vbCrLf + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf + PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY + "ERR.DESCRIPTION=" + Err.Description
            FRM_MSGBOX.Timer1.Enabled = True
        End If
    End If
    If (Not DirExist(PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO___)) Then
        Err.Clear
        MkDirNested PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO___
        If Err.Number <> 0 Then
            MSGBOX2 = "ERROR Clipboard Images Not Getting Saved" + vbCrLf + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf + PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO___ + "ERR.DESCRIPTION=" + Err.Description
            FRM_MSGBOX.Timer1.Enabled = True
        End If
    End If

    PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP = App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBDataNoTArchive"
    If (Not DirExist(PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP)) Then
        Err.Clear
        MkDirNested PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP
        If Err.Number <> 0 Then
            MSGBOX2 = "ERROR Clipboard Images Not Getting Saved" + vbCrLf + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf + PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP + "ERR.DESCRIPTION=" + Err.Description
            FRM_MSGBOX.Timer1.Enabled = True
        End If
    End If
    
    PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "\Day-Data"
    If (Not DirExist(PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA)) Then
        Err.Clear
        MkDirNested PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA
        If Err.Number <> 0 Then
            MSGBOX2 = "ERROR Clipboard Images Not Getting Saved" + vbCrLf + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf + PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DAY_DATA + "ERR.DESCRIPTION=" + Err.Description
            FRM_MSGBOX.Timer1.Enabled = True
        End If
    End If
    
    PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA
    If (Not DirExist(PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA)) Then
        Err.Clear
        MkDirNested PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA
        If Err.Number <> 0 Then
            MSGBOX2 = "ERROR Clipboard Images Not Getting Saved" + vbCrLf + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf + PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP_DATA + "ERR.DESCRIPTION=" + Err.Description
            FRM_MSGBOX.Timer1.Enabled = True
        End If
    End If

End Sub


Sub PictureClip()

Dim IP, TimeSet2
On Error Resume Next

'--------------------------------------------------------------------
If Mnu_SoundOn.Checked = True Then
    MMControl1.Command = "prev"
    MMControl1.Command = "Play"
End If
'--------------------------------------------------------------------

Call SET_FOLDER_CLIPBOARD_LOGGER

'--------------------------------------------------------------------
Filexxx$ = GetFileFromHwnd(CurProcHwnd)
'TOO DIFFERCULT TO COMPARE PICTURE TO LAST ONE SAVED AT PROGRAM START
'MAYBE AT FORM LOAD
'YES DONE - AND BOTH
CurProcHwnd = GetForegroundWindow
'CurProcHwnd= LAST SET = Sub PictureClip()
'CurProcHwnd = GetForegroundWindow
'GetWindowClass(CurProcHwnd)
hwnd9 = GetWindowRect(CurProcHwnd, Rect1)
'--------------------------------------------------------------------

'--------------------------------------------------------------------
Picture3.Picture = Clipboard.GetData(vbCFBitmap)
U1 = Picture3.Width
U2 = Picture3.Height
If U1 > Screen.Width And U2 > Screen.Height Then QQ3$ = "Screen" Else QQ3$ = "App"
'--------------------------------------------------------------------

'--------------------------------------------------------------------
'<---------------------------------
'CHECK DUPE EXIT IS TRUE
'IN THE APP FOLDER
'-----------------
FILENAME1 = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP + "\COMPARE DUPE.txt"
If Dir(FILENAME1) <> "" Then Kill FILENAME1

SavePicture Picture3.Picture, FILENAME1
'<---------------------------------
'CHECK DUPE EXIT IS TRUE
'--------------------------------------------------------------------
    
'--------------------------------------------------------------------
'--------------------------------------------------------------------
PICX2$ = PICXX$
FR1 = FreeFile
Open FILENAME1 For Binary As #FR1
    PICXX$ = Space$(LOF(FR1))
    Get #FR1, 1, PICXX$
    XLEN = LOF(1)
Close #1
'--------------------------------------------------------------------
'LESS PROBLEM WITH INUSE DELAY IF DELETE
'Kill FILENAME1
'THINK WE WANT TWO - ACOPY FOR DUPE CHECK COMPARE AT START OF PROGRAM

'SET TO DUPE BUT NOT FOR ANY REASON TO DO LATER

'--------------------------------
'NOT CHECKING DUPE PROPER PROBLEM
'--------------------------------
If PICX2$ = PICXX$ Then PICX2$ = "DUPE": Exit Sub Else PIX2$ = ""
'--------------------------------------------------------------------
'--------------------------------------------------------------------

FileInUseDelay PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP + "\HotKey-Shot % Pic Data.txt"
FR1 = FreeFile
Open PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP + "\HotKey-Shot % Pic Data.txt" For Append As #FR1
    Print #FR1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p")
Close #FR1

'If Wo / (Ri + Wo) * 100 >= 1 Then
'TS1 = PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY + "\HotKey-Capture- " + TimeSet2 + "-" + QQ3$ + ".jpg"

TimeSet2 = Format$(Now, "YYYY-MM-DD HH-MM-SS DDD")
TS1 = PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY + "\HotKey-Capture- " + TimeSet2 + ".jpg"
SavePicture Picture3.Picture, TS1
LAST_CLIPBOARD_FILE = TS1

FileInUseDelay PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP + "\HotKey-Shot % Pic Data Location.txt"
FR1 = FreeFile
Open App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBDataNoTArchive\HotKey-Shot % Pic Data Location.txt" For Append As #FR1
    Print #FR1, LAST_CLIPBOARD_FILE
Close #FR1

Set F = FS.GetFile(TS1)

Call Menu_clipboard_size(F.Size)
Call GetFormat_And_Display

'PATH_TO_CLIPBOARD_AND_FILENAME = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP + "\HotKey-Capture- " + TimeSet2 + "-" + QQ3$ + "-ClipText" + ".txt"
PATH_TO_CLIPBOARD_AND_FILENAME = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP + "\HotKey-Capture- " + TimeSet2 + "-ClipText" + ".txt"

'FileInUseDelay PATH_TO_CLIPBOARD_AND_FILENAME
FR1 = FreeFile
Open PATH_TO_CLIPBOARD_AND_FILENAME For Output As #FR1
    Print #FR1, "---------------------"
    
    If Tag2 = 0 Then Print #FR1, "Screen HotKey Application Shot FileName ="
    If Tag2 = 2 Then Print #FR1, "Screen HotKey Screen Shot FileName ="
    
    Print #FR1, TS1
    Print #FR1, "---------------------"
    Print #FR1, "Application Exe = "; Filexxx$
    Print #FR1, "---------------------"
    Print #FR1, "Application Title = "; GetWindowTitle(CurProcHwnd)
    Print #FR1, "---------------------"
    Print #FR1, "Application Class Title = "; GetWindowClass(CurProcHwnd)
    Print #FR1, "---------------------"
    Print #FR1, "Current Handle Hwnd ="; CurProcHwnd
    Print #FR1, "---------------------"
    Print #FR1, "Resolution -----"
    Print #FR1, "---------------------"
    Print #FR1, "Top ="; Rect1.Top
    Print #FR1, "Bottom ="; Rect1.Bottom
    Print #FR1, "Left ="; Rect1.Left
    Print #FR1, "Right ="; Rect1.Right
    Print #FR1, "---------------------"
    Print #FR1, "Width ="; Rect1.Right - Rect1.Left
    Print #FR1, "Height ="; Rect1.Bottom - Rect1.Top
    Print #FR1, "---------------------"
    Print #FR1, MNU_CB_SIZE_BYTE.Caption
    Print #FR1, "---------------------"
    If PICX2$ = "DUPE" Then
        Print #FR1, "PICTURE IMAGE IS A DUPLICATE"
        Print #FR1, "---------------------"
    Else
        Print #FR1, "PICTURE IMAGE IS NOT DUPLICATE"
        Print #FR1, "---------------------"
    End If
    
    
    GSX = 0
    If GGSize = True Then
        Print #FR1, "Clipboard Text too Bigger Limit to Store Before Image"
        GSX = 1
    End If
    
    If GSX = 0 And ADLEN = Len(AD$) Then
        Print #FR1, "ClipBaord TEXT Same Size as Last One"
    End If
    Print #FR1, "---------------------"
    
    If GSX = 0 And ADLEN <> Len(AD$) Then
        Print #FR1, Trim(Str(Len(AD$))) + " Byte's of Text Here - Held Before Picture"
        Print #FR1, "---------------------"
    End If
    If GSX = 0 And ADLEN <> Len(AD$) Then
        Print #FR1, "NOT ANY TEXT TO BE WRITEN - SAME LENGHT AS ONE BEFORE"
    End If
    If GSX = 0 And ADLEN <> Len(AD$) Then
        Print #FR1, "NOT ANY TEXT TO BE WRITEN - SAME LENGHT AS ONE BEFORE OR TOO BIGG - "
    End If

    STRING_VAR = "--------------------- TEXT ---------------------"
    If GSX = 0 And ADLEN <> Len(AD$) Then
        Print #FR1, STRING_VAR
        Print #FR1, AD$
        GSX = 1
    End If
Close #FR1
    
Dim XPOS1

FR1 = FreeFile
Open PATH_TO_CLIPBOARD_AND_FILENAME For Input As #FR1
    
    '-2 FOR VBCRLF
    PICTURE_INFO_TEXT = Input(LOF(FR1) - 2, FR1)
    XPOS1 = InStr(PICTURE_INFO_TEXT, STRING_VAR)
    If XPOS1 > 0 Then
        PICTURE_INFO_TEXT = Mid(PICTURE_INFO_TEXT, 1, XPOS1)
    End If
    '"---------------------"
    PICTURE_INFO_TEXT = String(40, "-") + vbCrLf + PATH_TO_CLIPBOARD_AND_FILENAME + vbCrLf + String(40, "-") + vbCrLf + PICTURE_INFO_TEXT + vbCrLf + String(40, "-")
    'PICTURE_INFO_TEXT = Replace(PICTURE_INFO_TEXT, vbCrLf, vbLf)
    'MsgBox PICTURE_INFO_TEXT
Close #FR1

' Clear the third picture (not necessary if not visible).
Picture3.Picture = LoadPicture()

'IF SAME LENGHT AD$ AS LAST TIME THEN DON'T STORE TWO VAR -- WITH PICTURE
'REALLY WANT MULTPLE LAST 3 TEXTES SAVED WITH PICTURE
'USE RTB2
ADLEN = Len(AD$)

Exit Sub
Errcode2:
DoEvents
Resume Next

End Sub




Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

Sub FileInUseDelay(FILENAME_IN_USE_CHECK As String)
    
    '--------------------------------------------------
    'REPEAT USE OF code
    'MAKE USE ONE CODE SUB ROUTINE THIS ONE LOOK BETTER
    'IsFileOpenDelay
    'REDO CHANGE UPDATE CODE REQUIRED
    '--------------------------------------------------
    Call IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    
    Exit Sub
    
    '-------------------------------
    'THIS CODE
    'FileInUseDelay
    'SEE DEPANDANT SUB ROUTINE ABOVE
    '-------------------------------
    
    'AND HERE ---- QQ_LIMIT = 40
    
    QQ_LIMIT = 90
    QQ = Now + TimeSerial(0, 0, QQ_LIMIT)
    Do
'        DoEvents
        ii = FileInUse(FILENAME_IN_USE_CHECK)
        If ii = True Then
            Sleep (200)
            DD_DIFF = DateDiff("S", Now, QQ)
            Form1.MNU_FILE_STUCK_IN_USE.Visible = True
            If DD_DIFF <> O_DD_DIFF Then
                Form1.MNU_FILE_STUCK_IN_USE.Caption = "FileInUseDelay --" + Str(DD_DIFF) + " Secs - File Stuck In Use Of Limit --" + Str(QQ_LIMIT) + " -- " + FILENAME_IN_USE_CHECK
            End If
            O_DD_DIFF = DD_DIFF
        End If
    Loop Until ii = False Or QQ < Now
    
    If ii = False Then
        'Form1.MNU_FILE_STUCK_IN_USE.Visible = False
        Form1.MNU_FILE_STUCK_IN_USE.Caption = "Last FileInUseDelay --" + Str(QQ_LIMIT)
    End If

    If ii = True Then
        Form1.MNU_FILE_STUCK_IN_USE.Caption = "Error FileInUseDelay --" + Str(DD_DIFF) + " Secs - File Stuck In Use Of Limit --" + Str(QQ_LIMIT) + " -- " + FILENAME_IN_USE_CHECK
        MSGBOX2 = "ERROR FileInUseDelay Stuck" + vbCrLf + FILENAME_IN_USE_CHECK + vbCrLf + Str(DD_DIFF) + " Sec of" + Str(QQ_LIMIT) + " -- " + FILENAME_IN_USE_CHECK
        FRM_MSGBOX.Timer1.Enabled = True
        If IsIDE = True Then Stop
    End If
End Sub




Public Function GetFileFromHwnd(lngHwnd) As String

'MsgBox getfilefromhwnd(Me.hwnd)

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim x

strFile = String$(256, 0)
x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameExA(hProcess, bla, strFile, Len(strFile))
GetFileFromHwnd = Left(strFile, C)

End Function


Function GetWindowClass(ByVal hwnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hwnd, sText, 255)
    sText = Left$(sText, Ret)
    GetWindowClass = sText
End Function



Private Sub Timer_JOYPAD_Timer()

Exit Sub
'To Do with the JoyPad Scrolling Notepad and Winamp and Outlook
'NOT WINAMP YET
'Call FrmJoy.Timer_XxXJOY
'Exit Sub

End Sub



'----------------       Latest Version Of Save Chks
'#### REMEBER SWITCHES

'Sub zzLoad_Checks()
'Put This In Declarations
'Dim OCheckQ
'Very Nice Code Will Save all your Check Boxes an Menu Checks and Values you can filter
'some out -- works best as I Know

'if you use menu checks you have to set them yourself on clicks
'If Mnu_VB.Checked = True Then Mnu_VB.Checked = False: Exit Sub
'If Mnu_VB.Checked = False Then Mnu_VB.Checked = True

'3 Parts
' ---
'1 Load
'2 Save
'3 Timer ' This Chks for Changes using XOR Hash
'   Best way I know with Timer Hardly CPU Usage Unfriendly

'zzCheckTimer.Enabled = True
'Dont Have Timer Enabled on Form Load

'Call Timer to Run On Form Unload ----- call zzCheckTimer_Timer
'Then you could set timer slow like 10 Secs - I Do

'-------------------------------
Sub zzLoad_Checks()

'zzCheckTimer.Enabled = False

Dim TH(), ATD, FR1
ReDim TH(Me.Controls.Count * 3)

On Error Resume Next
I = 0
FR1 = FreeFile
Open App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\0TextData\ChkSettings.txt" For Input As #FR1
Do
    If Not EOF(FR1) Then Line Input #FR1, vv$
    TH(I) = vv$
    I = I + 1
    If Not EOF(FR1) Then Line Input #FR1, vv$
    TH(I) = Val(vv$)
    I = I + 1
    If Not EOF(FR1) Then Line Input #FR1, vv$
    TH(I) = Val(vv$)
    I = I + 1
Loop Until EOF(1)
Close #1
    
tit = I
For Each Control In Me.Controls
    ATD = 1
    
    ppi = LCase(Control.Name)
    If InStr(ppi, LCase("Combo")) Then ATD = 0
    If InStr(ppi, LCase("Check")) Then ATD = 0
    If InStr(ppi, LCase("Mnu")) Then ATD = 0
    If InStr(ppi, LCase("Menu")) Then ATD = 0
    
    XXD = -1
    For r = 0 To tit
        If Control.Name = TH(r) Then
            XXD = r: Exit For
        End If
    Next
    
    If XXD > 0 And ATD = 0 Then
        On Error Resume Next
        If TH(XXD + 1) <> 0 Then Control.Value = TH(XXD + 1)
        If TH(XXD + 2) <> 0 Then Control.Checked = TH(XXD + 2)
        'Dont Let People Play Around If you Want to Hard Code In Designer To Enable Chk This
        'This Lets Happen Eg If Th(xxd + 2) <> 0
        
        A1 = 0
        If Mid(LCase(Control.Name), 1, 3) <> "mnu" Then
            A1 = Control.Value
        End If
        B1 = 0
        B1 = Control.Checked
        OCheckQ = OCheckQ + Str(A1) + Str(Val(B1))
        On Error GoTo 0
    End If
Next
'zzCheckTimer.Enabled = True
End Sub

'-------------------------------
Sub zzSave_Checks()
Dim A1, B1 As Integer, ATD
On Error Resume Next
'filename=
If FS.FolderExists(App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\0TextData") = False Then
    MkDir App.Path + "\# DATA"
    MkDir App.Path + "\# DATA\" + GetComputerName
    MkDir App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\0TextData"
End If

Open App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\0TextData\ChkSettings.txt" For Output As #1

For Each Control In Me.Controls
    Err.Clear
    A1 = 0
    A2 = 1 '=ERROR
    ATD = 1
    'If Mid(LCase(Control.Name), 1, 3) = "mnu" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 5) = "timer" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 3) = "pic" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "mmco" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "text" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 3) = "dir" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "line" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 5) = "label" Then ATD = 0
    If Mid(UCase(Control.Name), 1, 4) = "MNU_" Then ATD = 2
    If ATD = 1 Then
        A2 = 0
        A1 = Control.Value
        A2 = Err.Number
    End If
    
    Err.Clear
    B1 = 0
    B2 = 1 '=ERROR
    ATD = 1
    If Mid(LCase(Control.Name), 1, 5) = "timer" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 3) = "pic" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "mmco" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "text" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 3) = "dir" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "line" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 5) = "label" Then ATD = 0
    If Mid(UCase(Control.Name), 1, 4) = "COMM" Then ATD = 0
    If ATD = 1 Then
        B2 = 0
        B1 = Control.Checked
        B2 = Err.Number
    End If

    If A2 = 0 Or B2 = 0 Then
        Print #1, Control.Name
        Print #1, A1
        Print #1, B1
    End If
Next
Close #1

End Sub



'-------------------------------
Private Sub zzCheckTimer_Timer()


On Error Resume Next
For Each Control In Me.Controls
    A1 = 0
    B1 = 0
    NOT_FLAG = 0
    NOT_FLAG = 0
    'TOO HARD SOME USE VALUE WHEN THEY DONT USE CHECKED
    'USE ERROR LEVEL WITH PROGRAM OPTIONS ERROR BREAK
    If InStr(UCase(Control.Name), "TIMER") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "DIR") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "PICT") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "TIMER") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "PICT") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "MMCONT") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "COMM") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "LABEL") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "TEXT") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "MNU_") > 0 Then NOT_FLAG = -1
    If NOT_FLAG <= 0 Then
        If NOT_FLAG = -1 Then A1 = 0 Else A1 = Control.Value
        B1 = Control.Checked
    End If
    checkq = checkq + Str(Val(A1)) + Str(Val(B1))
Next

If checkq = OCheckQ Then Exit Sub

OCheckQ = checkq
Call zzSave_Checks

End Sub


Sub SUB_RENAME_JPG_URL(FDS)
'CALLED FROM Timer_WEATHER_URL_Timer
Exit Sub

Dim FDS_END
    
On Error Resume Next
If FS.FileExists(FDS) = True Then
    If FileInUse(FDS) = False Then
        
        Filespec1 = FDS
        Set F = FS.GetFile((Filespec1))
        NTIME = Format(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
        'FDS2 = Replace(FDS, "BBC WEATHER.JPG", NTIME)
        FDS2 = Mid(FDS, 1, InStr(FDS, "\URL SCREENSHOT") + Len("\URL SCREENSHOT"))
        FDS_END = Replace(FDS, FDS2, "")
        FDS2 = FDS2 + "- " + NTIME
        FDS2 = FDS2 + " " + FDS_END '" - BBC WEATHER.JPG"
        
        'If FileInUse(fds) = True Then
        Name FDS As FDS2
        'err.description
'        If MouseClicks > 0 Then Stop
        'If FS.FileExists(FDS2) = False Then Stop
        
            If MNU_PLAY_SOUND_ON_LOAD.Checked = True Or MODIFY_SOUND_SELECTION = True Then
        
                MMControl1.Command = "prev"
                MMControl1.Command = "Play"
            
            End If
            
'            If MouseClicks > 0 Then Stop
            If MouseClicks = 0 Then

                ED = FindWindow("IrfanView", vbNullString)
                If ED > 0 Then
                'Shell "C:\Program Files\IrfanView\i_view32.exe/killmesoftly"
                I = ExecCmdWAIT("C:\Program Files\IrfanView\i_view32.exe /killmesoftly")
                'i=0 NORMAL
            End If
            
            If IRFANVIEW_PATH <> "" Then
                If VAR1 = "IVIEW" Then
                    Shell IRFANVIEW_PATH + " """ + FDS2 + """ /fs /silent", vbMaximizedFocus
                End If
            Else
                Me.WindowState = vbNormal
                MsgBox "IRFANVIEW_PATH VAR -- PATH NOT FOUND FOR FILE" + vbCrLf + "NOT INSTALED AT EXPECTED LOCATION " + IRFANVIEW_PATH3 + vbCrLf + "OR" + vbCrLf + IRFANVIEW_PATH2, vbMsgBoxSetForeground
            End If

            '/reloadonloop
            
            Timer6.Enabled = True
            
'            Sleep 1000
            '/cmdexit
            '/closeslideshow - DONT WORK FOR WHAT WE WANT UNLESS SLIDESHOW
            'BOTH THESE DON'T WORK WHAT WE WANT AND BOTH TOGETHER EXIT PROGRAM BEFORE START
        End If
  
    End If
End If
On Error GoTo 0
    

End Sub

Private Sub Timer5_Timer()

End Sub

Private Sub Timer2_Timer()

End Sub

Private Sub Timer6_Timer()
'CALLED FROM SUB_RENAME_JPG_URL
Timer6.Enabled = False
Exit Sub

Dim I As Long
ED = FindWindow("IrfanView", vbNullString)
If ED > 0 Then
    ShowWindow ED, SW_SHOWMAXIMIZED
    SetForegroundWindow (ED)
'    Debug.Print "CODE SET FOREGROUND COMPLETE"
    I = GetForegroundWindow
    'If i <> ed Then Stop
    Timer6.Enabled = False
    
End If


End Sub


'-------------------------------------------------
'-------------------------------------------------
Private Sub Timer_1_MINUTE_Timer()

Timer_1_MINUTE.Interval = 60000

Call VB_PROJECT_CHECKDATE

End Sub
'-------------------------------------------------
'-------------------------------------------------
Sub VB_PROJECT_CHECKDATE()

'TOP DELCLARE -------------
'DIM XVB_DATE
'--------------------------

If Mid(App.Path, 1, 1) <> "D" Then Exit Sub

PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE'S\")

'---------------------------------------------------
'IF A NEW PROJECT NOT BEEN SYNC YET TO MIRROR FOLDER
'---------------------------------------------------
'MINOR WORK AROUND CREATE PATH
'---------------------------------------------------
If Dir(PATH_FILE_NAME2) = "" Then
    On Error Resume Next
        MkDir Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
        FS.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
    If Err.Number > 0 Then Exit Sub
End If

Set F = FS.GetFile(PATH_FILE_NAME2)

VB_DATE = F.DateLastModified

'--------
'01 OF 02
'----------------------------------
'WRITE INFO ABOUT DATE CHANGE NEWER
'----------------------------------

If XVB_DATE < VB_DATE And XVB_DATE > 0 Then
    
    '-----------------------------------
    'RUN A VBS TO COPY OVER AND RELOADER
    '-----------------------------------
    
'    FR1 = FreeFile
'    Open App.Path + "\RELOAD AND COPY.VBS" For Output As #FR1


    '----------------------------------
    'Mon 10 April 2017 16:30:50--------
    '----------------------------------
    'PROJECT REFERANCE ----------------
    'wshom.ocx
    'IF HARD FIND DO BROWSER
    'IN SYSTEM32 AND SYSWOW64
    'DIR wshom.ocx /S
    '----------------------------------
    'WINDOWS SCRIPT HOST OBJECT MODEL
    'AFTER FIND ALSO HAVE TO SELECT HER
    '----------------------------------
    '----------------------------------
    'Mon 10 April 2017 15:45:11--------
    '----
    'VISUAL BASIC wshom.ocx NOT WORKING - Google Search
    'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB721GB721&q=VISUAL+BASIC+wshom.ocx+NOT+WORKING&oq=VIUAL+BASIC+wshom.ocx+NOT+WORKING&gs_l=serp.3..30i10k1.1533.3987.0.4658.12.12.0.0.0.0.127.888.11j1.12.0....0...1c.1.64.serp..0.12.886...33i160k1j33i21k1.fYCPCpVPeVk
    '--------
    'WScript reference in VB
    'http://forums.codeguru.com/showthread.php?30458-WScript-reference-in-VB
    '----
    'Set objShell = WScript.CreateObject("WScript.Shell") - ------Not WORKER
    '----------------------------------
    '----------------------------------
    
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    WSHShell.Run """" + App.Path + "\RELOAD AND COPY.VBS" + """"
    Set WSHShell = Nothing
    Unload Me
    
End If

XVB_DATE = VB_DATE

End Sub

Private Sub TIMER_RETRY_WRITE_INFO_UNTIL_DONE1_Timer()
        
    'WRITE INFO ABOUT OUR PROGRAM TO BE RELOADED
    'PUT THE TIME DATE OF RELOAD REQUEST TO BE DONE
    'NOT MULTI COPY WHEN ERROR
    'NEED TO SAY WHAT COMPUTER WE ARE
    'OR MAKE SURE SYNC DONT SYNC THEM
    'FORMER OPTION
    
    On Error Resume Next
    VARTIME_RELOAD = "RELOAD ME -- " + Format(VB_DATE_REMOTE, "YYYY MM DD HH MM SS")
    
    FOLDERNAME2 = "D:\VB6-EXE'S\#00 RELOAD_MIRROR\" + GetComputerName
    If Dir(FOLDERNAME2, vbDirectory) = "" Then
        I = CreateFolderTree(FOLDERNAME2)
        If I = False Then Exit Sub
    End If
    
    FR1 = FreeFile
    Open "D:\VB6-EXE'S\#00 RELOAD_MIRROR\" + GetComputerName + "\" + VARTIME_RELOAD + ".TXT" For Output As #FR1
        'OUR APP PATH PROJECT EXE
        Print #FR1, PATH_FILE_NAME1
    Close #FR1
    
    If Err.Number = 0 Then
        TIMER_RETRY_WRITE_INFO_UNTIL_DONE1.Enabled = False
    End If

End Sub

Private Sub TIMER_RETRY_WRITE_INFO_UNTIL_DONE2_Timer()
    
    
    'APPEND A LOGG SOMEWHERE
    
    Exit Sub
    '-------------------------------------------------
    
    On Error Resume Next
    FR1 = FreeFile
    Open "D:\VB6-EXE'S\#00 RELOAD_APPEND_SCRIPT\## SCRIPT RELOAD APPEND.TXT" For Append As #FR1
        VARTIME = Format(Now, "YYYY MM DD HH MM SS") + " -- " + PATH_FILE_NAME1
    Close #FR1
    
    If Err.Number = 0 Then
        TIMER_RETRY_WRITE_INFO_UNTIL_DONE2.Enabled = False
    End If

End Sub

Sub TIMER_VB_PROJECT_DATE_TIMER()

TIMER_VB_PROJECT_DATE.Interval = 10000

Call VB_PROJECT_CHECKDATE

End Sub

Private Sub Timer3_Timer()
'Timer3.Enabled = False
'Exit Sub

'-----------------------------------
'COMPILER EXE VERSION UPTODATE TIMER
'-----------------------------------

'SORT EXE MAKE SURE MOST UPTO DATE OVER MIRROR FOLDER
'IMPROVE DON'T REQUIRE INDEX POINTER


'    Exit Sub
'    If IsIDE = False Then Exit Sub
    
    'Here 1 Second Timer
    'timer3.interval
    Timer3.Interval = 60000 ' A Min - BECUASE Too Much Disk Access at 1 Sec
    'Call Timer_APP_BEGIN_TIME_TIMER
    
'    Call MouseDetectMove
'    Call KeyboardDetectPress ' Also Done at Quicker Timer_MOUSE_KEYBOARD
    

    
    
    
    'FORM_STAY_UP_WITH_MSGBOX = STOP THE FORM WINDOW MINIMIZE WHEN A MSGBOX IS SHOWING
    
    If HHH < Now And HHH > 0 And FORM_STAY_UP_WITH_MSGBOX <> True Then
        Me.WindowState = vbMinimized: HHH = 0
    End If



    
    
    
'Filespec1 = App.Path + "\" + App.EXEName + ".vbp"
'Set F = FS.getfile((Filespec1))
'ADATE1 = F.datelastmodified
'Dim DateXVar
ADATE1 = 0
    
On Error GoTo END_EXIT
FileSpec2 = App.Path + "\" + App.EXEName + ".EXE"
Set F = FS.GetFile(FileSpec2)

ADATE1 = F.DateLastModified
APPEXEDATE = ADATE1
    
FILESPEC3 = App.Path + "\" + App.EXEName + ".EXE"
FILESPEC3 = Replace(FILESPEC3, ":\VB6\", ":\VB6-EXE'S\")

Set F = FS.GetFile(FILESPEC3)
ADATE_MIRROR_EXE_DATE = F.DateLastModified

'If ADATE_MIRROR_EXE_DATE <= ADATE1 Then
'    Exit Sub
'End If

Call Timer_EXE_DAY_AGE_TIMER
Call Timer_EXE_DAY_AGE_MIRROR_TIMER
    


Exit Sub



If APPEXEDATE < ADATE_MIRROR_EXE_DATE Then
    
    Shell App.Path + "\Network_Sync_EXE_Clipboard.exe """ + App.Path + "\" + App.EXEName + ".EXE" + """", vbNormalFocus
    
    
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub

End If
    
    
END_EXIT:

    
Exit Sub
    
'Dim Pro, TitleMSG As String

'Pro2 = " **"
'
'Pro2 = DateDiff("d", ADATE1, ADATE2)
'TIMERVAR = "Days"
'If Pro2 = 0 Then
'    Pro2 = DateDiff("h", ADATE1, ADATE2)
'    TIMERVAR = "Hours"
'End If
'TIMERVAR = "Hours"
'If Pro2 = 0 Then
'    Pro2 = DateDiff("n", ADATE1, ADATE2)
'    TIMERVAR = "Minutes"
'End If
'timervar2 = Mid(TIMERVAR, 1, 1)

'Pro = "Project EXE Compiled is Newer By ---[" + Str(Pro2) + " ]--- " + TIMERVAR + " -- " + Format(ADATE2, "DD-MMM-YYYY hh:mm:ss") + " -- " + Format(ADATE1, "DD-MMM-YYYY hh:mm:ss")
'Debug.Print Pro

'Mnu_Project_Source_Code.Caption = "New EXE Version is Here OR Going to Be Soon - New " + Format(ADATE2, "DD-MMM-YYYY hh:mm:ss") + " - Old " + Format(ADATE1, "DD-MMM-YYYY hh:mm:ss") + " -" + Str(Pro3) + " " + TIMERVAR3 + " Difference"
'MNU_VB.Caption = "VB ME" + Str(Pro3) + " " + TIMERVAR4

    
    
    Filespec1 = App.Path + "\" + App.EXEName + " - EXE VERSION DATE.NFO"
    If Dir(Filespec1) <> "" Then
        Set F = FS.GetFile((Filespec1))
        ADATE1 = F.DateLastModified
        If ADATE3_BEFORE = ADATE1 Then Exit Sub
        
        'ADATE3_BEFORE = NOTHING ON FIRST RUN
        ADATE3_BEFORE = ADATE1

        FR1 = FreeFile
        Open Filespec1 For Input As #FR1
            Line Input #FR1, DateXVar
        Close #FR1
    Else
        'NOT ANY DATE DATA YET SO GET AND WRITE SOME
        FileSpec2 = App.Path + "\" + App.EXEName + ".EXE"
        Set F = FS.GetFile((FileSpec2))
        ADATE2 = F.DateLastModified
        APPEXEDATE = ADATE2
        
        'We Got Our EXE MOST UPTODATE THAN BEFORE
        'BECAUSE IT WAS EMPTY BEFORE
        'SO UPDATE POINTER FILE
        FR1 = FreeFile
        Open Filespec1 For Output As #FR1
            Print #FR1, Format(ADATE2, "DD-MMM-YYYY hh:mm:ss")
        Close #FR1
        ADATE3_BEFORE = ADATE2
        'LATER A SYNC WILL SHOW WHAT IS MOST UPTODATE ONE
        
        'NOTHING LEFT TO DO HERE
        'BUT UPDATE SOME DATE DETAILS
        'Exit Sub
        
    End If
    
    If DateXVar <> "" Then
        Err.Clear
        On Error Resume Next
        ADATE1 = DateValue(DateXVar) + TimeValue(DateXVar)
        'ERROR READING FILE
        'TRY LATER
        If Err.Number > 0 Or ADATE1 = 0 Then
            Exit Sub
        End If
'        If ADATE1 = 0 Or Err.Number > 0 Then
'            Filespec2 = App.Path + "\" + App.EXEName + ".EXE"
'            Set F = FS.getfile((Filespec2))
'            ADATE2 = F.datelastmodified
'            APPEXEDATE = ADATE2
'        End If
'        On Error GoTo 0
        
    End If
    
    
    
    'READ IT THE DATE AGAIN
    FileSpec2 = App.Path + "\" + App.EXEName + ".EXE"
    Set F = FS.GetFile(FileSpec2)
    ADATE2 = F.DateLastModified
    APPEXEDATE = ADATE2
    
    'SURE BOTH GOT A DATE VALUE STEADY AHEAD
    
    Dim Pro, TitleMSG As String
    
    Pro2 = " **"
    
    
    '-----------
    '1. OF 2
    'IF THE EXE IS NEWER THAN THE TXT DATE DATA FILE POINTER
    'THEN
    'UPDATE THE POINTER AND LET OTHER NETWORK COMPUTER KNOW
    '-----------
    '2. OF 2
    'IF THE POINTER IS NEWER THAN THE EXE VERSION
    'FIND THE NEW VERSION
    'CHAIN THE PROGRAM THAT WILL COPY IT OVER
    '-----------
    'NOTE 1 OF %
    'IF THE EXE IS BEEN RUNNING LONG WON'T FIND ANY NEW EXE
    'ONLY THE POINTER
    
    
    '------------
    '01 OF 03
    '------------
    'IS EXE NEWER
    
'    Dim Pro, TitleMSG As String
    
    Pro2 = " **"
    If ADATE2 > ADATE1 Then
        
        If DateDiff("n", ADATE1, ADATE2) > 0 Then
            
            'DUD INFO
            'Pro = "Project Source Code Is More Advanced By ---[" + str(DateDiff("n", ADATE2, ADATE1)) + " ]--- Minutes Difference Than The Executable Program"
            
            Pro2 = DateDiff("d", ADATE1, ADATE2)
            TIMERVAR = "Days"
            If Pro2 = 0 Then
                Pro2 = DateDiff("h", ADATE1, ADATE2)
                TIMERVAR = "Hours"
            End If
            TIMERVAR = "Hours"
            If Pro2 = 0 Then
                Pro2 = DateDiff("n", ADATE1, ADATE2)
                TIMERVAR = "Minutes"
            End If
            timervar2 = Mid(TIMERVAR, 1, 1)
            
            Pro = "Project EXE Compiled is Newer By ---[" + Str(Pro2) + " ]--- " + TIMERVAR + " -- " + Format(ADATE2, "DD-MMM-YYYY hh:mm:ss") + " -- " + Format(ADATE1, "DD-MMM-YYYY hh:mm:ss")
            'Debug.Print Pro
        
        End If
    
        'We Got Our EXE MOST UPTODATE THAN BEFORE
        'SO UPDATE POINTER FILE
        'Close FR1
'        Reset
        
'        FR1 = FreeFile
'        Open Filespec1 For Output As #FR1
'        Print #FR1, Format(ADATE2, "DD-MMM-YYYY hh:mm:ss")
'        Close #FR1
'        ADATE3_BEFORE = ADATE2
        
        
        'START THE TIMER CODE
        Timer_EXE_DAY_AGE.Enabled = True
        Call Timer_EXE_DAY_AGE_TIMER
    
    End If
    
    '------------
    '02 OF 03
    '------------
    'IS POINTER SHOWING NEWER THAN EXE -- AVAILABLE HERE
    
    If ADATE2 < ADATE1 Then
    
        'Mnu_Project_Source_Code.Caption = "Check For More Up-to-date Project Source Code =" + Pro2
        
        Timer_EXE_DAY_AGE.Enabled = False
        Mnu_Project_Source_Code.Caption = "New EXE Version is Here OR Going to Be Soon - New " + Format(ADATE2, "DD-MMM-YYYY hh:mm:ss") + " - Old " + Format(ADATE1, "DD-MMM-YYYY hh:mm:ss") + " -" + Str(Pro3) + " " + TIMERVAR3 + " Difference"
        DATE_CODE = TIMERVAR4
        DATE_CODE = Replace(DATE_CODE, "D", "Day")
        DATE_CODE = Replace(DATE_CODE, "M", "Minute")
        MNU_VB.Caption = "VB ME" + Str(Pro3) + " " + DATE_CODE
        
        'WE NEED TO SWAP OUR NEW EXE IN
        'LOOK FOR OUR EXE IN MIRROR SYNC FOLDER IS HERE FIRST
        
        FILESPEC3 = App.Path + "\" + App.EXEName + ".EXE"
        'MIRROR
        FILESPEC3 = Replace(FILESPEC3, ":\VB6\", ":\VB6-EXE'S\")
        Set F = FS.GetFile(FILESPEC3)
        ADATENEWEXE = F.DateLastModified
        
        If ADATENEWEXE < ADATE2 Then
        
            'TAKE ACTION TO COPY IT OFF AND RESTART
            'SPEED IS RIGHT
            'WAIT FOR PROGRAM TO CLOSE
            'WHY NOT A BATCH FILE
                
            'i = MsgBox(Mnu_Project_Source_Code.Caption + vbCrLf + "Do You Want to Kick Start The New Version", vbYesNo + vbMsgBoxSetForeground)
            
            'Debug.Print App.Path + "\Network_Sync_EXE_Clipboard.exe """ + App.Path + "\" + App.EXEName + ".EXE" + """"
            Shell App.Path + "\Network_Sync_EXE_Clipboard.exe """ + App.Path + "\" + App.EXEName + ".EXE" + """", vbNormalFocus
            
            
            
            EXIT_TRUE = True
            Unload Me
            Exit Sub
        
        
        
        
        End If
                
                
        
        
        
        
        
        
    
    End If
    
    
    
    
    '------------
    '03 OF 03
    '------------
    'POINTER DATE AND EXE DATE ARE EQUAL
    
    If ADATE2 = ADATE1 Then
    
        'START THE TIMER CODE
        Timer_EXE_DAY_AGE.Enabled = True
        Call Timer_EXE_DAY_AGE_TIMER
    
    
    End If
    
    '-----------------------------------------
    Exit Sub
    '-----------------------------------------
    
    
'    If iMessageResultRECompile <> 10 Then
'        If DateDiff("n", ADATE2, ADATE1) > 10 And IsIDE = False Then
'            'TitleMSG = Me.Caption
'            Me.WindowState = vbNormal
'            Do
'                DoEvents
'                Me.Refresh
'                DoEvents
'                Sleep 100
'            Loop Until Me.WindowState <> vbMinimized
'
'            iMessageResultRECompile = MsgBox(Pro + vbCrLf + "Project ExE Code Is More Up-to-Date" + vbCrLf + "Run the VB IDE to Compile?", vbYesNo + vbMsgBoxSetForeground)
'            If iMessageResultRECompile = vbYes Then Call MNU_VB_Click
'            iMessageResultRECompile = 10
'        End If
'    End If
    
    
    
    

    
    
    
    
    
'    Filespec1 = App.Path + "\01 Woarble Tone 5 Mins 01.wav"
'    Set F = FS.getfile((Filespec1))
'    ADATE1 = F.datelastmodified
'    Filespec2 = App.Path + "\01 Woarble Tone 5 Mins 02.wav"
'    Set F = FS.getfile((Filespec2))
'    ADATE2 = F.datelastmodified
'    If ADATE2 > ADATE1 Then
'
'        ATidalX.MMControl2.Command = "Close"
'
'        FS2.CopyFile Filespec2, Filespec1
'
'        ATidalX.MMControl2.Notify = True
'        ATidalX.MMControl2.Wait = True
'        ATidalX.MMControl2.Shareable = False
'        ATidalX.MMControl2.DeviceType = "waveaudio"
'        ATidalX.MMControl2.FileName = App.Path + "\01 Woarble Tone 5 Mins 01.wav"
'        ATidalX.MMControl2.Command = "Open"
'
'    End If

End Sub


Private Sub Timer_APP_BEGIN_TIMER_Timer()

        'Timer_APP_BEGIN_TIME.Interval = 1000
        
        'JOINT SUB ROUTINE MAYBE -- CHECK -- MAYBE REUSED CODE
        
        'IF IN DAY MODE - - JOINT SUB ROUTINE
        Pro3 = DateDiff("d", ADATE_APP_BEGIN_DATE, Now)
        TIMERVAR3 = "Day"
        
        'DAY
        PROX = Str(DateDiff("d", ADATE_APP_BEGIN_DATE, Now)) + " " + Mid(TIMERVAR3, 1, Len(TIMERVAR3) - 1)
        
        'IF IN HOUR MODE - - JOINT SUB ROUTINE
        If Pro3 = 0 Then
            Pro3 = DateDiff("h", ADATE_APP_BEGIN_DATE, Now)
            TIMERVAR3 = "Hour"
        End If
        
        'HOUR
        PROX = PROX + Str(DateDiff("h", ADATE_APP_BEGIN_DATE, Now)) + " " + Mid(TIMERVAR3, 1, Len(TIMERVAR3) - 1) + " "
        
        'IF IN MINUTE MODE - - JOINT SUB ROUTINE
        If Pro3 = 0 Then
            Pro3 = DateDiff("n", ADATE_APP_BEGIN_DATE, Now)
            TIMERVAR3 = "Minute"
        End If
        
        'MINUTE
        PROX = PROX + Format((DateDiff("n", ADATE_APP_BEGIN_DATE, Now)), "00") + " " + Mid(TIMERVAR3, 1, Len(TIMERVAR3) - 1) + " "
        
        'IF IN SECDOND MODE - - JOINT SUB ROUTINE
        
        If Pro3 = 0 Then
            pro4 = DateDiff("s", ADATE_APP_BEGIN_DATE, Now)
            TIMERVAR4 = "Second"
        Else
            pro4 = ""
            TIMERVAR4 = ""
        End If
        
        'SEC
        PROX = PROX + Format((DateDiff("s", ADATE_APP_BEGIN_DATE, Now)), "00") + " " + Mid(TIMERVAR3, 1, Len(TIMERVAR3) - 1)
        
        
        '---------------------------------------
        'USED SOMEWEHRE ELSE NOT THIS SUBROUTINE
        'FIRST CHAR OF TIME ID
        '---------------------------------------
        TIMERVAR4 = Mid(TIMERVAR3, 1, 1)
        
        Mnu_APP_BEGIN_TIME_TIMER.Caption = "CLIPBOARD LOGGER BEGIN TIME = " + Format(ADATE_APP_BEGIN_DATE, "DD-MMM-YYYY hh:mm:ss") + " --" + Str(Pro3) + " " + TIMERVAR3 + " Age -- STARUP TIME SECONDS =" + Str(ADATE_APP_BEGIN_DATE2)
        
        MNU_RUN_TIME.Caption = "RUN TIME " + PROX


End Sub


Sub Timer_EXE_DAY_AGE_MIRROR_TIMER()

        'THIS NOT SEEM RUN AT START

        Pro3 = DateDiff("d", ADATE_MIRROR_EXE_DATE, Now)
        TIMERVAR3 = "Days"
        If Pro3 = 0 Then
            Pro3 = DateDiff("h", ADATE_MIRROR_EXE_DATE, Now)
            TIMERVAR3 = "Hours"
        End If
        If Pro3 = 0 Then
            Pro3 = DateDiff("n", ADATE_MIRROR_EXE_DATE, Now)
            TIMERVAR3 = "Minutes"
        End If
        
        TIMERVAR4 = Mid(TIMERVAR3, 1, 1)
        
        Mnu_Project_Source_Code_MIRROR.Caption = "EXE Compiler Code Current Version = " + Format(ADATE_MIRROR_EXE_DATE, "DD-MMM-YYYY hh:mm:ss") + " --" + Str(Pro3) + " " + TIMERVAR3 + " Age -- MIRROR VERSION"
        
        'MNU_VB.Caption = "VB ME" + Str(Pro3) + " " + TIMERVAR4

End Sub

Sub Timer_EXE_DAY_AGE_TIMER()

        
        If APPEXEDATE = 0 Then 'Nothing Then
            
            
            Mnu_Project_Source_Code.Caption = "EXE Compiler Code Current Version = VAR FOR APPEXEDATE HAS A NUL"
            MNU_VB.Caption = "VB ME -- NULL"
            
            Exit Sub
        End If
        

        Pro3 = DateDiff("d", APPEXEDATE, Now)
        TIMERVAR3 = "Days"
        If Pro3 = 0 Then
            Pro3 = DateDiff("h", APPEXEDATE, Now)
            TIMERVAR3 = "Hours"
        End If
        If Pro3 = 0 Then
            Pro3 = DateDiff("n", APPEXEDATE, Now)
            TIMERVAR3 = "Minutes"
        End If
        TIMERVAR4 = Mid(TIMERVAR3, 1, 1)
        
        Mnu_Project_Source_Code.Caption = "EXE Compiler Code Current Version = " + Format(APPEXEDATE, "DD-MMM-YYYY hh:mm:ss") + " --" + Str(Pro3) + " " + TIMERVAR3 + " Age"
        
        DATE_CODE = TIMERVAR4
        DATE_CODE = Replace(DATE_CODE, "D", "Day")
        DATE_CODE = Replace(DATE_CODE, "M", "Minute")

        MNU_VB.Caption = "VB_ME FILE_EXE_CREATION" + Str(Pro3) + " " + DATE_CODE

End Sub


Public Function FindShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    strPath = String$(165, 0)
    lngRes = GetShortPathName(strFileName, strPath, 164)
    FindShortPath = Left$(strPath, lngRes)
End Function

Sub Timer_KEYBOARD_Active_Timer()

    '-------------------------------------------
    If EXIT_TRUE = True Then Unload Me: Exit Sub

    If OIK2 = False Then
        IK2 = KeyboardDetectPress
        OIK2 = IK2
    End If
    
    'TestLowTimer = True
    If IsIDE = True And TestLowTimer = True Then
        Timer_KEYBOARD_Active.Interval = 10000
    End If
    
    Exit Sub
    
    
    'BIT COMPLEX IF KEYBOARD ONLY ACTIVE MAYBE
    'IF USE MOUSE KEY WHY MORE OR LESS THAN WHEN NOT
    '-----------------------------------------------
    If KEYBOARD_OR_MOUSE_ACTIVE_3_MIN > Now Then
        active1 = True
    Else
        active1 = False
    End If
    
    If IsIDE = True And active1 = True Then
        Timer_KEYBOARD_Active.Interval = 5000
    Else
        Timer_KEYBOARD_Active.Interval = 50
    End If
    

End Sub



Private Sub Timer_MOUSE_KEYBOARD_Timer()
    
    'Debug.Print Me.Enabled
    
    'Timer_MOUSE_KEYBOARD.Interval = 1000
    
    IM1 = MouseDetectMove
    'IK2 = KeyboardDetectPress
    
    If IK2 = False And IM1 = False Then Exit Sub
    OIK2 = False
    
    TIMER_KEYBOARD_AND_MOUSE_ACTIVE.Enabled = False
    TIMER_KEYBOARD_AND_MOUSE_ACTIVE.Interval = 59000
    TIMER_KEYBOARD_AND_MOUSE_ACTIVE.Enabled = True
    OverRideOnceExtra = True

    KEYBOARD_OR_MOUSE_ACTIVE = Now + TimeSerial(0, 0, 5)
    KEYBOARD_OR_MOUSE_ACTIVE_3_MIN = Now + TimeSerial(0, 3, 0)
    KEYBOARD_OR_MOUSE_ACTIVE_10_SEC = Now + TimeSerial(0, 0, 10)

    Mouse_Keyboard_Active_Time = Now
    
    '--------------------------------------
    'MAKE SURE NOT RESETING TIMER TO BEGIN AGAIN
    'USE ENABLE FALSE AND THEN TRUE TO DO THAT NORMALLY
    'CHANGE TIME TO 2 MIN BECAUSE ERROR GETTING
    Timer_MOUSE_1_MINUTE.Enabled = False
    Timer_MOUSE_1_MINUTE.Interval = 5000
    Timer_MOUSE_1_MINUTE.Enabled = True
    
    '----------------------
    'GRAB THIS BEFORE TAKEN
    '----------------------
    'WHAT FOR
    '----------------------
'    GO3 = 0
'
'    If InStr(MNU_IDLE_ACTIVE.Caption, "1 MIN") Then GO3 = 1
'
'    If MNU_IDLE_ACTIVE.Caption = "IDLE" Then GO3 = 1
'
'    If MNU_IDLE_ACTIVE.Caption <> MNU_IDLE_ACTIVE_VAR Then
        MNU_IDLE_ACTIVE.Caption = "IDLE STATE = ACTIVE"
'        MNU_IDLE_ACTIVE_VAR = MNU_IDLE_ACTIVE.Caption
'    End If
    
    If Me.WindowState <> vbMinimized Then
    
        'IF IT GOT THIS FAR INTO A MIN AND NOT ANY CLIPBOARD YET
        'AND NOT vbMinimized
        'MAYBE MAKE IT QUICKER AT BEEN TO IDLE DO TRIGGER
    
        '"1st Run" IS ALL IT CAN BE FOR ANSWER
        
        If Mnu_Run_Status.Caption = "1st Run" And GO3 = 1 Then
            Mnu_Run_Status.Visible = False
            Mnu_Clip_Status.Visible = False
        End If

    
    End If
    
    '-----------------------------------
    Timer_MOUSE_2_MINUTE.Enabled = False
    Timer_MOUSE_2_MINUTE.Interval = 59990 ' 2 MIN - TOP UP WHEN ACTIVE
    Timer_MOUSE_2_MINUTE.Enabled = True
    '-----------------------------------
    Timer_MOUSE_3_MINUTE.Enabled = False
    '--------------------------------------


End Sub

Private Sub Timer_EXPLORER_GONE_Timer()


If GETWinNT_Ver <> "WIN XP" Then
    Timer_EXPLORER_GONE.Enabled = False
    Exit Sub
End If


'1st FIT IN ANOTHER SUBROUTINE
'Call MNU_CLIPBOARD_API_PUBLIC_VAR_HOOK_Click


'-------------------
'If Explorer Crashes
'-------------------

'----------------
'EXPLORER DESKTOP
'----------------
If FindWindow("Progman", "Program Manager") = 0 Then

    ExplorerGone = True
    
    Me.WindowState = Normal
    DoEvents
    Me.Refresh
    DoEvents
    Me.SetFocus
    DoEvents
    Me.Refresh
    DoEvents
    
    
    'ONLY REQUIRE WIN XP
    FORM_STAY_UP_WITH_MSGBOX = True
    I = MsgBox("Do You Want to Reload Explorer, Crash -- Disappeared", vbYesNo + vbMsgBoxSetForeground)
    FORM_STAY_UP_WITH_MSGBOX = False

'    Me.WindowState = Normal

    If I = vbYes Then
        'Shell "c:\windows\Explorer.exe", vbNormalFocus
        
        'cmdLine$ = "c:\windows\Explorer.exe"
        'i = ExecCmd(cmdLine$)
    
        Shell "c:\windows\Explorer.exe", vbNormalFocus
    
        Do
            i2 = FindWindow("Progman", "Program Manager")
            DoEvents
            'Sleep 500
        Loop Until i2 <> 0
        A = A
    End If
    
    Exit Sub

End If

If FindWindow("Progman", "Program Manager") <> 0 And ExplorerGone = True Then

    'BRING WINDOWS FRONT
    I = FindWinPartExplorerGone(False) ' True = Quite Mode Don't Display  Result
'    Debug.Print str(i) + " Windows Brought Forward"
 
    MNU_BRing_Front.Caption = "Bring All Windows Front -- Explorer Crash/Terminated @ " + Format(Now, "DD-MMM-YYYY HH:MM:SS")
    ExplorerGone = False

End If



End Sub

Private Sub TimerSCROLL_Timer()

Exit Sub
'HIDEN REMOVED FROM MENU

'This is Not Enabled unless Selected from the Menu
'SOMETHING TO DO - NOTEPAD2 SCROLL DOWN

I = 0 ': KASC = 0
For I = 1 To 255
    BDF = GetAsyncKeyState(I)
    If BDF < -300 Then KCODE = Now + TimeSerial(0, 0, 3): Exit Sub
Next

'If KASC > 0 Then Stop

If KCODE > Now Then Exit Sub

NOTT = FindWindow("Notepad2", vbNullString)
If NOTT <> GetForegroundWindow Then Exit Sub

SendKeys "{down}", False

End Sub

Private Sub TIMER_KEYBOARD_AND_MOUSE_ACTIVE_TIMER()
    '1 MIN FOR ACTIVE
    TIMER_KEYBOARD_AND_MOUSE_ACTIVE.Enabled = False
    OverRideOnceExtra = True

End Sub

Private Sub TIMER_MOUSE_Timer()

    'Call Timer_SCREEN_SHOT_Timer
'------------------------------
    
    Timer_MOUSE.Enabled = False
    MouseClicks = 0
    Mousey = 0
    
    If Mnu_Exit_VAR <> "Exit" Then
        Mnu_Exit.Caption = "Exit"
        Mnu_Exit_VAR = Mnu_Exit.Caption
    End If
    
    'TimerCheckIntegrityOfProgram = True
    
'    OverRideOnceExtra = True
    
    'Form2.Label1 = Str(Val(Int(MouseClicks))) + " --" + Str(Val(Int(Mousey)))

End Sub

Public Sub FindCursor(x, y)

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
x = P.x ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
y = P.y '/ GetSystemMetrics(1) * Screen.Height

End Sub

Function KeyboardDetectPress()

Dim BDF, i2
Dim hwndtest As Long

KeyboardDetectPress = False

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer

'Private Type POINTAPI
'        X As Long
'        Y As Long
'End Type


i2 = 0: KASC = 0
'WIN 10 IS GIVE -- GetAsyncKeyState(255) ALWAYS PRESSED
For i2 = 1 To 254
    BDF = GetAsyncKeyState(i2)
    If BDF < -300 Then
        'KCODE = Now + TimeSerial(0, 0, 3): Exit Sub
    
        '----------------------------------------
        'DON'T USE F5 -- WHEN FOR DEBUGGER RESUME
        If i2 = 116 Then Exit Function
        '----------------------------------------
    
        'Debug.Print i2
        
        If Mnu_Exit_VAR <> "Exit *" Then
            Mnu_Exit.Caption = "Exit *"
            Mnu_Exit_VAR = Mnu_Exit.Caption
        End If
        
        Timer_MOUSE.Enabled = False
        Timer_MOUSE.Interval = 5000
        Timer_MOUSE.Enabled = True
        KeyboardDetectPress = True

        KASC = i2
        Exit For
    End If
Next


'If KASC > 0 Then Stop
'DEBUG.PRINT INFO


Dim I As Long
'ESCAPE KEY
If KASC = 27 And KASC_TRIGGER <> KASC Then
    I = FindWindow(vbNullString, "MSDN Library Visual Studio 6.0")
    If I > 0 And I = GetForegroundWindow Then
        
        'CRAPPER
        'SendKeys "{%}{F4)", True
        
        TargetHwnd = I
        TargetHwnd = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
        
        'PROCESS CLOSE KILL CRAPPER
        'Call CloseProgram(i, , False)
    
    End If
End If


KASC_TRIGGER = KASC


'    If KCODE > Now Then Exit Sub
'
'
'    NOTT = FindWindow("Notepad2", vbNullString)
'    If NOTT <> GetForegroundWindow Then Exit Sub
'
'    SendKeys "{down}", False
End Function



Function MouseDetectMove()

'Exit Function


MouseDetectMove = False

FindCursor nx, ny

'This Will Happen to Mouse If User Is Logged off
'If Nx = 0 And Ny = 0 Then
'LockSSaver = Now + LockSaverDelay
'Winamp Video
'Call SetLockMax
'Call ProgressLock
'End If

'If Nx = 0 Or Ny = 0 Then Exit Sub

'If Me.WindowState = vbMinimized Then
'    MouseClicks = 0
'    Mousey = 0
'    Mnu_Exit.Caption = "Exit"
'    Exit Sub
'End If

If Xmp5 = -1 And Ymp5 = -1 Then
    Xmp5 = nx: Ymp5 = ny
End If


'If Quake = 0 Then
    If (nx <> Xmp5 Or ny <> Ymp5) Then 'And (Nx <> ScreenWidth And Ny <> ScreenHeight) Then
        'Call WinonPoint
        Mousey = Mousey + Sqr(Abs(Xmp5 - nx) ^ 2 + Abs(Ymp5 - ny) ^ 2)
        MouseClicks = Sqr(Abs(Xmp5 - nx) ^ 2 + Abs(Ymp5 - ny) ^ 2)
                    
'        If MouseFirstRun = 0 Then
'            MouseFirstRun = 1
'            Mousey = 0
'        End If
    
        'Form2.Label1 = str(Val(Int(MouseClicks))) + " --" + str(Val(Int(Mousey)))
        
        'If MouseClicks > 3 Then
        If Mousey > 3 Or MouseClicks > 3 Then

'           Label21.Caption = Str$(Mousey)
            'THIS READ OF A MENU BAR -- MAKE MENU BAR SCREEN FLICKER IN THE BAR
            'STORE RESULT WANT IN A VAR INSTEAD
            ' -- NOT WORK
            ' -- STILL FLICKER
            
            If Mnu_Exit_VAR <> "Exit *" Then
                Mnu_Exit.Caption = "Exit *"
                Mnu_Exit_VAR = Mnu_Exit.Caption
            End If
            
            Timer_MOUSE.Enabled = False
            Timer_MOUSE.Interval = 5000
            Timer_MOUSE.Enabled = True
            
            'THIS IS SOURCE OF MAKE MENU BAR SCREEN FLICKER
            MouseDetectMove = True
        
        End If
        
        Xmp5 = nx: Ymp5 = ny
    
    End If
'End If

End Function


Sub Menu_clipboard_size(VarSize As Double)
Dim TTX, TTG

    If Clipboard.GetFormat(vbCFBitmap) = True Then SIZE_IMAGE_NOT_MATTER = True


    'WORK TO DO HERE MORE THAN TEXT AN IMAGE
    
    
    If Clipboard.GetFormat(vbCFText) = True Then TTX = "Text" Else TTX = "Image"
'    MNU_CB_SIZE.Caption = Format(VarSize / 1024 ^ 2, "0.0000") + " Meg " + TTX
    MNU_CB_SIZE_MEG.Caption = Format(VarSize / 1024 ^ 2, "0.000") + " Meg " + TTX
    MNU_CB_SIZE_BYTE.Caption = Format(VarSize, "0") + " Byte " + TTX
    '----------------------------------------------
    
    If SIZE_IMAGE_NOT_MATTER = False Then
        If LimitClipSize > 0 Then
            If VarSize > LimitClipSize Then
                MNU_CB_SIZE_BYTE_OVERSIZE.Visible = True
                MNU_CB_SIZE_BYTE_OVERSIZE.Caption = "CLIP OVERSIZE LIMIT -" + Str(VarSize) + ">" + Trim(Str(Int(LimitClipSize / 1000))) + " K"
            Else
                MNU_CB_SIZE_BYTE_OVERSIZE.Visible = False
            End If
        End If
    End If
    
    '----------------------------------------------
    
    
    TTG = Format(VarSize, "###,###,##0") + " Byte " + TTX
    'Mnu_LAST_CLIP_TIME.Caption = "@ " + Format(Now, "DD-MM-YYYY-HH:MM:SS")
    Mnu_LAST_CLIP_TIME.Caption = "Last Clip__" + Format(Now, "MMMM DD-MMM-YYYY__HH:MM:SS__HH:MMa/p")

    If Int(VarSize / 1024 ^ 2) > 3 Then
        'TOP MAIN MENU
        
        MNU_CB_SIZE.Caption = MNU_CB_SIZE_MEG.Caption + " && Mnu Opt"
        'Hitt Menu The Option Below to Copy and Paste

    Else
        'TOP MAIN MENU
        MNU_CB_SIZE.Caption = TTG + " && Mnu Opt" 'MNU_CB_SIZE_BYTE.Caption
    End If
End Sub



Sub TRIM_UP_FORM_LOGG_CAPTURE()
    
    
    '------------------------------------------------------
    'SUBPATH IN USE THEN DO LATER
    '------------------------------------------------------
    If ScanPath.ListView1.ListItems.Count > 0 Then Exit Sub
    '------------------------------------------------------
    
    
    '--------------------------------------------------------
    'THIS NEED BE BETTER MORE ROOT LEVEL PATH SCAN AND FILTER
    'THINK BY FATE SOME GET LEFT BEHIND
    'TAKING UP SPACE
    '2 VARIABLE SET ONE WITH PARTICAL SCAN PATH
    '--------------------------------------------------------
    
    'LAST_FILE_DATE_TIME = Now - 2
    'LAST_FILE_DATE_TIME = Now - TimeSerial(0, 30, 0)
    
    
    
    'COUNT OF SAVES SINCE PROGRAM BEGIN
    'COUNT_MESSURE = 2880
    COUNT_MESSURE = 10
    COUNT_MESSURE = 10
'    COUNT_MESSURE = 0
    
    ScanPath.cboMask = "*.JPG"
    ScanPath.chkSubFolders = vbChecked
    '---------------------------------------------------
    'FORM SWAPING
    ScanPath.TxtPath = FF_FORM1
    If FF_COUNT_FORM_CAPTURE1 > COUNT_MESSURE Then
        FF_COUNT_FORM_CAPTURE1 = 0
        LAST_FILE_DATE_TIME = Now - TimeSerial(0, 30, 0)
        ScanPath.CMDScan_NO_LIST_FAST_FORM_CAPTURE_Click
    End If
    '---------------------------------------------------
    'MOUSE / KEY RESUME FROM IDLE OR ENTER ACTIVE
    ScanPath.TxtPath = FF_FORM2
    If FF_COUNT_FORM_CAPTURE2 > COUNT_MESSURE Then
        FF_COUNT_FORM_CAPTURE2 = 0
        LAST_FILE_DATE_TIME = Now - TimeSerial(0, 30, 0)
        ScanPath.CMDScan_NO_LIST_FAST_FORM_CAPTURE_Click
    End If
    '---------------------------------------------------
    'ALWAYS 1 MIN
    ScanPath.TxtPath = FF_FORM3
    If FF_COUNT_FORM_CAPTURE3 > COUNT_MESSURE Then
        FF_COUNT_FORM_CAPTURE3 = 0
        LAST_FILE_DATE_TIME = Now - TimeSerial(0, 30, 0)
        ScanPath.CMDScan_NO_LIST_FAST_FORM_CAPTURE_Click
    End If
    '---------------------------------------------------
    'ALWAYS 10 MINS WHOLE SCREEN
    ScanPath.TxtPath = FF_FORM4
    If FF_COUNT_FORM_CAPTURE4 > COUNT_MESSURE Then
        FF_COUNT_FORM_CAPTURE4 = 0
        LAST_FILE_DATE_TIME = Now - TimeSerial(0, 80, 0)
        ScanPath.CMDScan_NO_LIST_FAST_FORM_CAPTURE_Click
    End If
    '---------------------------------------------------
    'ALWAYS 1 HOUR WHOLE SCREEN
    '---------------------------------------------------
    'ALWAYS MOUSE FLOAT OVER FORM SWAPPER
    '---------------------------------------------------

End Sub


Private Sub Timer_SCREEN_SHOT_Timer()

Timer_SCREEN_SHOT.Interval = 1000

'IDLE KEY SCREEN SHOT

Dim X22 As Long

YES_VAR = False
'If GetComputerName = "ASUS-BIGGER" Then YES_VAR = True
If GetComputerName = "4-ASUS-GL522VW" Then YES_VAR = True
If GetComputerName = "5-ASUS-P2520LA" Then YES_VAR = True




'TestLowTimer = True
'If TestLowTimer = True Then YES_VAR = False
If YES_VAR = False Then Timer_SCREEN_SHOT.Enabled = False: Exit Sub

'TO DO SCREEN OR FORM SHOT ONLY WHEN MOUSE / KEY
'ACTIVE RESUMED OR FINISHED IDLE AFTER 1 MIN
'VAR HERE
'If TIMER_KEYBOARD_AND_MOUSE_ACTIVE.Enabled = True Then YES_VAR = True

If TIMER_KEYBOARD_AND_MOUSE_ACTIVE.Enabled = True Then
    SCREEN_SHOT_VAR_MOUSE_KEY_ACTIVE = Now + TimeSerial(0, 15, 0)
End If

If SCREEN_SHOT_VAR_MOUSE_KEY_ACTIVE < Now Or 1 = 1 Then
    If SCREEN_SHOT_VAR_MOUSE_KEY_ACTIVE > 0 Or 1 = 1 Then
            SCREEN_SHOT_VAR_MOUSE_KEY_ACTIVE = 0
            Call TRIM_UP_FORM_LOGG_CAPTURE
    End If

    'TRIM UP THE BIG LOGG
    'FF_FORM1 = FORM_SWAP 1000
    'FF_FORM2 = MOUSE KEY 1000 '' OR 3 DAY WHAT EVER IS HIGHER
    'FF_FORM3 = FORM    1 MIN 800
    'FF_FORM4 = SCREEN 10 MIN 500
End If


DO_FORM = False: DO_SCREEN = False
YES_VAR = False
'PRIORITY 1 FORM SWAP
X22 = GetForegroundWindow
If OX22 <> X22 Then YES_VAR = True
OX22 = X22
If YES_VAR = True Then
    DO_FORM = True
    FF2$ = "FormCapture_" + Format$(Now, "YYYY-MM-DD-DDD") + "_FORM_SWAP\"
    FOLDERNAME_AUTO = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\AUTO_Form_Shot\" + FF2$
    'Debug.Print FOLDERNAME_AUTO
    If FS.FolderExists(FOLDERNAME_AUTO) = False Then
        'If IsIDE = True Then Stop
        I = CreateFolderTree(FOLDERNAME_AUTO)
        If I = False Then
            MSGBOX2 = "ERROR PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME_AUTO
            FRM_MSGBOX.Timer1.Enabled = True
            Exit Sub
        End If
    End If
    
    FF_FORM1 = FOLDERNAME_AUTO
    FF_COUNT_FORM_CAPTURE1 = FF_COUNT_FORM_CAPTURE1 + 1
    If FS.FolderExists(FOLDERNAME_AUTO) = False Then
        'If IsIDE = True Then Stop
        I = CreateFolderTree(FOLDERNAME_AUTO)
        If I = False Then
            MSGBOX2 = "ERROR PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME_AUTO
            FRM_MSGBOX.Timer1.Enabled = True
            Exit Sub
        End If
    End If
End If

'PRIORITY 2 MOUSE / KEY RESUME FROM IDLE OR ENTER ACTIVE program
If OverRideOnceExtra = True And YES_VAR = False Then
    DO_FORM = True
    OverRideOnceExtra = False
    FF2$ = "FormCapture_" + Format$(Now, "YYYY-MM-DD-DDD") + "_MOUSE_KEY_BUSY_IDLE_OR_ENTER_ACTIVE\"
    FOLDERNAME_AUTO = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\AUTO_Form_Shot\" + FF2$
    FF_FORM2 = FOLDERNAME_AUTO
    FF_COUNT_FORM_CAPTURE2 = FF_COUNT_FORM_CAPTURE2 + 1
End If

'PRIORITY 3 OF 4 ALWAYS 1 MIN
If YES_VAR = False Then
    X22 = Minute(Now)
    If OX25 <> X22 Then YES_VAR = True
    OX25 = X22
    If YES_VAR = True Then
        DO_FORM = True
        FF2$ = "FormCapture_" + Format$(Now, "YYYY-MM-DD-DDD") + "_MINUTE\"
        FOLDERNAME_AUTO = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\AUTO_Form_Shot\" + FF2$
        FF_FORM3 = FOLDERNAME_AUTO
        FF_COUNT_FORM_CAPTURE3 = FF_COUNT_FORM_CAPTURE3 + 1
        If FS.FolderExists(FOLDERNAME_AUTO) = False Then
            I = CreateFolderTree(FOLDERNAME_AUTO)
            If I = False Then
                MSGBOX2 = "ERROR PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME_AUTO
                FRM_MSGBOX.Timer1.Enabled = True
                Exit Sub
            End If
        End If
    End If
End If

'PRIORITY 4 OF 4 ALWAYS 10 MINS WHOLE SCREEN
If YES_VAR = False Then
    X22 = Val(Mid(Format(Minute(Now), "00"), 2, 1))
    If OX24 <> X22 And X22 = 0 Then YES_VAR = True
    OX24 = X22
    If YES_VAR = True Then
        DO_SCREEN = True
        FF2$ = "ScreenCapture_" + Format$(Now, "YYYY-MM-DD-DDD") + "_MINUTE_LONG\"
        FOLDERNAME_AUTO = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\AUTO_Screen_Shot\" + FF2$
        'LEN(FOLDERNAME_AUTO)
        FF_FORM4 = FOLDERNAME_AUTO
        FF_COUNT_FORM_CAPTURE4 = FF_COUNT_FORM_CAPTURE4 + 1
        If FS.FolderExists(FOLDERNAME_AUTO) = False Then
            I = CreateFolderTree(FOLDERNAME_AUTO)
            If I = False Then
                MSGBOX2 = "ERROR PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME_AUTO
                FRM_MSGBOX.Timer1.Enabled = True
                Exit Sub
            End If
        End If
    End If
End If
'NONE THESE HAPPEN SAME TIME

If YES_VAR = False Then Exit Sub

    If FS.FolderExists(FOLDERNAME_AUTO) = False Then
        I = CreateFolderTree(FOLDERNAME_AUTO)
        If I = False Then
            MSGBOX2 = "ERROR PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME_AUTO
            FRM_MSGBOX.Timer1.Enabled = True
            Exit Sub
        End If
    End If

If DO_FORM = True Then
    TT = PrintCurrentFormOntoForm(SCREEN_CAP)
    TS = FOLDERNAME_AUTO + "FormCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".jpg"
    'MsgBox TS
    On Error Resume Next
    SavePicture SCREEN_CAP.Picture, TS
    'ERR.NUMBER
    DONE_COUNT_FORM_CAPTURE = DONE_COUNT_FORM_CAPTURE + 1
End If

If DO_SCREEN = True Then
    TT = PrintScreenOntoForm(SCREEN_CAP)
    TS = FOLDERNAME_AUTO + "ScreenCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".jpg"
    On Error Resume Next
    SavePicture SCREEN_CAP.Picture, TS
    DONE_COUNT_FORM_CAPTURE = DONE_COUNT_FORM_CAPTURE + 1
End If



If Err.Number > 0 Then
    MSGBOX2 = "ERR.Description =" + Err.Description + vbCrLf + "ERR.Number =" + Str(Err.Number) + vbCrLf + "ERROR AT COMMAND :-" + vbCrLf + "SavePicture SCREEN_CAP.Picture, TS" + vbCrLf + "PICTURE IMAGE DID NOT SAVE" + vbCrLf + TS + vbCrLf + "Timer_SCREEN_SHOT -- WILL BE ENABLED=FALSE -- AFTER ERROR"
    FRM_MSGBOX.Timer1.Enabled = True
    Timer_SCREEN_SHOT.Enabled = False
    ERR_STAT = "ERROR -- "
Else
    ERR_STAT = ""
End If
    
    FOLDERNAME_AUTO = TS
    
    'Debug.Print TS
    MNU_TXT = Replace(TS, "D:\0 00 ART LOGGERS\", "__\")
    'MNU_TXT = Replace(MNU_TXT, GetComputerName, "...")
    'Debug.Print MNU_TXT
    
    Mnu_Explorer_Screen_Capture.Caption = ERR_STAT + "EXPLORE Auto Screen Image -- " + MNU_TXT
    Mnu_IVIEW_Screen_Capture.Caption = ERR_STAT + "IVIEW Auto Screen Image -- " + MNU_TXT
    
    Mnu_Explorer_Form_Capture.Caption = ERR_STAT + "EXPLORE Auto Screen -- " + MNU_TXT
    Mnu_IVIEW_Form_Capture.Caption = ERR_STAT + "IVIEW Auto Form IMAGE -- " + MNU_TXT
    
    'Explorer Auto Screen Capture
    'IVIEW - Auto Screen Capture
    'Mnu_IVIEW_Screen_Capture

Exit Sub

'------------------------------------------------------------



'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------


'OUT OF MEMORY ERROR
'ERR.NUMBER
'ERR.DESCRIPTION

'VAR TO USE AS FOLDER NAME
'FF1$ = "ScreenCapture_" + Format$(Now, "YYYY-MM-DD-DDD")
'FF2$ = "FormCapture_" + Format$(Now, "YYYY-MM-DD-DDD")
'
'On Error Resume Next
'FOLDERNAME1 = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\CLIP_Screen-Shots\" + FF1$
'FOLDERNAME_AUTO = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\CLIP_Form-Shots\" + FF2$
'
''ONLY USE FOLDERNAME_AUTO AT THE MOMENT
'If FS.FolderExists(FOLDERNAME_AUTO) = False Then
'    'Call SET_MAIN_FOLDER_ART_LOGGERS_APP_AND_SCREEN_SHOT
'    I = CreateFolderTree(FOLDERNAME_AUTO)
'    If I = False Then MsgBox "PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME_AUTO: Exit Sub
'End If

On Error GoTo 0

'TURN THE WHOLE SCREEN SHOT OFF NOW 30-11-2015
'ONLY FORM SHOT INSTEAD

'TT = PrintScreenOntoForm(SCREEN_CAP)
'TS = FOLDERNAME1 + "\ScreenCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".jpg"
'SavePicture SCREEN_CAP.Picture, TS

'TT = PrintCurrentFormOntoForm(SCREEN_CAP)
'TS = FOLDERNAME_AUTO + "\FormCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".jpg"
'On Error Resume Next
'SavePicture SCREEN_CAP.Picture, TS

'OUT OF MEMORY ERROR
'ERR.NUMBER
'ERR.DESCRIPTION

'If Err.Number > 0 Then MsgBox (Err.Description + vbCrLf + "ERR.Number =" + Str(Err.Number) + vbCrLf + "ERROR AT COMMAND :- SavePicture SCREEN_CAP.Picture, TS" + vbCrLf + "PICTURE IMAGE DID NOT SAVE" + vbCrLf + TS), vbMsgBoxSetForeground

'FileInUseDelay App.Path + "\VBDataNoArchive\Screen-Shot Pic Always.jpg"
'SavePicture Form1.Picture, App.Path + "\VBDataNoArchive\Screen-Shot Pic Always.jpg"

End Sub


