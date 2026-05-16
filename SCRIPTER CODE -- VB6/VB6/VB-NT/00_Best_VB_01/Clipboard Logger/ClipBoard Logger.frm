VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "ClipBoard Logger"
   ClientHeight    =   5205
   ClientLeft      =   3690
   ClientTop       =   6990
   ClientWidth     =   12750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ClipBoard Logger.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   12750
   WindowState     =   1  'Minimized
   Begin VB.Timer TIMER1_SPOT_CHECK_CLIPBOPARD_CORRECT_LENGHT 
      Interval        =   1000
      Left            =   4680
      Top             =   552
   End
   Begin VB.FileListBox File2 
      Height          =   324
      Left            =   3804
      TabIndex        =   16
      Top             =   60
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Timer TIMER_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4644
      Top             =   156
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9564
      Top             =   1584
   End
   Begin VB.Timer Timer_RESET_API_CLIPPER 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   9192
      Top             =   1584
   End
   Begin VB.Timer Timer_100_MILLISECOND 
      Interval        =   100
      Left            =   9192
      Top             =   1212
   End
   Begin VB.Timer Timer_MNU_COPY_2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9216
      Top             =   660
   End
   Begin VB.TextBox Txt_Search 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2724
      TabIndex        =   14
      Text            =   "Text Search"
      Top             =   504
      Width           =   1692
   End
   Begin VB.Timer Timer_GET_KEY_ESC_MINIMIZE 
      Interval        =   1
      Left            =   9048
      Top             =   140
   End
   Begin VB.Timer Timer1_PLUS 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8688
      Top             =   864
   End
   Begin VB.Timer TIMER_Clipboard_Set_Text 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8688
      Top             =   504
   End
   Begin VB.Timer TIMER_Clipboard_Clear 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8688
      Top             =   140
   End
   Begin VB.Timer Timer_SET_LABEL3_INFO_BAR 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8328
      Top             =   864
   End
   Begin VB.Timer Timer_1_SECOND 
      Interval        =   1000
      Left            =   8328
      Top             =   504
   End
   Begin VB.FileListBox File1 
      Height          =   576
      Left            =   10380
      TabIndex        =   13
      Top             =   2556
      Visible         =   0   'False
      Width           =   1080
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   900
      Left            =   36
      TabIndex        =   12
      Top             =   876
      Width           =   2592
      _ExtentX        =   4577
      _ExtentY        =   1588
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"ClipBoard Logger.frx":08CA
   End
   Begin RichTextLib.RichTextBox RTB_CLIPPER_FORMAT_CONVERTING 
      Height          =   792
      Left            =   48
      TabIndex        =   1
      Top             =   1812
      Visible         =   0   'False
      Width           =   4272
      _ExtentX        =   7541
      _ExtentY        =   1402
      _Version        =   393217
      TextRTF         =   $"ClipBoard Logger.frx":094C
   End
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   8328
      Top             =   140
   End
   Begin VB.Timer Timer_API_RESET_INFO 
      Interval        =   1000
      Left            =   5580
      Top             =   1536
   End
   Begin VB.Timer Timer_APP_BEGIN_TIMER 
      Interval        =   1000
      Left            =   7020
      Top             =   2244
   End
   Begin VB.Timer Timer_UNLOAD_FORM 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6660
      Top             =   2604
   End
   Begin VB.Timer Timer_KEYBOARD_Active 
      Interval        =   50
      Left            =   5580
      Top             =   2232
   End
   Begin VB.Timer TIMER_KEYBOARD_AND_MOUSE_ACTIVE 
      Enabled         =   0   'False
      Interval        =   59000
      Left            =   6660
      Top             =   492
   End
   Begin VB.FileListBox File3 
      Height          =   576
      Left            =   10380
      TabIndex        =   10
      Top             =   1848
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Timer TIMER_VB_PROJECT_DATE 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   6300
      Top             =   140
   End
   Begin VB.Timer TIMER_RETRY_WRITE_INFO_UNTIL_DONE1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6300
      Top             =   492
   End
   Begin VB.Timer TIMER_RETRY_WRITE_INFO_UNTIL_DONE2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6300
      Top             =   828
   End
   Begin VB.Timer Timer_MENU_HEIGHT_CHANGED 
      Interval        =   100
      Left            =   7020
      Top             =   1536
   End
   Begin VB.Timer Timer_MOUSE_5_MINUTE 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6660
      Top             =   2244
   End
   Begin VB.Timer Timer_Idle_Few_Second 
      Interval        =   1000
      Left            =   7020
      Top             =   1884
   End
   Begin VB.Timer Timer_MOUSE_4_MINUTE 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6660
      Top             =   1884
   End
   Begin VB.Timer Timer_MOUSE_3_MINUTE 
      Interval        =   59990
      Left            =   6660
      Top             =   1524
   End
   Begin VB.Timer TIMER_OutClipChunck_Txt 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5940
      Top             =   1188
   End
   Begin VB.Timer Timer_INFORAPID_MSGBOX 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5940
      Top             =   852
   End
   Begin VB.Timer Timer_MOUSE_2_MINUTE 
      Enabled         =   0   'False
      Interval        =   59990
      Left            =   6660
      Top             =   1164
   End
   Begin VB.Timer Timer_MOUSE_1_MINUTE 
      Enabled         =   0   'False
      Interval        =   59990
      Left            =   6660
      Top             =   828
   End
   Begin VB.Timer Timer_EXE_DAY_AGE 
      Interval        =   60000
      Left            =   5940
      Top             =   140
   End
   Begin VB.Timer Timer_API_OKAY_COLOUR 
      Interval        =   1
      Left            =   5580
      Top             =   1176
   End
   Begin VB.Timer Timer_API_Test 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5580
      Top             =   828
   End
   Begin VB.Timer Timer_WEATHER_URL 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7020
      Top             =   140
   End
   Begin VB.DirListBox Dir1 
      Height          =   624
      Left            =   10380
      TabIndex        =   7
      Top             =   1116
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Timer Timer_EXPLORER_GONE 
      Interval        =   1000
      Left            =   5580
      Top             =   140
   End
   Begin VB.Timer Timer_MOUSE_KEYBOARD 
      Interval        =   1000
      Left            =   5580
      Top             =   1884
   End
   Begin VB.Timer Timer_SCREEN_SHOT_AUTO 
      Interval        =   100
      Left            =   5580
      Top             =   492
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7020
      Top             =   492
   End
   Begin VB.Timer Timer_MOUSE 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6660
      Top             =   140
   End
   Begin VB.Timer Timer_FORM_RESIZE 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   5940
      Top             =   516
   End
   Begin VB.Timer TimerSCROLL 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   7020
      Top             =   828
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5208
      Top             =   492
   End
   Begin VB.Timer TIMER_JOYPAD 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   7020
      Top             =   1188
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   400
      Left            =   10980
      ScaleHeight     =   345
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   400
      Left            =   10392
      ScaleHeight     =   345
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   588
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture2 
      Height          =   400
      Left            =   10980
      ScaleHeight     =   345
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   140
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Height          =   400
      Left            =   10392
      ScaleHeight     =   345
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   140
      Visible         =   0   'False
      Width           =   540
   End
   Begin MCI.MMControl MMControl1 
      Height          =   336
      Left            =   7476
      TabIndex        =   0
      Top             =   2232
      Visible         =   0   'False
      Width           =   2364
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5208
      Top             =   140
   End
   Begin MCI.MMControl MMControl2 
      Height          =   336
      Left            =   7476
      TabIndex        =   6
      Top             =   2604
      Visible         =   0   'False
      Width           =   2364
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   576
      Left            =   2940
      TabIndex        =   15
      Top             =   1092
      Width           =   2328
   End
   Begin VB.Label Label3 
      Caption         =   "Label3 -- INFO"
      Height          =   300
      Left            =   60
      TabIndex        =   11
      Top             =   492
      Width           =   2568
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "TIMERS NOT USED >>"
      Height          =   1788
      Left            =   7476
      TabIndex        =   9
      Top             =   140
      Visible         =   0   'False
      Width           =   804
   End
   Begin VB.Label Label1 
      Caption         =   "Label1 -- COLOUR BAR"
      Height          =   252
      Left            =   48
      TabIndex        =   8
      Top             =   140
      Width           =   2568
   End
   Begin VB.Menu Mnu_Exit 
      Caption         =   "Exit"
   End
   Begin VB.Menu Mnu_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_EXE_NEWER_IN_MIRROR 
      Caption         =   "**** VB EXE NEWER IN MIRROR ****"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_ALWAYS_ON_TOP 
      Caption         =   "ON TOP"
   End
   Begin VB.Menu MNU_AUDIO_WANT_ON 
      Caption         =   "AUDIO WANT ON"
   End
   Begin VB.Menu MNU_API_RESET 
      Caption         =   "API"
   End
   Begin VB.Menu MNU_RESET_FORM 
      Caption         =   "RESET FORM"
   End
   Begin VB.Menu MNU_SCROLL_BOTTOM_OFF__WANT_ON_ASK_PRESS 
      Caption         =   "SCROLL BOTTOM OFF -- WANT ON ASK PRESS"
   End
   Begin VB.Menu MNU_SELECTOR_WHOLE_LINE_MODE 
      Caption         =   "SELECTOR WHOLE LINE MODE=FALSE"
   End
   Begin VB.Menu MNU_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER 
      Caption         =   "GATHER OTHER COMPUTER"
   End
   Begin VB.Menu MNU_AUTO_MINIMIZE_OFF_ON 
      Caption         =   "AUTO MINIMIZE=OFF"
   End
   Begin VB.Menu AUTO_CLIPBOARD_SELECTOR 
      Caption         =   "AUTO_CLIPBOARD_SELECTOR_>2"
   End
   Begin VB.Menu MNU_EXE_LANUCHER_2 
      Caption         =   "EXE MENU"
      Begin VB.Menu MNU_EXE_LANUCHER 
         Caption         =   "TREE SIZE FREE"
         Index           =   1
      End
      Begin VB.Menu MNU_EXE_LANUCHER 
         Caption         =   "TREESIZE PRO"
         Index           =   2
      End
      Begin VB.Menu MNU_EXE_LANUCHER 
         Caption         =   "TREESIZE SEARCH"
         Index           =   3
      End
      Begin VB.Menu MNU_EXE_LANUCHER 
         Caption         =   "FLICKR YAHOO"
         Index           =   4
      End
      Begin VB.Menu MNU_EXE_LANUCHER 
         Caption         =   "FLCIKR SYNC"
         Index           =   5
      End
      Begin VB.Menu MNU_EXE_LANUCHER 
         Caption         =   "VB MULTI MENU"
         Index           =   6
      End
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
      Begin VB.Menu MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU 
         Caption         =   "CODE_MENU"
         Index           =   1
      End
      Begin VB.Menu MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU 
         Caption         =   "EXPLORER -- ME_VB"
         Index           =   2
      End
      Begin VB.Menu MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU 
         Caption         =   "EXPLORER -- SEND TO SYSTEM FOLDER"
         Index           =   3
      End
      Begin VB.Menu MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU 
         Caption         =   "EXPLORER -- SEND TO FAT 32 FOLDER"
         Index           =   4
      End
      Begin VB.Menu MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU 
         Caption         =   "EXPLORER -- PIN TO START MENU - FOLDER"
         Index           =   5
      End
      Begin VB.Menu MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU 
         Caption         =   "EXPLORER -- DESKTOP FOLDER - PUBLIC"
         Index           =   6
      End
      Begin VB.Menu MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU 
         Caption         =   "EXPLORER -- DESKTOP FOLDER - USER"
         Index           =   7
      End
      Begin VB.Menu MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU 
         Caption         =   "MIRROR COPY CUSTOM SEND_TO FOLDER TO ANY OPERATING SYSTEM SEND_TO SPECIAL FOLDER"
         Index           =   8
      End
      Begin VB.Menu MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU 
         Caption         =   "ABORT SHUTDOWN \system32\shutdown.exe /a"
         Index           =   9
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
      Begin VB.Menu MNU_SOUND_OPTION 
         Caption         =   "MNU_SOUND_OPTION"
         Index           =   7
      End
      Begin VB.Menu MNU_SOUND_OPTION 
         Caption         =   "MNU_SOUND_OPTION"
         Index           =   8
      End
      Begin VB.Menu MNU_SOUND_OPTION 
         Caption         =   "MNU_SOUND_OPTION"
         Index           =   9
      End
      Begin VB.Menu MNU_SOUND_OPTION 
         Caption         =   "MNU_SOUND_OPTION"
         Index           =   10
      End
      Begin VB.Menu MNU_SOUND_OPTION 
         Caption         =   "MNU_SOUND_OPTION"
         Index           =   11
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
      Caption         =   "__* LOGG MENU *__"
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "Open Logg Explorer"
         Index           =   1
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "INFO RAPID ALL FOLDERS"
         Index           =   2
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "INFO RAPID ALL FOLDERS AND SMALL FILES BEGIN ClipBoard-*.TXT"
         Index           =   3
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "INFO RAPID MY USER"
         Index           =   4
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "INFO RAPID TRIM LOGG"
         Index           =   5
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "Hex Open Recent Trim Logg"
         Index           =   6
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "----"
         Index           =   7
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "Edit Recent Trim Logg"
         Index           =   8
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "Edit This Logg"
         Index           =   9
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "Edit Total Logg"
         Index           =   10
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "Edit Strip Logg"
         Index           =   11
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "----"
         Index           =   12
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "TEXTVIEW RECENT TRIM LOGG"
         Index           =   13
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "TEXTVIEW THIS LOGG"
         Index           =   14
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "TEXTVIEW TOTAL LOGG"
         Index           =   15
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "TEXTVIEW STRIPER"
         Index           =   16
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "----"
         Index           =   17
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "Shell T -- This Logg"
         Index           =   18
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   "Shell T -- Total Logg"
         Index           =   19
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   20
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   21
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   22
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   23
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   24
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   25
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   26
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   27
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   28
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   29
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   30
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   31
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   32
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   33
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   34
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   35
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   36
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   37
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   38
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   39
      End
      Begin VB.Menu MNU_LOGGEXPLORER 
         Caption         =   ""
         Index           =   40
      End
   End
   Begin VB.Menu MNU_LAST_ART_PIC2 
      Caption         =   "* IMAGE MENU *"
      Begin VB.Menu MNU_SCREEN_SHOT 
         Caption         =   "OPEN ART LOGG FOLDER"
      End
      Begin VB.Menu MNU_FOLDER_AUTO_CLIPBOARD 
         Caption         =   "FOLDER AUTO CLIPBOARD"
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
   Begin VB.Menu MNU_TEXT_OPEN_RECENT_2 
      Caption         =   "TEXTVIEW TRIM LOGG"
   End
   Begin VB.Menu MNU_TEXT_OPEN_STRIP_LOGGER_2 
      Caption         =   "TEXTVIEW STRIP LOGG"
   End
   Begin VB.Menu MNU_SHOW_IMAGE_1 
      Caption         =   "SHOW IMAGE 01 of 03 LAST OR FIND"
   End
   Begin VB.Menu MNU_SHOW_IMAGE_2_GLOBAL 
      Caption         =   "SHOW IMAGE 02 of 03 GLOBAL"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_SHOW_IMAGE_3_COMMON 
      Caption         =   "SHOW IMAGE 03 of 03 MINE COMMON FOLDER"
   End
   Begin VB.Menu MNU_CLIPBOARD_IMAGE_FILENAME 
      Caption         =   "CLIPBOARD IMAGE FILENAME"
   End
   Begin VB.Menu MNU_EXPLORER_IMAGE 
      Caption         =   "EXPLORER IMAGE"
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
         Caption         =   "LAST GRAB - ALL CAPS - ON ALL CLIPBOARDER IF ENABLED"
      End
      Begin VB.Menu MNU_LAST_GRAB_CAPS 
         Caption         =   "LAST GRAB - CAPS - ONCE"
      End
      Begin VB.Menu MNU_LAST_GRAB_PRO_CAPS 
         Caption         =   "LAST GRAB - PRO - CAPS - ON ALL CLIPBOARDER IF ENABLED"
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
         Caption         =   "CLIPBOARD A SPACE CHARACTOR"
      End
      Begin VB.Menu MNU_REFORMAT_ADD_A_DASH 
         Caption         =   "ADD DOUBLE DASH BEFORE EVERY CR-LINEFEED - TEXTBOXES THE REMOVE CR-LINEFEED  PROBLEM"
      End
      Begin VB.Menu MNU_REFORMAT_REMOVE_THE_DASH 
         Caption         =   "REMOVE DOUBLE DASH BEFORE EVERY CR-LINEFEED"
      End
      Begin VB.Menu mnu_Keyboard_Logger_Remove_Doubel_Char_Into_One 
         Caption         =   "REPLACE DOUBLE CHAR INTO ONE -- KEYBOARD LOGGER"
      End
      Begin VB.Menu MNU_REPLACE_ALL_STRING_CHAR_TO_UNDERLINE 
         Caption         =   "REPLACE ALL STRING CHAR TO UNDERLINE"
      End
      Begin VB.Menu MNU_REPLACE_CODER_STRING_UNDERLINE_AND_SPECIAL_CHAR 
         Caption         =   "REPLACE CODER STRING UNDERLINE && SPECIAL CHAR"
      End
      Begin VB.Menu MNU_REPLACE_DASH_SPACE_TO_UNDERLINE 
         Caption         =   "REPLACE DASH && SPACE TO UNDERLINE"
      End
      Begin VB.Menu MNU_REFORMAT_REPLACE_DASH_TO_UNDERLINE 
         Caption         =   "REPLACE DASH TO UNDERLINE"
      End
      Begin VB.Menu MNU_REFORMAT_REPLACE_SPACE_FOR_DASH 
         Caption         =   "REPLACE SPACE TO DASH"
      End
      Begin VB.Menu MNU_REFORMAT_REPLACE_SPACE_TO_NOTHING 
         Caption         =   "REPLACE SPACE TO NOTHING"
      End
      Begin VB.Menu MNU_REFORMAT_REPLACE_SPACE_TO_UNDERLINE 
         Caption         =   "REPLACE SPACE TO UNDERLINE"
      End
      Begin VB.Menu MNU_REFORMAT_REPLACE_CRLF_DOUBLE_TO_SINGLE 
         Caption         =   "REPLACE TRIPLE CRLF TO DOUBLE"
      End
      Begin VB.Menu MNU_REFORMAT_REPLACE_CRLF_DOUBLE_TO_SINGLE_02 
         Caption         =   "REPLACE DOUBLE CRLF TO SINGLE"
      End
      Begin VB.Menu MNU_REFORMAT_REPLACE_TAB_TO_4_SPACE 
         Caption         =   "REPLACE TAB TO 4 SPACES -- CODER"
      End
   End
   Begin VB.Menu MNU_SPACE_DASH_TO_UNDERSCORE 
      Caption         =   "TXT SPACE TO DASH"
   End
   Begin VB.Menu MNU_TEXT_CAPITAL_MODE_ON 
      Caption         =   "TEXT_CAPITAL_MODE__OFF"
   End
   Begin VB.Menu MNU_CALC_MODE_OFF_AND_ON 
      Caption         =   "CALC MODE=ON"
   End
   Begin VB.Menu MNU_CALC_GIVE_RESULT 
      Caption         =   "CALC TOTAL"
   End
   Begin VB.Menu MNU_CALC_RESULT_AND_SUM 
      Caption         =   "CALC SUM && TOTAL"
   End
   Begin VB.Menu MNU_CALC_POUND_ALL_CHUNK 
      Caption         =   "CALC £ ALL CHUNK"
   End
   Begin VB.Menu MNU_LAST_GRAB_LOW_CASE_TOP_BAR 
      Caption         =   "TXT LOW"
   End
   Begin VB.Menu MNU_LAST_GRAB_CAPS_TOP_BAR 
      Caption         =   "TXT CAPS"
   End
   Begin VB.Menu MNU_LAST_GRAB_TEXT_INVERT 
      Caption         =   "TXT INVERT"
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
   Begin VB.Menu Mnu_Clipboard_Info_Spacer 
      Caption         =   "ClipBoard INFO"
   End
   Begin VB.Menu Mnu_LAST_CLIP_TIME 
      Caption         =   "Last Clip Time"
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
      Begin VB.Menu Mnu_URL_Browser 
         Caption         =   "URL Launch In Browser"
      End
      Begin VB.Menu MNU_REG_JUMP 
         Caption         =   "REG JUMP"
      End
      Begin VB.Menu MNU_CPC 
         Caption         =   "URL CPC WEB Page Offer Search 100 Web Pages 100%"
      End
      Begin VB.Menu Mnu_HTML_URL 
         Caption         =   "URL ADD WRAPPER  HTML <a href=""""..."
      End
   End
   Begin VB.Menu MNU_CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY 
      Caption         =   "URL_WEB_FOLDER_FILE_REG_KEY_CPC"
   End
   Begin VB.Menu MNU_CPC_TOP_MENU 
      Caption         =   "* CPC 00 40 80 90 *"
   End
   Begin VB.Menu MNU_404_CPC_PAGES 
      Caption         =   "* CPC 404 PAGE REMOVE *"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_READ_CLIPBOARD_OF_HTTP_WEB_CHUNK_AND_OPEN 
      Caption         =   "READ CLIPBOARD OF HTTP WEB CHUNK AND OPEN"
   End
   Begin VB.Menu MNU_EBAY_MULTIPLE_PAGE 
      Caption         =   "EBAY MULTIPLE PAGE"
   End
   Begin VB.Menu MNU_EBAY_YEAR_1 
      Caption         =   "EBAY 2019"
   End
   Begin VB.Menu MNU_EBAY_10 
      Caption         =   "EBAY 10"
   End
   Begin VB.Menu MNU_EBAY_YEAR_2 
      Caption         =   "EBAY 2018"
   End
   Begin VB.Menu MNU_EBAY_YEAR_1_10 
      Caption         =   "EBAY 2019 10"
   End
   Begin VB.Menu MNU_EBAY_YEAR_3 
      Caption         =   "EBAY 2017"
   End
   Begin VB.Menu MNU_EBAY_MULTI_MENU_PULL 
      Caption         =   "EBAY MULTI MENU PULL"
      Begin VB.Menu MNU_EBAY_MULTI 
         Caption         =   "LOADER 2019 2018 2017"
         Index           =   1
      End
      Begin VB.Menu MNU_EBAY_MULTI 
         Caption         =   "LOADER 2019"
         Index           =   2
      End
      Begin VB.Menu MNU_EBAY_MULTI 
         Caption         =   "LOADER 2019 2018"
         Index           =   3
      End
      Begin VB.Menu MNU_EBAY_MULTI 
         Caption         =   "LOADER 2018"
         Index           =   4
      End
      Begin VB.Menu MNU_EBAY_MULTI 
         Caption         =   "LOADER 2018 2017"
         Index           =   5
      End
      Begin VB.Menu MNU_EBAY_MULTI 
         Caption         =   "LOADER 2017"
         Index           =   6
      End
   End
   Begin VB.Menu MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS 
      Caption         =   "*--- NOT SCROLL IN  OPTION MENU---*"
   End
   Begin VB.Menu Mnu_Show_Error_Script_Frm_Msgbox 
      Caption         =   "** Show Error Script **"
   End
   Begin VB.Menu MNU_ERROR_INFO 
      Caption         =   "ANY ERROR MSG OR INFO"
   End
   Begin VB.Menu Mnu_Audio_ON 
      Caption         =   "** /\/\ Audio Is Off /\/\ ** Hitt On Here for ON ***"
   End
   Begin VB.Menu MNU_SELECTED 
      Caption         =   "SELECTED"
   End
   Begin VB.Menu MNU_CLIPPER_REMOVE_DOUBLE_VBCRLF 
      Caption         =   "REMOVE DOUBLE VBCRLF FROM CLIPPER"
   End
   Begin VB.Menu MNU_CLIPPER_WRAPPER_HTML_LINK_HREF 
      Caption         =   "WRAP ALL BUNCH HTML LINK IN A WRAPPER <HREF>"
   End
   Begin VB.Menu Mnu_GET_LAST_CLIPPER_IN_TINYURL 
      Caption         =   "____ LAST CLIP TINYURL ____"
   End
   Begin VB.Menu MNU_COPY_2_CLIPPER_BEFORE 
      Caption         =   "COPY 2 CLIPPER BEFORE"
   End
   Begin VB.Menu MNU_COPY 
      Caption         =   "<*** COPY ***>"
   End
   Begin VB.Menu MNU_COPY_2 
      Caption         =   "< COPY >"
   End
   Begin VB.Menu MNU_EMPTY_CLIPBOARD 
      Caption         =   "EMPTY"
   End
   Begin VB.Menu Mnu_Menu_Item_Count 
      Caption         =   "Menu Counter"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' WORK ON HERE DISABLED THE LOT
' Timer_SCREEN_SHOT_AUTO_Timer
' WORK ON HERE HAD TO DISABLE THE LOT SEARCH
    ' TIMER_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER.Enabled = False
    ' Exit Sub


' -------------------------------------------------------------------------------------------------
' MORE CODE TO DO MAKE EFFECIENT -- Sat 14-Sep-2019 20:42:40
' -------------------------------------------------------------------------------------------------
Dim AlwaysOnTop_MODE
Dim STARTUP_RUN

Const WM_COPY = &H301
' --------------------------------------------------------------------
' TO HERE -- Const WM_COPY = &H301
' LINE GRAB MODE AND PASTE IT
' LIKE -- CLIPBOARD.SETTEXT
' SendMessage Text1.hWnd, WM_COPY, 0&, 0& 'Copy
' --------------------------------------------------------------------

Dim API_LOAD


Dim MISSING_PULSE_PULSE_TRAIN
Dim TIME_STAMP_CLIPBOARD
Dim OLD_CLIP_INFO_04

Dim AR_DAY_NOW()
Dim OLD_DAY_NOW
Dim FOLDERNAME_AUTO_SWAP
Dim FOLDERNAME_AUTO_MOUSE_ACTIVE
Dim FOLDERNAME_AUTO_MINUTE
Dim FOLDERNAME_AUTO_MINUTE_LONG

' -------------------------------------------------------------------------------------------------
Dim OLD_MNU_SOUND_WAV_1_CHECKER_INDEX
Dim OLD_MNU_SOUND_WAV_2_CHECKER_INDEX

Dim MMCONTROL1__FILENAME_PARTIAL
Dim RUN_ONCE_AT_START
Dim MMCONTROL1__FILENAME
Dim MMCONTROL2__FILENAME
' Tue 02-Jun-2020 18:36:32
' -------------------------------------------------------------------------------------------------

' -------------------------------------------------------------------------------------------------
' MORE WORKER TO DO HERE --
' FROM -- Sat 14-Sep-2019 21:18:00
' -------------------------------------------------------------------------------------------------
Dim VALUE_START_STOP
Dim ONCE_RUN_ALL_FILE_DATE_POPULATE__VARIABLE_NOT_DO_AGAIN__
Dim ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_DATE(100)
Dim ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_SIZE_LEN(100)
Dim ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_CRC32(100)
Dim ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_NAME(100)

Dim MNU_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_VALUE

Dim MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER



' -------------------------------------------------------------------------------------------------
' WORK ON CLIPBOARD LOGGER
' -------------------------------------------------------------------------------------------------
' Fri 27-Apr-2018 04:19:23
' Fri 27-Apr-2018 08:18:00 4 HOURS THROUGH NIGHT DUSK AND DAY
' -------------------------------------------------------------------------------------------------
' 1. DO THE CALC WORKING PROPER ADD BUTTON TO GRAB LAST SUM AND TOTAL OR
' JUST THE TOTAL IF CALC IN OFF MODE -- CHANGED TO STARTUP IN OFF MODE
' 2. TIME SINCE CLIPBOARD PROPER
' 3. SMOOTHER OPERATION RESIZE FORM AND INNER TEXT BOX RESIZE
' NOW USING POINTXY UNDER MOUSE IF IN FOREGROUND T ADJUST INNER BOX
' WAS A BUG WITH INNERBOX NOT SIZING BUT FOUND THAT BUG AND
' PUT EXTRA IN WITH POINTXY AND SPEEDIER TIMER RESIZE
' STILL POSIBLE TO SEE THE BOX HAS MINOR GAPS RE-AJUSTING
' WITH FULL SPEED TIMER
' 4. RELOAD FORM DONE BY ANOTHER FORM CAN'T DO WITHIN FORM
' TO HELP WITH THE API NEED RESETING OPTION REQUIRE TO START
' SILENTLY PREHAPS TO DO AUTO IF ERROR FOUND MAYBE WHEN CONTROL C
' PRESS AND NOTHING HAPPENS CHECK ERROR RESET
' -------------------------------------------------------------------------------------------------

' -------------------------------------------------------------------------------------------------
' WORK ON CLIPBOARD LOGGER
' -------------------------------------------------------------------------------------------------
' 1..
' ADD CODE __ MENU'S NOW HAVE BRAKETS AROUND THEM BUT EXTRA LEARN DONN'T HAVE IT ON SUB MENU ITEM
' YET TO BE UPGRADED IN ELITESPY+ WHERE GOT MOST WORK FROM IN FIRST PLACE
' 2..
' ADD CODE __ NOW HAS A SEARCH BOX - WHICH QUITE DIFFERCULT AS MAIN FOCUS CONTROL IN FORM IS THE RICH TEXT BOX
' SO KEYS HAVE TO BE PASSED OVER TO SEARCH BOX WHENEVER TYPING KEYS
' AND CONTROL KEY FILTERING A DELEETE AND BACK SPACE HAS TO LEARN WHERE CURSOR IS AND DEL FORWARD OR
' BACK FROM THERE
' TAB FOR FOCUS
' TEXT ARE TO DISPLAY SEACH BOX AT FIRST BUT REMOVED SOON AS ENTER WITH TYPING OR THE STARTUP IN
' FOCUS IS LOST FOCUSING
' LONG TIME AND TODAY SPEND AN HOUR ON CODE 1ST OF ALL AND THEN 4 HOURS ON A CODE
'
' -------------------------------------------------------------------------------------------------
' Tue 23-Oct-2018 15:49:55
' Tue 23-Oct-2018 21:30:00 NEALY SIX HOUR
' -------------------------------------------------------------------------------------------------


' -------------------------------------------------------------------------------------------------
' WORK ON CLIPBOARD LOGGER
' -------------------------------------------------------------------------------------------------
' 1..
' MAKE CLIPBOARD USE TWO CLIPS TO MERGE GET TINYURL WITH TITLE
' 2..
' MAKE CLIPBOARD SIZE SHOW AT INFO BAR TOP
' 3..
' MAKE CPC MENU BUTTON SHOW IF URL NOT ABLE TO PROCESSOR
' AND TIDY PRESENTATION
'
' -------------------------------------------------------------------------------------------------
' Thu 14-Mar-2019 09:32:27
' Thu 14-Mar-2019 10:48:00
' -------------------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------------------
' THIS WILL STOP THE PROBLEM OF SELECT SOME TEXT AND THE WINDOWS RESIZE CHANGE SO SELECTED TEXT GONE
' THAT ALSO HAPPEN WHEN SOME MENU ITEM CHANGE LENGHT
' BUT BETTER STILL NEXT TRY NOT RESIZE ANY WHILE DRAG HAPPEN
' --------------------------------------------------------------------------------------------------


' -----------------------------------------------------------------------------------------------
' Text1.Visible = False
' Label4.Caption = Str(X) + " -- " + Str(Y)
' WANT TO TRY AND GET IF USER IS RESIZE THE FORM
' NONE SEARCHER
' LOOK AT MOUSE CHANGE SHAPE
' THE CORDINATE BOTTOM LEFT
' BUT WIN 10 HAS SIDE OF FORM
' AND THEN BE DETERMINE IF WANT OT RESIZE MANY INTERNAL ITEM
' AS NOT WANT NORMAL WHEN USER SELECT TO NOT JUMP TO BOTTOM AT ANYTIME LIKE WHEN CLIPBOARD ARRIVE
' 2019 JUNE 09
' -----------------------------------------------------------------------------------------------


' -----------------------------------------------------------------------------------------------
' -----------------------------------------------------------------------------------------------
' SESSION 010
' -----------------------------------------------------------------------------------------------
' DO WORK -- AND MULTI PROJECT
' -----------------------------------------------------------------------------------------------
' SEARCH -- STOP SHIFT AND END KEY COPY PASTE KEYBOARD FORM HAVE ERROR BY ENTER SEARCH
' -----------------------------------------------------------------------------------------------
' WORK LEFT TO DO
' Project_Check_Date
' -----------------------------------------------------------------------------------------------
' ADD THE CALL
' Public Sub TIMER_CLIPBOARD_TIMER_RETRY_Timer()
' NOW ABLE TO USE TIMER WITH CLIPBOARD
' AND NOT LOOP EVER
' IF CLIPBOARD FAIL HAS COOMAND TO EXIT PROCEDURE AND THEN
' TIMER CALL BACK
' CallByName Form, VAR_TIMER_CLIPBOARD_TIMER_RETRY, VbMethod
' ----------------------------------------------------------------------
' Fri 24-Jan-2020 02:00:00
' Fri 24-Jan-2020 05:54:00 -- 4 HOUR 54 MINUTE
' ----------------------------------------------------------------------


' --------------------------------------------------------------------
' TIDY UP HERE WORK
' SETUP_SOUND_FILE(VARSOUND)
' AND AT ROUTINE --
' Private Sub MNU_SOUND_OPTION_Click(Index As Integer)
' --------------------------------------------------------------------
' Tue 02-Jun-2020 21:24:05
' Wed 03-Jun-2020 01:01:16 -- 3 HOUR 37 MINUTE
' --------------------------------------------------------------------




Dim MIDNIGHT_TIMER As Date

' FOR USE CHANGE TO UPPER CASE EVERY CLIPBOARD IN
' -----------------------------------------------
Dim O_2_AD

Dim EBAY_YEAR_IN

Dim FILE_NAME_USER As String


Dim CLIPBOARD_RETURN_TIMER_ERROR_1
Dim CLIPBOARD_RETURN_TIMER_ERROR_2
Dim CLIPBOARD_RETURN_TIMER_ERROR_3
Dim CLIPBOARD_RETURN_TIMER_ERROR_4
Dim CLIPBOARD_RETURN_TIMER_ERROR_5
Dim CLIPBOARD_RETURN_TIMER_ERROR_6
Dim CLIPBOARD_RETURN_TIMER_ERROR_7
Dim CLIPBOARD_RETURN_TIMER_ERROR_8
Dim CLIPBOARD_RETURN_TIMER_ERROR_9


Dim Clipboard_GetFormat_vbCFText_OO_1
Dim Clipboard_GetFormat_vbCFBitmap_OO_2


Dim O_GMCL_I

Dim O_Menu_Height_1
'----
Dim STOP_RESIZE_WHILE_MOUSE_OR_KEY_DOWN


Dim FORM_LOAD_SIZE_NOT_SET_YET

Dim CLIPBOARD_SIZE_VAR

Dim RecursiveSearch_Most_Recent_STORE_DATE

Dim CALC_SET_GO

Const DontWaitUntilFinished = False, WSCRIPT_ShowWindow = 1, WSCRIPT_DontShowWindow = 0, WaitUntilFinished = True

Dim X_COUNT_VAR
Dim O_FLAG_WW_VAR
' -------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------
'DO THE TEAMVIEWER CLIPBOARD SOUND EFFECTING PROBLEM
'-------------------------------------------------------------------------------------------
Dim DONT_PLAY_SOUND_TEAM_VIEWER
Dim OLD_FOREGROUND_WINDOW_TEAM_VIEWER
'-------------------------------------------------------------------------------------------


Dim MMControl2_Counting_Var
Dim O_MMControl2_GFW_Var

'WORK HERE T0 AUTOUPDATE SCRIPT

Dim GG$

Dim CALC_NEW_INPUT
Dim CALC_NEW_INPUT_MNU
Dim Clipboard_Clear_SUPPRESS_MESSAGE
Dim ENTER_TEXT_IN_LOGGER_FOREGROUND_OVERRIDE
Public TIMER_Clipboard_Clear_VAR
Public TIMER_Clipboard_Set_Text_VAR
Public Clipboard_Set_Text_VAR

Public EXIT_TRUE

'mcontrol

Dim SET_LABEL3_INFO_BAR_READY_TO_GO_VAR

Dim CALC_ARRAY_VAR() As String

Dim O2_CLIPPER_CALC
Dim CALC_STRINGER_S As String
Dim O_CALC_STRINGER_S As String
Public CALC_MODE


'SEARCHER
'STILL NEEDS WORK
Dim VAL_MINUTE_API
Dim Test_Min_Var
Dim OLD_VAR_TRIPPING

Dim PATH_CHECK As String

Dim ARRAY_ICON_NOT_WANT() As String
Dim ARRAY_ICON_DE_DUPE_VAR_() As String
Dim ARRAY_ICON_DE_DUPE_FILE() As String
Public RecursiveSearch_Most_Recent_VAR
Public TT_VAR_API_RESET
Public TT_VAR_LAST_CLIP_TIME_Second
Public TT_VAR_RESULT_MNU_RUN_TIME

Dim A_TimeIdle
Dim DIE()


Dim O_Text1_SelLength

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

'DO THIS -- MNU_INFO_RAPID_ALL_Click()
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

Public O_VAL_MINUTE_API
Dim oMNU_IDLE_ACTIVE

Dim COMPARE_EXE_1, COMPARE_EXE_2, COMPARE_EXE_3


Dim VAR_MNU_404_HITT_ONCE

Dim GO3

Dim I5()


Dim TEXTBOX1_SELECTION_CHANGE


Dim RCPC
Dim R_EBAY
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

Dim FOLDERNAME_AUTO '' IMAGE
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

Dim OLTLWH_1
Dim OLTLWH_2
Dim OLTLWH_3

Dim RESET_RRR
Dim DUPE_IMAGE_AT_LOAD_FORM
'Dim RESIZE_AT_LOAD, RESIZE_AT_LOAD2

Public ART$, ART2$, GGSize

Public QQ2$

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


Dim OFH, A1 As String, B1 As String, F1 As String

Private Declare Function FlatSB_SetScrollPos Lib "comctl32" (ByVal hWnd As Long, ByVal code As Long, ByVal nPos As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32" (ByVal hWnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO) As Boolean
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32" (ByVal hWnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO, ByVal fRedraw As Boolean) As Long

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
        Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

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
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Private Const GW_HWNDNEXT = 2
'Private Const GW_CHILD = 5
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long



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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
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

Dim SEARCH_F3

Dim SEARCH_BOX_NEVER_BEFORE_FOCUS


Dim VAR_TIMER_CLIPBOARD_TIMER_RETRY

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CYCAPTION = 4
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

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
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
'Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long


Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long



Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long  'MODULE 1135

Const hWnd_TOPMOST = -1
Const hWnd_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

'Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hWnd As Long)
    Dim flags
    SetWindowPos hWnd, hWnd_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function



' Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

'------------------------------------------------

Property Get TitleBarHeight() As Long
    TitleBarHeight = GetSystemMetrics(SM_CYCAPTION)
End Property




Public Sub TIMER_CLIPBOARD_TIMER_RETRY_Timer()
    ' ---------------------------------------
    ' CLIPBOARD_TIMER
    ' MAKE ALL CALLBACK CODE PUBLIC
    ' ---------------------------------------
    Dim Form As Form
    For Each Form In Forms
       If Form.Name = Me.Name Then
            CallByName Form, VAR_TIMER_CLIPBOARD_TIMER_RETRY, VbMethod
        End If
    Next
    TIMER_CLIPBOARD_TIMER_RETRY.Enabled = False
End Sub


Sub HEIGHT_BAR_TOP_ROUTINE()

    If Txt_Search.Left < 100 Then Txt_Search.Left = 100

    HEIGHT_BAR_TOP = 550
    Txt_Search.Width = 4000
    ' FOR XP THING LOW END COMPUTER SCREEN SIZE
    ' -----------------------------------------
    ' HIGHEND   -- 23040
    ' LOWER END -- 15360
    If Screen.Width < 18000 Then
        Label3.FontSize = 8.2
        HEIGHT_BAR_TOP = 450
        Txt_Search.Width = 2400
    Else
        HEIGHT_BAR_TOP = 550
        Txt_Search.Width = 4000
    End If
    
    Txt_Search.Top = Label1.Top + Label1.Height
    Txt_Search.Left = Me.Width - Txt_Search.Width
    Txt_Search.Height = HEIGHT_BAR_TOP '340 ' 336 IS MINIMUM HEIGHT 1 LINE ADD A FEW FOR NICER NUMBER
    Txt_Search.BackColor = RGB(255, 255, 255)
    
    
    Label3.Left = 0
    Label3.Top = Label1.Top + Label1.Height
    Label3.Height = HEIGHT_BAR_TOP
    If Txt_Search.Left > 0 Then
        Label3.Width = Txt_Search.Left - 40
    End If
    Label3.BackColor = RGB(255, 255, 255)
    Label3.FontSize = Label3.FontSize
    

End Sub

Private Sub Form_Load()

STARTUP_RUN = True

LATCH_RUN_ONCE = False
OLTLWH_1 = 0
OLTLWH_3 = 0

SEARCH_BOX_NEVER_BEFORE_FOCUS = True

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
If App.PrevInstance = True Then End

Set FSO = CreateObject("Scripting.FileSystemObject")

Debug.Print "FORM_LOAD"


Dim reg_valuename, WShell, Cmd, cmdLine, I
GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv").EnumValues &H80000003, "S-1-5-19\Environment", reg_valuename
If IsArray(reg_valuename) <> 0 Then
    RequireAdmin = 1
End If

If IsIDE = False Then
    If RequireAdmin = 0 Then
        I_TEXT = I_TEXT + "Set UAC = CreateObject(""SHELL.APPLICATION"")" + vbCrLf
        I_TEXT = I_TEXT + "UAC.ShellExecute """ + App.Path + "\" + App.EXEName + """, """", """", ""RUNAS"", 1" + vbCrLf
                
        VBSCRIPT_PATH = App.Path + "\RELOAD WITH ADMINISTRATOR RIGHTS.VBS"
        FR1 = FreeFile
        Open VBSCRIPT_PATH For Output As #FR1
        Print #FR1, I_TEXT
        Close #FR1
        
        Dim WSHShell
        Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.Run """" + VBSCRIPT_PATH + """"
        Set WSHShell = Nothing
        'Unload Me
        End
    End If
End If
If IsIDE = True Then
    If RequireAdmin <> 1 Then
        MsgBox "Visual Basic IDE Not in Administrator Rights Mode"
        End
    End If
End If
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Debug.Print "CHECK FOR ADMIN DONE"

FIRST_RUN_SPOT_CHECK_CLIPBOARD = True

FirstRun = True
HOOKSTATold = -10

MNU_SELECTOR_WHOLE_LINE_MODE.Caption = "RUN_FULL_LINE_SELECT=FALSE"

'SET THE DEFAULT OPTION -- TRUE -- PROPER
Call MNU_AUTO_MINIMIZE_OFF_ON_Click




LimitClipSize = 200000
ADATE_APP_BEGIN_DATE = Now


'--------------------------------------------------------------------
Call SET_FOLDER_CLIPBOARD_LOGGER


Debug.Print "CALL SET_FOLDER_CLIPBOARD_LOGGER"


'For r = 1 To 127
'    Debug.Print Str(r) + " -- " + GetSpecialfolder(r)
'Next


If App.PrevInstance = True Then End

Set FS = CreateObject("Scripting.FileSystemObject")


'Debug.Print "PROGRAM BEGIN__"



'PID HAS TO BE -1 -- Process Id Return in PID
'Var - False if Not Exist or PID Remain -1

DO_GO = 0
If UCase(GetComputerName) = "1-ASUS-X5DIJ" Then DO_GO = 1
If UCase(GetComputerName) = "2-ASUS-EEE" Then DO_GO = 1
If UCase(GetComputerName) = "3-LINDA-PC" Then DO_GO = 1
If UCase(GetComputerName) = "5-ASUS-P2520LA" Then DO_GO = 1

If DO_GO = 1 Then
    If IsIDE = False Then
        PID = -1
        Var = cProcesses.GetEXEID(PID, App.Path + "\" + App.EXEName + ".exe")
        If PID <> -1 Then
            Call Process_HIGH_PRIORITY_CLASS(PID)
        End If
    End If
End If

Debug.Print "SET IT OWN HIGH PRIORITY"


'Dim fileName As String
'fileName = "C:\PStart\Progs\#_PortableApps\PortableApps\PicPickPortable\App\picpick\picpick.exe"
'LOAD_PICPICK = False
'If Dir(fileName) <> "" Then
'    PID = -1
'    Var = cProcesses.GetEXEID(PID, fileName)
'    If PID = -1 Then
'        Shell fileName + " /startup", vbNormalNoFocus
'        LOAD_PICPICK = True
'    End If
'    If PID > 0 Then LOAD_PICPICK = True
'End If
'fileName = "C:\Program Files (x86)\PicPick\picpick.exe"
'If Dir(fileName) <> "" And LOAD_PICPICK = False Then
'    PID = -1
'    Var = cProcesses.GetEXEID(PID, fileName)
'    If PID = -1 Then
'        Shell fileName + " /startup", vbNormalNoFocus
'        LOAD_PICPICK = True
'    End If
'    If PID > 0 Then LOAD_PICPICK = True
'End If
'fileName = "C:\Program Files\PicPick\picpick.exe"
'If Dir(fileName) <> "" And LOAD_PICPICK = False Then
'    PID = -1
'    Var = cProcesses.GetEXEID(PID, fileName)
'    If PID = -1 Then
'        Shell fileName + " /startup", vbNormalNoFocus
'    End If
'End If
'
''Dim WSHShell
'On Error Resume Next
'EXECUTE_FILE_NAME = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-Autohotkeys Set AutoRun.ahk"
'If FSO.FileExists(EXECUTE_FILE_NAME) Then
'    Set WSHShell = CreateObject("WScript.Shell")
'    WSHShell.Run """" + EXECUTE_FILE_NAME + """", WSCRIPT_DontShowWindow, DontWaitUntilFinished
'    Set WSHShell = Nothing
'End If
'On Error GoTo 0
'
'If UCase(GetComputerName) <> "3-LINDA-PC" Then
'    If GetWindowsVersion > 5.1 Then
'        If IsIDE = False Then
'            PID = -1
'            Var = cProcesses.GetEXEID(PID, "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe")
'            If PID = -1 Then
'                Shell "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe", vbNormalNoFocus
'            End If
'        End If
'    End If
'End If

MNU_SCANPATH_COUNTER.Visible = False

'Mnu_Minimize.Caption = "MIN"
'Mnu_Max.Caption = "MAX"
'Mnu_Norm.Caption = "NORM"

'Mnu_Height.Visible = False
'FrmJoy.Show

Xmp5 = -1: Ymp5 = -1

Call ScanPath.SET_VAR_FS
Debug.Print "CALL SCANPATH.SET_VAR_FS -- DONE"


'Set FS = CreateObject("Scripting.FileSystemObject")

'Picture2.Picture = LoadPicture(App.Path + "\# DATA\"+GetComputerName+ "\VBIcon4.bmp")
    
'Me.Show
'DoEvents

'STILL NEEDS WORK
IRFANVIEW_PATH3 = "C:\Program Files (X86)\IrfanView\i_view32.exe"
If Dir(IRFANVIEW_PATH3) <> "" Then IRFANVIEW_PATH = IRFANVIEW_PATH3
IRFANVIEW_PATH2 = "C:\Program Files (X86)\IrfanView\i_view32.exe"
If Dir(IRFANVIEW_PATH2) <> "" Then IRFANVIEW_PATH = IRFANVIEW_PATH2
    
    


FileName_er = PATH_TO_CLIPBOARD_APP_DATA_1 + "\Hot Key Capture Text\HotKey-0-Shot % Pic Data Location.txt"
If Dir(FileName_er) <> "" Then
    FR1 = FreeFile
    Open FileName_er For Input As #FR1
        Line Input #FR1, LAST_CLIPBOARD_FILE
    Close #FR1
End If
'--------------------------------------------------------------------
    
Debug.Print "OPEN \HOT KEY CAPTURE TEXT\HOTKEY-0-SHOT % PIC DATA LOCATION.TXT"
    
    
    
Dim FileSpec, TT As Long

'PATH_TO_CLIPBOARD_TEXT_INFO_APP_DAY_DATA

If FSO.FolderExists(App.Path + "\Sound_Wav's") = False Then
    CreateFolderTree App.Path + "\Sound_Wav's"
End If

If FSO.FolderExists(App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\Hot Key Capture Text") = False Then
    CreateFolderTree App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\Hot Key Capture Text"
End If
'--------------------------------------------------------------------

If FSO.FolderExists("D:\# MY DOCS\# 01 My Documents\03 NotePad Files") = False Then
    CreateFolderTree "D:\# MY DOCS\# 01 My Documents\03 NotePad Files"
End If

'--------------------------------------------------------------------


Call ICON_NOT_WANTING


Debug.Print "TRIM DOWN __ ICON_NOT_WANTING"



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
    Open App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\VBIcon4.bmp" For Binary As #FR1
        'Open App.Path + "\VBIcon4.bmp" For Binary As #FR1
        'Open App.Path + "\VBIcon.bmp" For Binary As #1
        Pic1$ = Space$(LOF(FR1))
        Get #FR1, 1, Pic1$
    Close #FR1
End If

Debug.Print "CHECK ICON CAME AS LOAD UP"


FR1 = FreeFile
Open PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total--TRIM.txt" For Binary As #FR1
If LOF(FR1) > 1 * 1024 ^ 2 Then
    Pic12$ = Space$((1 * 1024 ^ 2) + 1)
    Seek FR1, LOF(1) - (1 * 1024 ^ 2)
    Get #FR1, , Pic12$
    Close #FR1
    
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
    
    Kill PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total--TRIM.txt"
    FR1 = FreeFile
    Open PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total--TRIM.txt" For Binary As #FR1
        Put #FR1, , Pic12$
    Close #FR1
    
    'Simple Copy File

    A1 = PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total--TRIM.txt"
    B1 = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\00_ClipBoard_Total--TRIM-" + GetComputerName + "-" + GetUserName + ".txt"
    'fsO.CopyFile ,
    ShellFileCopy A1, B1

End If
Close #1

Debug.Print "GRAB THE TRIM LOGG -- OVER"


'------------
FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\#OutClipChunck.Txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Binary As #FR1
    AD$ = Input$(LOF(FR1), FR1)
Close #FR1


Debug.Print "DO THE OUTCLIP CHUNK FROM TRIM LOHH GRAB IN"


ADTEST$ = AD$
ADTEST_BEFORE$ = AD$
'------------

'COMPARE_EXE_2 = AD$

'MNU_COMPARE.Caption = "COMPARE" + Str(Len(COMPARE_EXE_2)) + " 0"

Call COMPARE_FOR_EXE

Debug.Print "CALL COMPARE_FOR_EXE"

Call CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY(False)

Debug.Print "CALL CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY(FALSE)"

DUPE_CLIPPER_AT_LOAD_FORM = False

Call FORM_LOADER_GET_CLIPBOARD_INFO

Debug.Print "CALL FORM_LOADER_GET_CLIPBOARD_INFO"

'Form1.Show

Call zzLoad_Checks
    
Debug.Print "CALL ZZLOAD_CHECKS"
    
'Text1.SelStart = 0
'Text1.SelLength = Len(Text1)
'Text1.SelColor = &HFF00&
    
vPathSOUND2 = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\"
If Dir(vPathSOUND2, vbDirectory) = "" Then
    I = MkDirNested(vPathSOUND2)
    If I = False Then
        ' MsgBox "UABLE TO MKDIR NESTED" + vbCrLf + vPathSOUND2, vbMsgBoxSetForeground
    End If
End If

    
Call SETUP_SOUND_FILE("")
' Call RESET_SETUP_SOUND_FILE("NOTSOUND")

Debug.Print "CALL SETUP_SOUND_FILE"

'Call MNU_Norm_Click
'Me.WindowState = vbMaximized
'vbMaximized
'vbNORML
'vbMinimized



'Form1.WindowState = vbMinimized
If GetComputerName = "5-ASUS-P2520LA" Then
    Text1.Font.Size = 12
Else
    Text1.Font.Size = 12
End If

Text1.SelStart = Len(Text1.Text)

Label1.Top = 0
Label1.Left = 0
Label1.Caption = ""
Label1.Height = 60
Label1.Left = 0
Label1.Width = Me.Width

Call HEIGHT_BAR_TOP_ROUTINE

Label3.Caption = "INFO _ NOT FIRST CLIPBOARD YET"

Text1.Locked = True

'Text1.ScrollBars = rtfVertical

'SET ALSO IN FORM_RESIZE AKA SETUP_DIMENSIONS_INNER_FORM
'Call MNU_Norm_Click

'Load FrmJoy

'Exit Sub
API_LOAD = True
Load FRMCLIPTEST01
'Load FRMCLIPTEST02





Debug.Print "LOAD FRMCLIPTEST"


'WE CAN SET THE EXECUTE TO TRUE AFTER _ NOTHING WILL CLIPBOARD AT FIRST RUN
Call Timer1_Timer



Debug.Print "CALL TIMER1_TIMER"

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
Timer_SCREEN_SHOT_AUTO.Enabled = True

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
Debug.Print "CALL MNU_MIRROR_SEND_TO_OPERATING_SYSTEM_SET_MENU_CAPTION"
'--------------------------------------------------------
'SEND TO
Call SPECIAL_FOLDER_SEND_TO
Debug.Print "CALL SPECIAL_FOLDER_SEND_TO"
'--------------------------------------------------------
VAR_FLAG_EXPLORER_LOAD_NOT = True
'START MENU
Call MNU_SPECIAL_FOLDER_PIN_TO_START_MENU_Click
Debug.Print "CALL MNU_SPECIAL_FOLDER_PIN_TO_START_MENU_CLICK"
'DESKTOP
Call MNU_SPECIAL_FOLDER_DESKTOP_USER_Click
Debug.Print "CALL MNU_SPECIAL_FOLDER_DESKTOP_USER_CLICK"
Call MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC_Click
Debug.Print "CALL MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC_CLICK"
'--------------------------------------------------------
'CHECK HERE - WINMERGE
'Call SPECIAL_FOLDER_SENT_TO
Call SPECIAL_FOLDER_PIN_TO_START_MENU
'--------------------------------------------------------
Debug.Print "CALL SPECIAL_FOLDER_PIN_TO_START_MENU"

' MNU_EXPLORER_ME_VB.Caption = "EXPLORER ME_VB -- " + App.Path
Call SET_MNU_EXPLORER_ME_VB("EXPLORER ME_VB -- " + App.Path)
Debug.Print "CALL SET_MNU_EXPLORER_ME_VB(""EXPLORER ME_VB -- "" + APP.PATH)"

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
Debug.Print "CALL SETUP_DIMENSIONS_NORM"
Debug.Print "CALL SETUP_DIMENSIONS_INNER_FORM"


'RESIZE_AT_LOAD = False
DONT_RESIZE_WHILE_SETUP = False

If Timer_API_Test_Logick_Var2 = 0 Then
    VARTEXT = "THE TIME THE API CLIPBOARD FUNCTION LAST ACCESSED = NOT YET"
    MNU_TIME_API_FUNCTION_ACCESS.Caption = VARTEXT
End If

'STARTUP TIME
ADATE_APP_BEGIN_DATE2 = DateDiff("S", ADATE_APP_BEGIN_DATE, Now)

MNU_IDLE_ACTIVE.Caption = "IDLE ><>< ACTIVE"

'If IsIDE = True Then Mnu_VB.Enabled = False

'DoEvents

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

' Call MNU_SOUND_OPTION_Click(0)
' Call RESET_SETUP_SOUND_FILE("")

'MNU_404_CPC_PAGES.Visible = False

'If IsIDE = True Then
'    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = True
'Else
'    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = False
'End If

'If MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = False Then
'    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Visible = True
'End If

Call MNU_NIRSOFT_SETUP


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

If IsIDE = True Then
    CALC_MODE = True
    MNU_CALC_MODE_OFF_AND_ON.Caption = "CALC MODE=ON"
Else
    CALC_MODE = False
    MNU_CALC_MODE_OFF_AND_ON.Caption = "CALC MODE=OFF"
End If
Timer_SET_LABEL3_INFO_BAR.Enabled = True

Timer1_PLUS.Enabled = True

CALC_MODE = False
MNU_CALC_MODE_OFF_AND_ON.Caption = "CALC MODE=OFF"


EBAY_YEAR_IN = 1
Call MENU_EBAY_CAPTION_1
Call MENU_EBAY_CAPTION_2

Call SET_MENU_PADD_WORK
Debug.Print "CALL SET_MENU_PADD_WORK"

' #NFS_EX EXCLUDE FROM SYNCER
' ---------------------------

On Error Resume Next
FR1 = FreeFile
VAR_1 = App.Path + "\# DATA\GETCOMPUTERNAME_&_GETUSERNAME_SET__"
VAR_2 = GetComputerName + "__" + GetUserName + "__"
If Dir(App.Path + "\# DATA", vbDirectory) = "" Then
    MkDir App.Path + "\# DATA"
End If
Open VAR_1 + VAR_2 + "#NFS.TXT" For Output As #FR1
    Print #FR1, "TIME DATE BE HOLD IN FILE SYSTEM AND USER TO DETECT THE RECENT COMPUTER FOR PURPOSE CLIPBOARD CHUNK SEND GATHER FROM OTHER COMPUTER"
Close #FR1
' DETECT ALL BUT OWN
' TIMER
' TIMER_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER -- DEFAULT ENABLE -- FALSE
' MENU MNU _BUTTON GO
' MNU_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER
' DECLARE
' MNU_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_VALUE
' TIMER
' "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\4-ASUS-GL522VW_MATT 01\#OutClipChunck.Txt"
' Thu 28-May-2020 17:35:05
TIMER_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER.Enabled = True
TIMER_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER.Interval = 1000
On Error GoTo 0


FORM_LOAD_SIZE_NOT_SET_YET = True

On Error Resume Next
If IsIDE = True Then
    Me.WindowState = vbNormal
Else
    Me.WindowState = vbMinimized
End If
On Error GoTo 0

'TIMER_VB_PROJECT_DATE.Enabled = True
'TIMER_VB_PROJECT_DATE.Interval = 4000
'Debug.Print "TIMER_VB_PROJECT_DATE.ENABLED = TRUE"

Debug.Print "FORM_LOADED ON"

STARTUP_RUN = False


End Sub



Sub FORM_LOADER_GET_CLIPBOARD_INFO()

On Error GoTo EXIT_RETRY_AGAIN_ERROR
Err.Clear
OO_1 = Clipboard.GetFormat(vbCFText)
OO_2 = Clipboard.GetFormat(vbCFBitmap)

Clipboard_GetFormat_vbCFText_OO_1 = OO_1
Clipboard_GetFormat_vbCFBitmap_OO_2 = OO_2

On Error GoTo 0

If Clipboard_GetFormat_vbCFText_OO_1 = True Then
    Mnu_Clip_Description.Caption = "Clip Format:- " + "Text (.txt file)"
    Mnu_Clip_Description.Caption = "Clip Format:- " + "Text" ' (.txt file)"
    
    'MEMORY TEXT AT STARTER
    '----------------------
    On Error GoTo EXIT_RETRY_AGAIN_ERROR
    AD2 = Trim(Clipboard.GetText)
    On Error GoTo 0
    
    If AD2 <> "" Then
        If AD$ = AD2 Then
            AD2 = ""
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

Call CLIPBOARD_READ_DATA_LOGG_AT_FORM_LOAD

Call NEXT_LOADING_FORM_LOADER_GET_CLIPBOARD_INFO



' ----------------------------------------------
Exit Sub

' ----------------------------------------------
EXIT_RETRY_AGAIN_ERROR:
' ----------------------------------------------
CLIPBOARD_RETURN_TIMER_ERROR_9 = 1
Timer_MNU_COPY_2.Enabled = True


End Sub


Sub CLIPBOARD_READ_DATA_LOGG_AT_FORM_LOAD()

Text1 = ""
On Error Resume Next
If Dir(PATH_TO_CLIPBOARD_APP_DATA_1 + "\Start.txt") <> "" Then
    FR1 = FreeFile
    Open PATH_TO_CLIPBOARD_APP_DATA_1 + "\Start.txt" For Input As #FR1
    Line Input #FR1, Star
    Close #FR1
End If
On Error GoTo 0

'If Star = "" Then Star = Str(Now - DateSerial(0, 0, 40))

start = Now
If Star = "" Then Star = Now


'#IF ITS THE SAME DAY
CountR2$ = ""
If DateValue(Star) = DateValue(start) Then
    FR1 = FreeFile
    Open PATH_TO_CLIPBOARD_APP_DATA_1 + "\#ClipBoard.Txt" For Binary As #FR1
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
    
    FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\#Count.Txt"
    FR1 = FreeFile
    If Dir(FILENAME_IN_USE_CHECK) <> "" Then
        Open FILENAME_IN_USE_CHECK For Input As #FR1
        If Not (EOF(FR1)) Then Line Input #FR1, CountR2$
        Close #FR1
    End If
    CountR = Val(CountR2$)

Else
    tq = PATH_TO_CLIPBOARD_APP_DATA_1 + "\#ClipBoard_Old.Txt"
    If Dir(tq) <> "" Then
        Kill tq
        If Dir(PATH_TO_CLIPBOARD_APP_DATA_1 + "\#ClipBoard.Txt") <> "" Then
            Name PATH_TO_CLIPBOARD_APP_DATA_1 + "\#ClipBoard.Txt" As PATH_TO_CLIPBOARD_APP_DATA_1 + "\#ClipBoard_Old.Txt"
        Else
            Open PATH_TO_CLIPBOARD_APP_DATA_1 + "\#ClipBoard.Txt" For Output As #FR1
            Close #FR1
        End If
    End If
    CountR = 0
End If

FR1 = FreeFile
Open PATH_TO_CLIPBOARD_APP_DATA_1 + "\Start.txt" For Output As #FR1
    Print #FR1, Format$(start, "DD-MM-YYYY")
Close #FR1

'After load most import recent load a backlogg for viewing
FR1 = FreeFile
Open PATH_TO_CLIPBOARD_APP_DATA_1 + "\#ClipBoard_Old.Txt" For Binary As #FR1
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

End Sub


Sub NEXT_LOADING_FORM_LOADER_GET_CLIPBOARD_INFO()

On Error GoTo EXIT_RETRY_AGAIN_ERROR

Dim COMPARE_1 As String, COMPARE2 As String
'PICTURE COMPARE
DUPE_IMAGE_AT_LOAD_FORM = False


If Clipboard_GetFormat_vbCFBitmap_OO_2 = True Then
    Picture3.Picture = Clipboard.GetData(vbCFBitmap)
End If
    
On Error GoTo 0

If Clipboard_GetFormat_vbCFBitmap_OO_2 = True Then
    SavePicture Picture3.Picture, App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\Hot Key Capture Text\HotKey-Shot Pic-TEST.jpg"
    'LAST_CLIPBOARD_FILE = TS1

    If Dir(App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\Hot Key Capture Text\HotKey-Shot Pic-TEST.jpg") <> "" Then
        FR1 = FreeFile
        Open App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\Hot Key Capture Text\HotKey-Shot Pic-TEST.jpg" For Binary As #FR1
            COMPARE_1 = Space$(LOF(FR1))
            Get #FR1, 1, COMPARE_1
        Close #FR1
        Kill App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\Hot Key Capture Text\HotKey-Shot Pic-TEST.jpg"
        'ANOTHER COMPARE IN MAIN ROUTINE
        PICXX$ = COMPARE_1
    End If

    FILENAME1 = App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\Hot Key Capture Text\COMPARE DUPE.JPG"
    
    'If Dir(App.Path + "\# DATA\" + GetComputerName +"_"+GETUSERNAME+ "\Hot Key Capture Text\HotKey-Shot Pic.jpg") <> "" Then
    If Dir(FILENAME1) <> "" Then
        
        FR1 = FreeFile
        Open FILENAME1 For Binary As #FR1
            COMPARE2 = Space$(LOF(FR1))
            Get #FR1, 1, COMPARE2
        Close #FR1
        
        If COMPARE_1 = COMPARE2 Then
            Call Menu_Clipboard_size(Len(COMPARE_1))
            If FirstRun = True Then
                'FirstRun = False
                'Mnu_LAST_CLIP_TIME.Caption = Mnu_LAST_CLIP_TIME.Caption + " First Run"
                Mnu_Run_Status.Caption = "1st Run"

            End If

            COMPARE_1 = "": COMPARE2 = ""
            DUPE_IMAGE_AT_LOAD_FORM = True
            DUPE_CLIPPER_AT_LOAD_FORM = True
            
            Mnu_Clip_Status.Caption = "Dupe Image at Starter"
        Else
            Mnu_Clip_Status.Caption = "Image at Starter"
        End If
        
        'Mnu_Clip_Description.Caption = "Clip Format:- " + "Bitmap (.bmp file)"
        GetFormat_And_Display
        Mnu_Clip_Status.Caption = "At Starter"
        xyz2020 = True
    
    End If
End If

If xyz2020 = False Then
    GetFormat_And_Display
    Mnu_Clip_Status.Caption = "At Starter"
End If

' ----------------------------------------------
Exit Sub

' ----------------------------------------------
EXIT_RETRY_AGAIN_ERROR:
' ----------------------------------------------
CLIPBOARD_RETURN_TIMER_ERROR_10 = 1
Timer_MNU_COPY_2.Enabled = True



End Sub



Function GET_MENU_CONTROL_LENGTH(I)

Dim Control As Control
Dim Text_Checker_Form_Menu As String
GET_MENU_CONTROL_LENGTH = 0

For I = 0 To Forms.Count - 1
    If Forms(I).hWnd = Me.hWnd Then
        Text_Checker_Form_Menu = ""
        frmListMenu.GetMenuInfo_Not_SUB GetMenu(Forms(I).hWnd), 0, "", Text_Checker_Form_Menu
        GET_MENU_CONTROL_LENGTH = Len(Text_Checker_Form_Menu)
    End If
Next

End Function



Sub SET_MENU_PADD_WORK()

Dim i_Menu_Count, i_Form_Counter
Dim i_Menu_Not_Visa_Count

Dim Control As Control, LABEL_44, LABEL_48

Dim R_NEXT

Dim Text_Checker_Form_Menu As String

For I = 0 To Forms.Count - 1
    
    For Each Control In Forms(I).Controls
        If InStr(UCase(Control.Name), "MNU_") And InStr(UCase(Control.Name), "TIMER_") = 0 Then
            
            On Error Resume Next
                XCONTROL_V = -2
                XCONTROL_V = Control.Visible
                If Err.Number > 0 Then Debug.Print Control.Name + " -- " + "HAS ERROR TO USER HERE -- SUB SET_MENU_PADD_WORK"
            On Error GoTo 0
            If XCONTROL_V = True Then
                i_Menu_Count = i_Menu_Count + 1
            End If
            i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count + 1

        End If
    Next
Next

i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count - i_Menu_Count

Mnu_Menu_Item_Count.Caption = "Menu Item Count = " + Str(i_Menu_Count) + " &&" + Str(i_Menu_Not_Visa_Count) + " Not Visible"
'Mnu_Form_Count.Caption = "Form Counter = " + Str(Forms.Count - 1) '  + " Really 7"
'Mnu_Form_Count.Visible = False

i_Menu_Count = 0



For I = 0 To Forms.Count - 1
    Text_Checker_Form_Menu = ""
    frmListMenu.GetMenuInfo_Not_Indented GetMenu(Forms(I).hWnd), 0, "", Text_Checker_Form_Menu
    'frmListMenu.GetMenuInfo_Not_Indented GetMenu(Forms(I).hWnd), 0, "", Text_Checker_Form_Menu
    Text_Checker_Form_Menu = UCase(Text_Checker_Form_Menu)
    For Each Control In Forms(I).Controls
        If InStr(UCase(Control.Name), "MNU_") > 0 And InStr(UCase(Control.Name), "TIMER_") = 0 Then
            MENU_ITEM_VAR = Replace(Control.Caption, "[__ ", "")
            MENU_ITEM_VAR = Replace(MENU_ITEM_VAR, " __]", "")
            MENU_ITEM_VAR = UCase(Trim(MENU_ITEM_VAR))
            If InStr(Text_Checker_Form_Menu, "SUB MENU ----" + MENU_ITEM_VAR) = 0 Then
                
                'i_Menu_Count = i_Menu_Count + 1
                If InStr(Trim(Control.Caption), "[__ ") = 0 Then
                    LABEL_44 = Trim(Control.Caption)
                    LABEL_48 = LABEL_44
                    'LABEL_48 = Replace(LABEL_44, " ", "_")
                    LABEL_48 = Replace(LABEL_48, "___", "__")
                    LABEL_48 = "[__ " + LABEL_48 + " __]"
                    LABEL_48 = Replace(LABEL_48, "[__ [__ ", "[__ ")
                    LABEL_48 = Replace(LABEL_48, " __] __]", " __]")
                    If LABEL_48 <> LABEL_44 Then
                        Control.Caption = LABEL_48
                    End If
                End If
            End If
        End If
    Next
Next

'Stop

''MNU_BRing_Front
''---------------
'i_Form_Counter = Forms.Count
'i_Form_Counter = 0
''for each f
''For i = 0 To Forms.Count - 1
''    Load Forms(i)
''    Show Forms(i)
''Next
'
'Dim Form As Form
'i_Form_Counter = 0
'For Each Form In Forms
'    i_Form_Counter = i_Form_Counter + 1
'    Load Form
'    Form.Show
'    Show Form
'    'Set Form = Nothing
'Next Form
'
'i_Form_Counter = 0
'For Each Form In Forms
'    i_Form_Counter = i_Form_Counter + 1
'    Load Form
'    Form.Show
'    'Set Form = Nothing
'Next Form

i_Form_Counter = Forms.Count - 1
Me.Refresh
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Text1.Visible = False
'Label4.Caption = Str(X) + " -- " + Str(Y)

End Sub

Private Sub Form_Paint()

' -----------------------------------------------------------------------------------------------
' Text1.Visible = False
' Label4.Caption = Str(X) + " -- " + Str(Y)
' WANT TO TRY AND GET IF USER IS RESIZE THE FORM
' NONE SEARCHER
' LOOK AT MOUSE CHANGE SHAPE
' THE CORDINATE BOTTOM LEFT
' BUT WIN 10 HAS SIDE OF FORM
' AND THEN BE DETERMINE IF WANT OT RESIZE MANY INTERNAL ITEM
' AS NOT WANT NORMAL WHEN USER SELECT TO NOT JUMP TO BOTTOM AT ANYTIME LIKE WHEN CLIPBOARD ARRIVE
' -----------------------------------------------------------------------------------------------

End Sub

Private Sub Label3_Click()
' Label3.CAPTION
' Label3.WIDTH
' Label3.FONT

End Sub

Private Sub MNU_ABORT_SHUTDOWN_Click()
    
    'CSIDL_WINDOWS_SYSTEM32 = 37
    Shell GetSpecialfolder(37) + "\shutdown.exe /a", vbNormalFocus

End Sub

Private Sub Label5_Click()

End Sub





Private Sub Mnu_API_Reload_Click()

On Error Resume Next
API_LOAD = True
Load FRMCLIPTEST01

End Sub

Public Sub MNU_API_RESET_Click()

'-----------------------------
O_VAL_MINUTE_API = Now

Exit Sub

'Call Form1.MNU_API_RESET_SUB
'-----------------------------
'Call MNU_API_RESET_SUB
Call RESET_SETUP_SOUND_FILE("")

Call Mnu_API_Unload_Click
Call Mnu_API_Reload_Click

End Sub

Private Sub MNU_AUTO_MINIMIZE_OFF_ON_Click()

SET_1 = False
SET_2 = False
If InStr(MNU_AUTO_MINIMIZE_OFF_ON.Caption, "AUTO MINIMIZE=FALSE") Then SET_1 = True
If InStr(MNU_AUTO_MINIMIZE_OFF_ON.Caption, "AUTO MINIMIZE=TRUE") Then SET_2 = True
If SET_1 = False And SET_2 = False Then
    MNU_AUTO_MINIMIZE_OFF_ON.Caption = "[__ AUTO MINIMIZE=TRUE __]"
    Exit Sub
End If

If InStr(MNU_AUTO_MINIMIZE_OFF_ON.Caption, "AUTO MINIMIZE=FALSE") Then
    MNU_AUTO_MINIMIZE_OFF_ON.Caption = "[__ AUTO MINIMIZE=TRUE __]"
    Exit Sub
End If
If InStr(MNU_AUTO_MINIMIZE_OFF_ON.Caption, "AUTO MINIMIZE=TRUE") Then
    MNU_AUTO_MINIMIZE_OFF_ON.Caption = "[__ AUTO MINIMIZE=FALSE __]"
    Exit Sub
End If

End Sub

Private Sub MNU_CALC_GIVE_RESULT_Click()
        
    EXECUTE_TIMER_ENABLED = False
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    
    Do
        Err.Clear
        Clipboard.SetText Trim(Mid(CALC_STRINGER_S, InStrRev(CALC_STRINGER_S, "=") + 1))
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    AD$ = Trim(Mid(CALC_STRINGER_S, InStrRev(CALC_STRINGER_S, "=") + 1))
    Beep
    EXECUTE_TIMER_ENABLED = True
        
    Exit Sub
    
    SET_GO = True
    If SET_GO = True Then
        On Error Resume Next
        TIMER_Clipboard_Clear_VAR = False
        '10-10-10
        TIMER_Clipboard_Clear.Enabled = True
        Do
            DoEvents: Sleep 10
        Loop Until TIMER_Clipboard_Clear_VAR = True
            
        Clipboard_Set_Text_VAR = Trim(Mid(CALC_STRINGER_S, InStrRev(CALC_STRINGER_S, "=") + 1))
                
        Call PUT_TEXT_CLIPBOARD(Clipboard_Set_Text_VAR)
        
        On Error GoTo 0
    End If
    
End Sub

Sub PUT_TEXT_CLIPBOARD(Clipboard_Set_Text_VAR)
        
        ENTER_TEXT_IN_LOGGER_FOREGROUND_OVERRIDE = True
            X_COUNTER_VAR = 0
            Do
                X_COUNTER_VAR = X_COUNTER_VAR + 1
                Err.Clear
                Clipboard.SetText Clipboard_Set_Text_VAR
                
                'DoEvents: Sleep 10
                'If TIMER_Clipboard_Set_Text.Enabled = False Then
                    'TIMER_Clipboard_Set_Text.Interval = 1
                '    TIMER_Clipboard_Set_Text.Enabled = False
                '    TIMER_Clipboard_Set_Text.Interval = 10
                '    TIMER_Clipboard_Set_Text.Enabled = True
                'End If
                'If X_COUNTER_VAR = 50 Then
                '    If IsIDE = True Then Stop
                    '----------------------------------------------------------------------------------------------------------------
                    ' WEIRD ERROR HAPPEN HERE THE  TIMER_Clipboard_Set_Text.ENABLED = TRUE
                    ' IF WE PUSH A RUN CALL THROUGH RESET WORKING AGAIN
                    ' BUT NORMAL TIMER AND ERROR ITERMITANTLY STOPS THE CALL FROM TIMER
                    ' WORK AROUND THIS WAY
                    ' SOMETHING TO DO WITH THE CLIPBOARD CALL SET
                    ' CALLING A TIMER AND WAITING IN A  DO LOOP SEEMS A BIT MUCH
                    ' SLEEP AND DOEVENTS WOULD DO THE SAME
                    '----------------------------------------------------------------------------------------------------------------
                '    Call TIMER_Clipboard_Set_Text_Timer
                'End If
            
            If X_COUNTER_VAR = 100 Then
                MSGBOX2 = "ERROR UNABLE TO SET CLIPBOARD TO" + vbCrLf + vbCrLf + Clipboard_Set_Text_VAR
                frm_MSGBOX.Timer1.Enabled = True
                Exit Do
            End If
            
            If Err.Number > 0 Then Sleep 100
            
            DoEvents
        
        Loop Until Clipboard.GetText = Clipboard_Set_Text_VAR Or X_COUNTER_VAR > 100
        Clipboard_Set_Text_VAR = ""

End Sub

Private Sub MNU_CALC_MODE_OFF_AND_ON_Click()
If MNU_CALC_MODE_OFF_AND_ON.Caption = "CALC MODE=ON" Then
    CALC_MODE = False
    MNU_CALC_MODE_OFF_AND_ON.Caption = "CALC MODE=OFF"
Else
    CALC_NEW_INPUT = True
    CALC_NEW_INPUT_MNU = True
    CALC_MODE = True
    MNU_CALC_MODE_OFF_AND_ON.Caption = "CALC MODE=ON"
    Call SET_LABEL3
End If
End Sub

Private Sub MNU_CALC_POUND_ALL_CHUNK_Click()

TT_VAR = ""
XX_VAR = 1
EXIT_LOOP = False
Do
    SEARCH_VAR = InStr(XX_VAR, AD$, "£")

    If SEARCH_VAR = 0 Then
        EXIT_LOOP = True
        Exit Do
    End If

    For r = 1 To 20
        NUMBER_VAR = IsNumeric(Mid(AD$, SEARCH_VAR + 1, r))
        If NUMBER_VAR = False Then
            Exit For
        End If
    Next
    

    TT_VAR = TT_VAR + Mid(AD$, SEARCH_VAR, r - 1) + vbCrLf

    XX_VAR = SEARCH_VAR + 1

Loop Until EXIT_LOOP = True


AD$ = UCase(TT_VAR)

    'EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        Clipboard.SetText AD$
    'EXECUTE_TIMER_ENABLED = True

'Me.WindowState = vbMinimized
AD_DATE = 0

Sleep 1000

Call MNU_CALC_RESULT_AND_SUM_Click

Beep

End Sub

Private Sub MNU_CALC_RESULT_AND_SUM_Click()

    On Error Resume Next
    TIMER_Clipboard_Clear_VAR = False
    TIMER_Clipboard_Clear.Enabled = True
    
    Do
        DoEvents
        If TIMER_Clipboard_Clear_VAR = False Then Sleep 10
    Loop Until TIMER_Clipboard_Clear_VAR = True
    
    ENTER_TEXT_IN_LOGGER_FOREGROUND_OVERRIDE = True
    Clipboard_Set_Text_VAR = CALC_STRINGER_S
    Call PUT_TEXT_CLIPBOARD(Clipboard_Set_Text_VAR)

    Beep
    
    'Me.WindowState = vbMinimized
    
    On Error GoTo 0

End Sub

Private Sub MNU_CLIPBOARD_IMAGE_FILENAME_Click()
    
    If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
        Call RECURSIVE_SEARCH_LAST_CLIPBOARD_ART_FOLDER
    End If
    
        If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
        MsgBox "File None Found Last Art Clipboard Search Var _ LAST_CLIPBOARD_FILE"
        Exit Sub
    End If
    
    If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
        MsgBox "File Found Not Valid Last Art Clipboard Search Var _ LAST_CLIPBOARD_FILE"
        Exit Sub
    End If
    
    Clipboard.Clear
    Clipboard.SetText LAST_CLIPBOARD_FILE
    

    Dim ReturnValue As Long
    If IRFANVIEW_PATH <> "" Then
'        ReturnValue = Shell(IRFANVIEW_PATH + " """ + LAST_CLIPBOARD_FILE + """ /fs /silent", vbMaximizedFocus)
        ' REM
        
'        ReturnValue2 = Str(ReturnValue)
       On Error Resume Next
'        AppActivate ReturnValue2
        AppActivate ReturnValue
        ' IF DON'T WANT FULL SCREEN
        ' SENDKEY ENTER TO THE THING
        On Error Resume Next
'        SendKeys "{ENTER}", False
        On Error GoTo 0
        ' Shell IRFANVIEW_PATH + " """ + LAST_CLIPBOARD_FILE + """ /bf /display=(100,100,300,300,50,0,0) /silent", vbNormalFocus ' , vbMaximizedFocus
    Else
        Me.WindowState = vbNormal
        MsgBox "IRFANVIEW_PATH VAR -- PATH NOT FOUND FOR FILE" + vbCrLf + "NOT INSTALED AT EXPECTED LOCATION " + IRFANVIEW_PATH3 + vbCrLf + "OR" + vbCrLf + IRFANVIEW_PATH2, vbMsgBoxSetForeground
    End If

    Me.WindowState = vbMinimized

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

Sub MNU_TSF_CLICK()

End Sub

Sub MNU_TSP_CLICK()

End Sub

Sub MNU_TSS_CLICK()

End Sub

Sub MNU_FLICKR_YAHOO_CLICK()

End Sub

Sub MNU_FLCIKR_SYNC_CLICK()

End Sub



Public Sub MNU_COPY_2_CLIPPER_BEFORE_Click()

    Dim STRING_VAR_011
    Dim STRING_VAR_012
    VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_COPY_2_CLIPPER_BEFORE_Click"
    
    A1 = "---------------------" + vbCrLf + "---------------------" + vbCrLf + "Count ="
    
    FPOS = InStrRev(Text1.Text, A1)
    FPOS = InStrRev(Text1.Text, A1, FPOS - Len(A1) - 1)
    STRING_VAR_00 = Mid(Text1.Text, FPOS + Len("---------------------" + vbCrLf))

'    MsgBox STRING_VAR_00

    EXECUTE_TIMER_ENABLED = False
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    Err.Clear
    Clipboard.SetText STRING_VAR_00

    If Err.Number > 0 Then
        ' PUBLIC HERE -- NAME OF SUB WHOLE
        ' USE VARIBALE AT TOP -- VAR_TIMER_CLIPBOARD_TIMER_RETRY
        ' ------------------------------------------------------
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
        Exit Sub
    End If
    EXECUTE_TIMER_ENABLED = False
    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_EBAY_10_Click()
    EBAY_YEAR_IN = "10,"
    Call MENU_EBAY_CAPTION_2
    Call MNU_EBAY_MULTIPLE_PAGE_Click
End Sub

Private Sub MNU_EBAY_YEAR_1_10_Click()
    EBAY_YEAR_IN = "10-YEAR_1"
    Call MENU_EBAY_CAPTION_2
    Call MNU_EBAY_MULTIPLE_PAGE_Click
End Sub

Private Sub MNU_EBAY_YEAR_1_Click()
    EBAY_YEAR_IN = "1,"
    Call MENU_EBAY_CAPTION_2
    Call MNU_EBAY_MULTIPLE_PAGE_Click
End Sub

Private Sub MNU_EBAY_YEAR_2_Click()
    EBAY_YEAR_IN = "2,"
    Call MENU_EBAY_CAPTION_2
    Call MNU_EBAY_MULTIPLE_PAGE_Click
End Sub

Private Sub MNU_EBAY_YEAR_3_Click()
    EBAY_YEAR_IN = "3,"
    Call MENU_EBAY_CAPTION_2
    Call MNU_EBAY_MULTIPLE_PAGE_Click
End Sub

Private Sub MNU_EBAY_MULTI_Click(Index As Integer)
    
    Select Case Index
    Case 1
        EBAY_YEAR_IN = "1,2,3"
    Case 2
        EBAY_YEAR_IN = "1,"
    Case 3
        EBAY_YEAR_IN = "1,2,"
    Case 4
        EBAY_YEAR_IN = "2,"
    Case 5
        EBAY_YEAR_IN = "2,3,"
    Case 6
        EBAY_YEAR_IN = "3,"
    End Select
    
    Call MENU_EBAY_CAPTION_2
    Call MNU_EBAY_MULTIPLE_PAGE_Click
    
    
End Sub

Sub MENU_EBAY_CAPTION_1()
    MNU_EBAY_YEAR_1.Caption = "EBAY " + Trim(Str(Year(Now)))
    MNU_EBAY_YEAR_2.Caption = "EBAY " + Trim(Str(Year(Now)) - 1)
    MNU_EBAY_YEAR_3.Caption = "EBAY " + Trim(Str(Year(Now)) - 2)
    
    i1 = " " + Trim(Str(Year(Now)))
    i2 = " " + Trim(Str(Year(Now)) - 1)
    I3 = " " + Trim(Str(Year(Now)) - 2)
    
    MNU_EBAY_MULTI(1).Caption = "EBAY " + i1 + i2 + I3
    MNU_EBAY_MULTI(2).Caption = "EBAY " + i1
    MNU_EBAY_MULTI(3).Caption = "EBAY " + i1 + i2
    MNU_EBAY_MULTI(4).Caption = "EBAY " + i2
    MNU_EBAY_MULTI(5).Caption = "EBAY " + i2 + I3
    MNU_EBAY_MULTI(6).Caption = "EBAY " + I3

End Sub

Sub MENU_EBAY_CAPTION_2()
    
    EBAY_MNU_CAP_YEAR_IN = "EBAY MULTIPLE PAGE - "
    For r = 1 To 3
        If InStr(EBAY_YEAR_IN, Trim(Str(r)) + ",") Then
            EBAY_MNU_CAP_YEAR_IN = EBAY_MNU_CAP_YEAR_IN + " " + Trim(Str(Year(Now) - (r - 1)))
        End If
    Next
    MNU_EBAY_MULTIPLE_PAGE.Caption = EBAY_MNU_CAP_YEAR_IN

End Sub


Private Sub MNU_EBAY_MULTIPLE_PAGE_Click()

    Call MENU_EBAY_CAPTION_2

    Me.WindowState = vbMinimized
    
    URL_PATH_ORGINAL = "https://www.ebay.co.uk/myb/PurchaseHistory#PurchaseHistoryOrdersContainer?ipp=50&Period=2&Filter=1&radioChk=1&GotoPage=1"
    URL_PATH = URL_PATH_ORGINAL
    
    URL_PATH = Mid(URL_PATH, 1, InStr(LCase(URL_PATH), LCase("&GotoPage=")) + Len("&GotoPage=1") - 2)
    
    Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    Sleep 1000
    ' WAIT FOR CHROME TO LOAD A BIT-AH
    ' ---------------------------------
    
    ' CURRENT YEAR
    ' ONLY LOAD PAGE HOW FAR BACK CURRENT YEAR
    ' AND ADDTIONAL PAGE AFTER NONE LEFT WILL BE SAME LAST ONE
    ' AND THEN CHANGE THE YEAR
    ' EXAMPLE I HAVE 13 PAGE OF 50 IN 2019
    ' NOTE EBAY STRESS MEAN LESS ABLE TO LOAD A 100 ITEM PAGE NOW-A-DAY
    ' SO 50 WILL DO
    ' --------------------------------------------------------
    
    YEAR_2020_PAGE_COUNT = 14
    ' 2020 -- WILL CHANGE WHEN NEW YEAR ARRIVE
    ' IS TOP YEAR
    If InStr(EBAY_YEAR_IN, "1,") Then
    
        URL_COUNT = 0
        ' -----------------------------------
        ' 2019 IS CURRENTLY 15 PAGE SET OF 50
        ' 2020 IS THE NEW
        ' -----------------------------------
        ' NOW CHANGE IS 30 PAGE SET
        ' SEEM LAYOUT HAS MOVER AS SELECT HOW MANY A PAGE NOT THERE
        ' -----------------------------------
        ' 2020 SET HOW MANY PAGE OF 1 TO 10
        For R_EBAY = 1 To YEAR_2020_PAGE_COUNT
    
            MNU_EBAY_MULTIPLE_PAGE.Caption = MNU_CAP + Format(R_EBAY, "00")
            ShellExecute Me.hWnd, "open", URL_PATH + Format(R_EBAY, "0"), vbNullString, vbNullString, conSwNormal
    
            Sleep 1000
            URL_COUNT = URL_COUNT + 1
    
            If URL_COUNT > 100 Then
                URL_COUNT = 0
                Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", vbNormalFocus
                Sleep 1000
            End If
    
        Next R_EBAY
    
    End If
    
    ' 2019 IS CURRENTLY 30 PAGE SET OF 50
    ' -----------------------------------
    ' NOW 2019 AS ---- Thu 02-Jan-2020 13:00:54 -- 30 PAGE
    ' -----------------------------------
    If InStr(EBAY_YEAR_IN, "2,") Then
        ' DROP DOWN TO THE NEXT LOWER YEAR
        ' --------------------------------
        URL_PATH = URL_PATH_ORGINAL
        URL_PATH_1 = Mid(URL_PATH, 1, InStr(LCase(URL_PATH), LCase("&Period=")) + Len("&Period=") - 1)
        URL_PATH_2i = Mid(URL_PATH, InStr(LCase(URL_PATH), LCase("&Period=")) + Len("&Period="), 1)
        URL_PATH_2i = Trim(Str(Val(URL_PATH_2i) + 1))
        URL_PATH_3 = Mid(URL_PATH, InStr(LCase(URL_PATH), LCase("&Period=")) + Len("&Period=") + 1)
        URL_PATH = URL_PATH_1 + URL_PATH_2i + URL_PATH_3
        
        URL_PATH = Mid(URL_PATH, 1, InStr(LCase(URL_PATH), LCase("&GotoPage=")) + Len("&GotoPage=1") - 2)
        Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        Sleep 1000
        URL_COUNT = 0
        For R_EBAY = 1 To 30
            MNU_EBAY_MULTIPLE_PAGE.Caption = MNU_CAP + Format(R_EBAY, "00")
            ShellExecute Me.hWnd, "open", URL_PATH + Format(R_EBAY, "0"), vbNullString, vbNullString, conSwNormal
        
            Sleep 1000
            URL_COUNT = URL_COUNT + 1
        
            If URL_COUNT > 100 Then
                URL_COUNT = 0
                Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", vbNormalFocus
                Sleep 1000
            End If
        Next R_EBAY
    End If

    If InStr(EBAY_YEAR_IN, "3,") Then
        ' DROP DOWN TO THE NEXT LOWER YEAR
        ' --------------------------------
        ' -------------------------
        ' 2018 IS 29 PAGE SET OF 50
        ' Wed 06-Nov-2019 12:55:17
        ' -------------------------
        URL_PATH = URL_PATH_ORGINAL
        URL_PATH_1 = Mid(URL_PATH, 1, InStr(LCase(URL_PATH), LCase("&Period=")) + Len("&Period=") - 1)
        URL_PATH_2i = Mid(URL_PATH, InStr(LCase(URL_PATH), LCase("&Period=")) + Len("&Period="), 1)
        URL_PATH_2i = Trim(Str(Val(URL_PATH_2i) + 2))
        URL_PATH_3 = Mid(URL_PATH, InStr(LCase(URL_PATH), LCase("&Period=")) + Len("&Period=") + 1)
        URL_PATH = URL_PATH_1 + URL_PATH_2i + URL_PATH_3
        
        URL_PATH = Mid(URL_PATH, 1, InStr(LCase(URL_PATH), LCase("&GotoPage=")) + Len("&GotoPage=1") - 2)
        Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        Sleep 1000
        URL_COUNT = 0
        For R_EBAY = 1 To 29
            MNU_EBAY_MULTIPLE_PAGE.Caption = MNU_CAP + Format(R_EBAY, "00")
            ShellExecute Me.hWnd, "open", URL_PATH + Format(R_EBAY, "0"), vbNullString, vbNullString, conSwNormal
        
            Sleep 1000
            URL_COUNT = URL_COUNT + 1
        
            If URL_COUNT > 150 Then
                URL_COUNT = 0
                Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", vbNormalFocus
                Sleep 1000
            End If
        Next R_EBAY
    End If
        
    ' -------------------------
    ' 2017 IS 14 PAGE SET OF 50
    ' -------------------------
    ' NOW GONE -- FORGOT TO COLLECT END OF YEAR
    ' Thu 02-Jan-2020 13:04:40
    ' -------------------------
    
    If InStr(EBAY_YEAR_IN, "10,") Then
        ' ---------------------------------
        ' 2020 -- 10 SET OF THE MOST RECENT
        ' ---------------------------------
        URL_COUNT = 0
        For R_EBAY = 1 To 10
    
            MNU_EBAY_MULTIPLE_PAGE.Caption = MNU_CAP + Format(R_EBAY, "00")
            ShellExecute Me.hWnd, "open", URL_PATH + Format(R_EBAY, "0"), vbNullString, vbNullString, conSwNormal
    
            Sleep 800
            URL_COUNT = URL_COUNT + 1
    
            If URL_COUNT > 100 Then
                URL_COUNT = 0
                Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", vbNormalFocus
                Sleep 1000
            End If
    
        Next R_EBAY
            
    End If

    If InStr(EBAY_YEAR_IN, "10-YEAR_1") Then
        ' 10 SET OF YEAR BEFORE IS 2019  WHEN 2020
        ' ----------------------------------------
        ' DROP DOWN TO THE NEXT LOWER YEAR
        ' ----------------------------------------
        URL_PATH = URL_PATH_ORGINAL
        URL_PATH_1 = Mid(URL_PATH, 1, InStr(LCase(URL_PATH), LCase("&Period=")) + Len("&Period=") - 1)
        URL_PATH_2i = Mid(URL_PATH, InStr(LCase(URL_PATH), LCase("&Period=")) + Len("&Period="), 1)
        URL_PATH_2i = Trim(Str(Val(URL_PATH_2i) + 1))
        URL_PATH_3 = Mid(URL_PATH, InStr(LCase(URL_PATH), LCase("&Period=")) + Len("&Period=") + 1)
        URL_PATH = URL_PATH_1 + URL_PATH_2i + URL_PATH_3
        
        URL_PATH = Mid(URL_PATH, 1, InStr(LCase(URL_PATH), LCase("&GotoPage=")) + Len("&GotoPage=1") - 2)
        Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        Sleep 1000
        URL_COUNT = 0
        For R_EBAY = 1 To 10
            MNU_EBAY_MULTIPLE_PAGE.Caption = MNU_CAP + Format(R_EBAY, "00")
            ShellExecute Me.hWnd, "open", URL_PATH + Format(R_EBAY, "0"), vbNullString, vbNullString, conSwNormal
        
            Sleep 1000
            URL_COUNT = URL_COUNT + 1
        
            If URL_COUNT > 100 Then
                URL_COUNT = 0
                Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", vbNormalFocus
                Sleep 1000
            End If
        Next R_EBAY
    End If

End Sub

Private Sub MNU_EMPTY_CLIPBOARD_Click()

On Error Resume Next

Clipboard.Clear

End Sub

Private Sub MNU_EXE_LANUCHER_Click(Index As Integer)
    If Index = 1 Then Call MNU_TSF_CLICK
    If Index = 2 Then Call MNU_TSP_CLICK
    If Index = 3 Then Call MNU_TSS_CLICK
    If Index = 4 Then Call MNU_FLICKR_YAHOO_CLICK
    If Index = 5 Then Call MNU_FLCIKR_SYNC_CLICK
    If Index = 6 Then Call MNU_MULTI_MENU_CLICK
End Sub

Sub MNU_CODE_MENU_CLICK()

MsgBox "NOTHING HERE"

End Sub

Private Sub MNU_MIRROR_SEND_TO_OPERATING_SYSTEM_SET_MENU_CAPTION()
    'MNU_MIRROR_SEND_TO_OPERATING_SYSTEM.Caption = "MIRROR COPY OUR FAT32 SPECIAL SEND-TO FOLDER TO ANY OPERATING SYSTEM - SPECIAL FOLDER SEND TO"
    MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU.Item(8).Caption = "MIRROR COPY OUR FAT32 SPECIAL SEND-TO FOLDER TO ANY OPERATING SYSTEM - SPECIAL FOLDER SEND TO"
End Sub

Sub SET_MNU_SEND_TO_SYSTEM_FOLDER(FOLDER_SENDTO_SYSTEM)
    ' MNU_SEND_TO_SYSTEM_FOLDER.Caption = "SEND TO -- " + FOLDER_SENDTO_SYSTEM
    MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU.Item(3).Caption = "SEND TO -- " + FOLDER_SENDTO_SYSTEM
End Sub


Sub SET_MNU_SEND_TO_FAT32_FOLDER(FOLDER_SENDTO_FAT32_STORE)
    'MNU_SEND_TO_FAT32_FOLDER.Caption = "SEND TO -- " + FOLDER_SENDTO_FAT32_STORE
    MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU.Item(4).Caption = "SEND TO -- " + FOLDER_SENDTO_FAT32_STORE
End Sub


Sub SET_MNU_SPECIAL_FOLDER_DESKTOP_USER(FOLDER_SYSTEM_4)
    'MNU_SPECIAL_FOLDER_DESKTOP_USER.Caption = STRING_VAR + FOLDER_SYSTEM
    MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU.Item(7).Caption = FOLDER_SYSTEM_4
End Sub



Sub SET_MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC(FOLDER_SYSTEM_5)
    'MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC.Caption = STRING_VAR + FOLDER_SYSTEM
    MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU.Item(6).Caption = FOLDER_SYSTEM_5
End Sub


'Sub SET_MNU_SPECIAL_FOLDER_PIN_TO_START_MENU(FOLDER_SYSTEM)
'    ' MNU_SPECIAL_FOLDER_PIN_TO_START_MENU.Caption = "START MENU -- PIN TO -- " + FOLDER_SYSTEM
'    MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU.Item(5).Caption = "START MENU -- PIN TO -- " + FOLDER_SYSTEM
'End Sub
Sub SET_MNU_SPECIAL_FOLDER_PIN_TO_START_MENU(FOLDER_SYSTEM)
    ' MNU_SPECIAL_FOLDER_PIN_TO_START_MENU.Caption = "EXPLORER -- PIN TO START MENU -- " + FOLDER_SYSTEM
    MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU.Item(5).Caption = "EXPLORER -- PIN TO START MENU -- " + FOLDER_SYSTEM
End Sub

Sub SET_MNU_EXPLORER_ME_VB(FOLDER_SYSTEM_8)
    'Call SET_MNU_EXPLORER_ME_VB("EXPLORER ME_VB -- " + App.Path)
    MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU.Item(5).Caption = FOLDER_SYSTEM_8
End Sub


Private Sub MNU_EXPLORER_FOLDER_TOOL_SHOP_MENU_Click(Index As Integer)


If Index = 1 Then Call MNU_CODE_MENU_CLICK
' If Index = 2 Then Call MNU_TEST_STOP_ALL_TIMER_CLICK
If Index = 2 Then Call MNU_EXPLORER_ME_VB_Click
If Index = 3 Then Call MNU_SEND_TO_SYSTEM_FOLDER_Click
If Index = 4 Then Call MNU_SEND_TO_FAT32_FOLDER_Click
If Index = 5 Then Call MNU_SPECIAL_FOLDER_PIN_TO_START_MENU_Click
If Index = 6 Then Call MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC_Click
If Index = 7 Then Call MNU_SPECIAL_FOLDER_DESKTOP_USER_Click
If Index = 8 Then Call MNU_MIRROR_SEND_TO_OPERATING_SYSTEM_Click
If Index = 9 Then Call MNU_ABORT_SHUTDOWN_Click


End Sub

Private Sub MNU_EXPLORER_IMAGE_Click()
Call MNU_LAST_ART_PIC_SELECTOR_EXPLORER_Click
End Sub

Private Sub MNU_FILE_LOCATOR_IMAGE_Click()


'SHELL" D:\0 00 Art Loggers\# APP AND SCREEN\7-ASUS-GL522VW\FILE LOCATOR __ SavedCriteria.srf

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")

EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE " + EXECUTE_PARAM

EXECUTE_FILE_NAME = "D:\0 00 ART LOGGERS\# APP AND SCREEN\SavedCriteria IMAGE.srf"

If Dir(EXECUTE_FILE_NAME) = "" Then
    MsgBox "FILE NOT EXIST" + vbCrLf + vbCrLf + EXECUTE_FILE_NAME
    Exit Sub
End If

WSHShell.Run """" + EXECUTE_FILE_NAME + """"

Set WSHShell = Nothing

Rem ---------------------------------------------------------
Rem EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE /SELECT, C:\"
Rem VBS CODE WILL NOT RUN FROM NOTEPAD++
Rem ---------------------------------------------------------


End Sub

Private Sub MNU_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_Click()
' GATHER OTHER COMPUTER
MNU_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_VALUE = Not MNU_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_VALUE
If MNU_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_VALUE = -1 Then
    MNU_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_VALUE.Caption = "[__ GATHER OTHER COMPUTER = TRUE __]"
End If
If MNU_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_VALUE = 0 Then
    MNU_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_VALUE.Caption = "[__ GATHER OTHER COMPUTER __]"
End If

End Sub

Private Sub MNU_FOLDER_AUTO_CLIPBOARD_Click()

' FOLDERNAME_AUTO_SWAP

'    If FSO.FileExists(FOLDERNAME_AUTO_SWAP) = False Then
'        Call RECURSIVE_SEARCH_LAST_CLIPBOARD_ART_FOLDER
'    End If
    
    If FSO.FolderExists(FOLDERNAME_AUTO_SWAP) = False Then
        MsgBox "FILE NONE FOUND IMAGE AUTO LOGGER SEARCH VAR __CLIPBOARD_FILE" + vbCrLf + vbCrLf + FOLDERNAME_AUTO_SWAP
        Exit Sub
    End If
  
    Me.WindowState = vbMinimized
    ' Shell "Explorer.exe /SELECT," + FOLDERNAME_AUTO_SWAP, vbMaximizedFocus
    Shell "Explorer.exe /E," + FOLDERNAME_AUTO_SWAP, vbMaximizedFocus

End Sub

Private Sub Mnu_GET_LAST_CLIPPER_IN_TINYURL_Click()

Dim STRING_VAR_011
Dim STRING_VAR_012

STRING_VAR_00 = Mid(Text1.Text, Len(Text1.Text) - 3000)
XX_01 = Len(STRING_VAR_00)
STRING_VAR_011 = InStrRev(LCase(STRING_VAR_00), "http://tinyurl.com/", XX_01)
STRING_VAR_012 = InStrRev(LCase(STRING_VAR_00), "https://tinyurl.com/", XX_01)

If STRING_VAR_012 > STRING_VAR_011 Then STRING_VAR_01 = STRING_VAR_012
If STRING_VAR_011 > STRING_VAR_012 Then STRING_VAR_01 = STRING_VAR_011

' -------------------------------------------------
' NEW CHANGE TINY URL BEGIN TO USER --
' -------------------------------------------------
' https://tinyurl.com/ FROM http://tinyurl.com/
' -------------------------------------------------
' [ Monday 19:28:00 Pm_22 July 2019 ]
' -------------------------------------------------
If STRING_VAR_01 = 0 Then
    MsgBox "MUST BE PROBLEM -- NOT FIND" + vbCrLf + "http://tinyurl.com/" + vbCrLf + "OR" + vbCrLf + "https://tinyurl.com/"
    Exit Sub
End If

xx_02 = STRING_VAR_01
xx_03 = InStr(STRING_VAR_01, STRING_VAR_00, vbCrLf)
STRING_VAR_02 = InStrRev(LCase(STRING_VAR_00), "http", xx_02 - 1)
xx_04 = InStr(STRING_VAR_02, STRING_VAR_00, vbCrLf)
STRING_VAR_03 = InStrRev(LCase(STRING_VAR_00), vbCrLf, STRING_VAR_02 - 2)
If STRING_VAR_02 = 0 Then
    xx_02 = STRING_VAR_01
    STRING_VAR_02 = InStrRev(LCase(STRING_VAR_00), "www", xx_02 - 1)
End If
If STRING_VAR_02 > 0 Then
    STRING_VAR_05 = Mid(STRING_VAR_00, STRING_VAR_03 + 2, STRING_VAR_02 - STRING_VAR_03 - 4)
    ' STRING_VAR_07 = Mid(STRING_VAR_00, STRING_VAR_02, xx_04 - STRING_VAR_02)
    STRING_VAR_08 = Mid(STRING_VAR_00, STRING_VAR_01, xx_03 - STRING_VAR_01)
    STRING_VAR_05 = Replace(STRING_VAR_05, " - Google Drive", "")
    STRING_VAR_05 = Replace(STRING_VAR_05, "at master · Matthew-Lancaster/Matthew-Lancaster ", "")
    STRING_VAR_05 = Replace(STRING_VAR_05, "Matthew-Lancaster/Matthew-Lancaster", "Matthew-Lancaster")
    STRING_VAR_05 = Replace(STRING_VAR_05, "Matthew-Lancaster/", "")

    END_OUTPUT_STRING = ""
    END_OUTPUT_STRING = END_OUTPUT_STRING + "--" + vbCrLf
    END_OUTPUT_STRING = END_OUTPUT_STRING + STRING_VAR_05 + vbCrLf
    END_OUTPUT_STRING = END_OUTPUT_STRING + UCase(STRING_VAR_08) + vbCrLf
    END_OUTPUT_STRING = END_OUTPUT_STRING + "--" + vbCrLf

    END_OUTPUT_STRING = Replace(END_OUTPUT_STRING, "HTTPS:", "HTTP:")
End If

EXECUTE_TIMER_ENABLED = False
On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 100
Loop Until Err.Number = 0

Do
    Err.Clear
    Clipboard.SetText END_OUTPUT_STRING
    If Err.Number > 0 Then Sleep 100
Loop Until Err.Number = 0
EXECUTE_TIMER_ENABLED = True

Me.WindowState = vbMinimized
AD_DATE = 0
Beep

End Sub


Private Sub MNU_LAST_ART_PIC_SELECTOR_EXPLORER_Click()

    If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
        Call RECURSIVE_SEARCH_LAST_CLIPBOARD_ART_FOLDER
    End If
    
    If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
        MsgBox "File None Found Last Art Clipboard Search Var _ LAST_CLIPBOARD_FILE"
        Exit Sub
    End If
  
    Me.WindowState = vbMinimized
    Shell "Explorer.exe /SELECT," + LAST_CLIPBOARD_FILE, vbMaximizedFocus

End Sub


Private Sub MNU_LAST_GRAB_PRO_CAPS_Click()
    'AD$
    'MNU_LAST_GRAB_ALL_CAPS.Checked = False Then
    If MNU_LAST_GRAB_PRO_CAPS.Checked = False Then
        MNU_LAST_GRAB_PRO_CAPS.Checked = True
        GoTo EXIT1
    End If

    If MNU_LAST_GRAB_PRO_CAPS.Checked = True Then
        MNU_LAST_GRAB_PRO_CAPS.Checked = False
    End If
    
EXIT1:

Call Mnu_Fix1st_Letters_Click

' MNU_LAST_GRAB_PRO_CAPS
' MNU_PROCAPS

Me.WindowState = vbMinimized

Beep


End Sub


Private Sub MNU_LAST_GRAB_TEXT_INVERT_Click()

For r = 1 To Len(AD$)

    If Asc(Mid(AD$, r, 1)) < 91 Then
        Mid(AD$, r, 1) = LCase(Mid(AD$, r, 1))
    Else
        Mid(AD$, r, 1) = UCase(Mid(AD$, r, 1))
    End If

Next

EXECUTE_TIMER_ENABLED = False
    Clipboard.Clear
    Clipboard.SetText AD$
EXECUTE_TIMER_ENABLED = True

Me.WindowState = vbMinimized
AD_DATE = 0
Beep


End Sub

Private Sub MNU_ALWAYS_ON_TOP_Click()
    ' MENU FLIPPER
    If InStr(MNU_ALWAYS_ON_TOP.Caption, "ON TOP __ CORRECT") > 0 Then
        MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __]"
        AlwaysOnTop_MODE = False
        RESULT_API = NotAlwaysOnTop(Me.hWnd)
    Else
        MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __ CORRECT __]"
        AlwaysOnTop_MODE = True
        RESULT_API = AlwaysOnTop(Me.hWnd)
    End If
    Me.SetFocus

End Sub

Private Sub MNU_READ_CLIPBOARD_OF_HTTP_WEB_CHUNK_AND_OPEN_Click()

    Me.WindowState = vbMinimized
    On Error Resume Next
    
    Dim ARRAY_HTTP_1() As String
    Dim ARRAY_HTTP_2() As String
    ReDim Preserve ARRAY_HTTP_2(1000)
    ' ------------------------------------
    TEXT_STRING = Clipboard.GetText
    ' ------------------------------------
    YY_C = -1
    ARRAY_HTTP_1 = Split(TEXT_STRING, vbCrLf)
    For R_HTTP = 0 To UBound(ARRAY_HTTP_1)
        XX = InStr(UCase(ARRAY_HTTP_1(R_HTTP)), "HTTP")
        If XX > 0 Then
            YY_C = YY_C + 1
            ARRAY_HTTP_2(YY_C) = Mid(ARRAY_HTTP_1(R_HTTP), XX)
        End If
    Next
    
    ReDim Preserve ARRAY_HTTP_2(YY_C)
    Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", vbNormalFocus
    Sleep 1000

    COUNT_LIMITATION_OF_FACEBOOK = 30
    For R_HTTP = 0 To UBound(ARRAY_HTTP_2)
        ShellExecute Me.hWnd, "open", ARRAY_HTTP_2(R_HTTP), vbNullString, vbNullString, conSwNormal

        Sleep 1000
        URL_COUNT = URL_COUNT + 1

        If URL_COUNT > 40 Then
            URL_COUNT = 0
            Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", vbNormalFocus
            Sleep 1000
        End If
    COUNT_LIMITATION_OF_FACEBOOK = COUNT_LIMITATION_OF_FACEBOOK - 1
    If COUNT_LIMITATION_OF_FACEBOOK = 0 Then Exit For
    Next

End Sub

Private Sub MNU_REFORMAT_REPLACE_CRLF_DOUBLE_TO_SINGLE_02_Click()

Me.WindowState = vbMinimized


Dim EE As String, O_EE As String

I = MsgBox("REPLACE DOUBLE VBCRLF TO SINGLE YES NO", vbYesNo)
If I = vbNo Then Exit Sub

EE = AD$
'Do
    
    'EE = Replace(EE, vbCrLf + vbCrLf + vbCrLf, vbCrLf + vbCrLf)
    EE = Replace(EE, vbCrLf + vbCrLf, vbCrLf)

'    If O_EE = EE Then Exit Do
'    O_EE = EE
'Loop Until 1 = 2

AD$ = EE



EXECUTE_TIMER_ENABLED = False
On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 100
Loop Until Err.Number = 0

Do
    Err.Clear
    Clipboard.SetText AD$
    If Err.Number > 0 Then Sleep 100
Loop Until Err.Number = 0
Beep
EXECUTE_TIMER_ENABLED = True


End Sub



Private Sub MNU_REFORMAT_REPLACE_CRLF_DOUBLE_TO_SINGLE_Click()

Me.WindowState = vbMinimized


Dim EE As String, O_EE As String

I = MsgBox("REPLACE TRIPLE VBCRLF TO DOUBLE YES NO", vbYesNo)
If I = vbNo Then Exit Sub

EE = AD$
'Do
    
    EE = Replace(EE, vbCrLf + vbCrLf + vbCrLf, vbCrLf + vbCrLf)
    'EE = Replace(EE, vbCrLf + vbCrLf, vbCrLf)
'
'    If O_EE = EE Then Exit Do
'    O_EE = EE
'Loop Until 1 = 2

AD$ = EE



EXECUTE_TIMER_ENABLED = False
On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 100
Loop Until Err.Number = 0

Do
    Err.Clear
    Clipboard.SetText AD$
    If Err.Number > 0 Then Sleep 100
Loop Until Err.Number = 0
Beep
EXECUTE_TIMER_ENABLED = True


End Sub


Private Sub MNU_REFORMAT_REPLACE_DASH_TO_UNDERLINE_Click()

    Dim EE As String, O_EE As String
    Dim R1
    Dim TEST_BITAR
    EE = AD$
    Do
        TEST_BITAR = "-"
        For R1 = 1 To Len(TEST_BITAR)
            EE = Replace(EE, Mid(TEST_BITAR, R1, 1), "_")
        Next
        If O_EE = EE Then Exit Do
        O_EE = EE
    Loop Until 1 = 2
    
    AD$ = EE
    
    EXECUTE_TIMER_ENABLED = False
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    
    Do
        Err.Clear
        Clipboard.SetText AD$
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    Beep
    EXECUTE_TIMER_ENABLED = True
    
    Me.WindowState = vbMinimized



End Sub

Private Sub MNU_REFORMAT_REPLACE_SPACE_FOR_DASH_Click()

Dim EE As String, O_EE As String

EE = AD$
Do
    EE = Replace(EE, " ", "-")
    EE = Replace(EE, "--", "-")
    If O_EE = EE Then Exit Do
    O_EE = EE
Loop Until 1 = 2

AD$ = EE



EXECUTE_TIMER_ENABLED = False
On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 100
Loop Until Err.Number = 0

Do
    Err.Clear
    Clipboard.SetText AD$
    If Err.Number > 0 Then Sleep 100
Loop Until Err.Number = 0
Beep
EXECUTE_TIMER_ENABLED = True


Me.WindowState = vbMinimized


End Sub

Private Sub MNU_REFORMAT_REPLACE_SPACE_TO_NOTHING_Click()
    Dim EE As String, O_EE As String
    
    EE = AD$
    Do
        EE = Replace(EE, " ", "")
        If O_EE = EE Then Exit Do
        O_EE = EE
    Loop Until 1 = 2
    
    AD$ = EE
    
    
    
    EXECUTE_TIMER_ENABLED = False
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    
    Do
        Err.Clear
        Clipboard.SetText AD$
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    Beep
    EXECUTE_TIMER_ENABLED = True
    
    
    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_REFORMAT_REPLACE_SPACE_TO_UNDERLINE_Click()
    Dim EE As String, O_EE As String
    
    EE = AD$
    Do
        EE = Replace(EE, " ", "_")
        If O_EE = EE Then Exit Do
        O_EE = EE
    Loop Until 1 = 2
    
    AD$ = EE
    
    
    
    EXECUTE_TIMER_ENABLED = False
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    
    Do
        Err.Clear
        Clipboard.SetText AD$
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    Beep
    EXECUTE_TIMER_ENABLED = True
    
    
    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_REFORMAT_REPLACE_TAB_TO_4_SPACE_Click()

Dim EE As String, O_EE As String

EE = AD$
Do
    EE = Replace(EE, vbTab, "    ")
    If O_EE = EE Then Exit Do
    O_EE = EE
Loop Until 1 = 2

AD$ = EE



EXECUTE_TIMER_ENABLED = False
On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 100
Loop Until Err.Number = 0

Do
    Err.Clear
    Clipboard.SetText AD$
    If Err.Number > 0 Then Sleep 100
Loop Until Err.Number = 0
Beep
EXECUTE_TIMER_ENABLED = True




End Sub

Private Sub MNU_REPLACE_ALL_STRING_CHAR_TO_UNDERLINE_Click()

    Dim EE As String, O_EE As String
    Dim R1
    
    EE = AD$
    EE = String(Len(EE), "_")
    
    AD$ = EE
    
    EXECUTE_TIMER_ENABLED = False
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    
    Do
        Err.Clear
        Clipboard.SetText AD$
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    Beep
    EXECUTE_TIMER_ENABLED = True
    
    Me.WindowState = vbMinimized



End Sub

Private Sub MNU_REPLACE_CODER_STRING_UNDERLINE_AND_SPECIAL_CHAR_Click()

    Dim EE As String, O_EE As String
    Dim R1
    Dim TEST_BITAR
    EE = AD$
    Do
        TEST_BITAR = """!£$%^&*()[]-_+={};:'@#~,<.>/?|\`¬ "
        For R1 = 1 To Len(TEST_BITAR)
            EE = Replace(EE, Mid(TEST_BITAR, R1, 1), "_")
        Next
        If O_EE = EE Then Exit Do
        O_EE = EE
    Loop Until 1 = 2
    
    AD$ = EE
    
    EXECUTE_TIMER_ENABLED = False
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    
    Do
        Err.Clear
        Clipboard.SetText AD$
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    Beep
    EXECUTE_TIMER_ENABLED = True
    
    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_REPLACE_DASH_SPACE_TO_UNDERLINE_Click()
    
    Dim EE As String, O_EE As String
    Dim R1
    Dim TEST_BITAR
    EE = AD$
    Do
        TEST_BITAR = "- "
        For R1 = 1 To Len(TEST_BITAR)
            EE = Replace(EE, Mid(TEST_BITAR, R1, 1), "_")
        Next
        If O_EE = EE Then Exit Do
        O_EE = EE
    Loop Until 1 = 2
    
    AD$ = EE
    
    EXECUTE_TIMER_ENABLED = False
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    
    Do
        Err.Clear
        Clipboard.SetText AD$
        If Err.Number > 0 Then Sleep 100
    Loop Until Err.Number = 0
    Beep
    EXECUTE_TIMER_ENABLED = True
    
    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_RESET_FORM_Click()
    
    Call VBSCRIPT_CLOSE_AND_RELOAD
        
    Exit Sub
    
    EXIT_TRUE = True
    Unload Me
    DoEvents
    Reset
    EXIT_TRUE = False
    Load Form4_Reload_Form
    
    
End Sub

Private Sub MNU_SCROLL_DOWN_ARRAY_MNU_Click(Index As Integer)

If Index = 1 Then Call MNU_400_Click
If Index = 2 Then Call MNU_100_Click
If Index = 3 Then Call MNU_SCROLL_OFF_Click

End Sub

Private Sub MNU_SELECTOR_WHOLE_LINE_MODE_Click()

If InStr(MNU_SELECTOR_WHOLE_LINE_MODE.Caption, "RUN_FULL_LINE_SELECT=FALSE") Then
    MNU_SELECTOR_WHOLE_LINE_MODE.Caption = "[__ RUN_FULL_LINE_SELECT=TRUE __]"
    Exit Sub
End If
If InStr(MNU_SELECTOR_WHOLE_LINE_MODE.Caption, "RUN_FULL_LINE_SELECT=TRUE") Then
    MNU_SELECTOR_WHOLE_LINE_MODE.Caption = "[__ RUN_FULL_LINE_SELECT=FALSE __]"
    Exit Sub
End If
End Sub



Public Sub MNU_SPACE_DASH_TO_UNDERSCORE_Click()

    If MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER = "" Then
        On Error Resume Next
        Err.Clear
        MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER = Clipboard.GetText
        If Err.Number > 0 Then
            ' PUBLIC HERE -- NAME OF SUB WHOLE
            VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_SPACE_DASH_TO_UNDERSCORE_Click"
            TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
            Exit Sub
        End If
    End If
    
    Dim i2
    i2 = MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER
    i2 = Replace(i2, " ", "_")
    i2 = Replace(i2, "-", "_")
    i2 = Replace(i2, "-", "_")
    i2 = Replace(i2, ".", "_")
    i2 = Replace(i2, "?", "_")
    i2 = Replace(i2, "'", "_")
    i2 = Replace(i2, "`", "_")
    i2 = Replace(i2, """", "_")
    ' RENAME -- YYYY_MM_DD MMM_DDD HH_MM_SS__MA_.MP4 -- BATCH
    ' RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA__MP4____BATCH
    MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER = i2
    
    Me.WindowState = vbMinimized
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    Err.Clear
    Clipboard.SetText MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER
    If Err.Number > 0 Then
        ' PUBLIC HERE -- NAME OF SUB WHOLE
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_SPACE_DASH_TO_UNDERSCORE_Click"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
        Exit Sub
    End If

    MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER = ""
End Sub

Private Sub MNU_TEXT_CAPITAL_MODE_ON_Click()
    Call MNU_LAST_GRAB_ALL_CAPS_Click
End Sub

Private Sub Mnu_TEXT_Open_Recent_2_Click()
    Call Mnu_TEXT_Open_Recent_Click
End Sub

Private Sub MNU_TEXT_OPEN_STRIP_LOGGER_2_Click()
    Call Mnu_TEXT_StripLogg_Click
End Sub

Private Sub Text1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    
    STOP_RESIZE_WHILE_MOUSE_OR_KEY_DOWN = True

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    STOP_RESIZE_WHILE_MOUSE_OR_KEY_DOWN = False

End Sub

Private Sub Timer_1_Second_Timer()
        
    If A_TimeIdle + TimeSerial(0, 0, 5) < Now Then
        If Mnu_Exit_VAR <> "[__ Exit * __]" Then
            Mnu_Exit.Caption = "[__ Exit * __]"
            Mnu_Exit_VAR = "[__ Exit * __]"
        End If
    Else
        If Mnu_Exit_VAR <> "[__ Exit __]" Then
            Mnu_Exit.Caption = "[__ Exit __]"
            Mnu_Exit_VAR = "[__ Exit __]"
        End If
    End If


    If MIDNIGHT_TIMER = 0 Then
        MIDNIGHT_TIMER = Now + 1
        MIDNIGHT_TIMER = Int(MIDNIGHT_TIMER)
    End If
    
    ' ---------------------------------------------------------
    ' MIDNIGHT RESET AS BUG CHCKING
    ' SOMETHING IS GO INTO A SPIN DRIER LOOP AFTER LONG DAY AND
    ' AND LOOP IS DRIVE CPU PROCESSOR DOWN
    ' RESET A DAY SHOULD CURE FOR NOW
    ' AND LONG TIME FIND IT - SHUT CODE DOWN ONE AT A TIME FINDER
    ' ---------------------------------------------------------
    ' Sun 22-Sep-2019 04:31:20
    ' ---------------------------------------------------------
    If MIDNIGHT_TIMER < Now Then
        MIDNIGHT_TIMER = Now + 1
        MIDNIGHT_TIMER = Int(MIDNIGHT_TIMER)
        Call VBSCRIPT_CLOSE_AND_RELOAD
        ' ------------------------------------
        ' CALL MNU_RESET_FORM_Click
        ' STILL HERE FROM MENU CALL LINE ABOVE
        ' ------------------------------------
    End If

End Sub

Private Sub Timer_100_MILLISECOND_Timer()
    
    '---------------------------------------
    'IF MENU SET TO UPPER CASE ALL CLIPBOARD
    '---------------------------------------
    ' Call LAST_GRAB_ALL_CAPS_002_Click
    '---------------------------------------

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

' ---- WEIRD NAME GOT THERE
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

TT_VAR_API_RESET = "API " + Format(VAL_MINUTE_API, "00") '+ Test_Min_Var

Call SET_LABEL3

If MNU_API_RESET.Caption <> "API RESETER" Then
    MNU_API_RESET.Caption = "API RESETER"
End If
End Sub

Private Sub Mnu_API_Unload_Click()
    Exit Sub
    
    API_LOAD = False
    Unload FRMCLIPTEST01
End Sub


Private Sub Mnu_API_Unload_Reload_Click()
    Exit Sub
    
    API_LOAD = False
    Unload FRMCLIPTEST01
    DoEvents
    API_LOAD = True
    Load FRMCLIPTEST01
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

MMControl2_Counting_Var = 0

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
    If FSO.FileExists(FileSpec) Then
        GO3 = True
        Shell "Explorer.exe /select, " + FileSpec, vbNormalFocus
        Me.WindowState = vbMinimized
    End If
    
    '----------------------------------------
    'FOLDER SPEC
    '----------------------------------------
    If FSO.FolderExists(FileSpec) = True And GO3 = False Then
        GO3 = True
        'Shell "Explorer.exe " + FileSpec, vbNormalFocus
    End If

    If GO3 = False Then
        FileSpec_TMP = FindTreeLowerLevelWorkingExistPath(FileSpec)
        FileSpec = FileSpec_TMP
        If FSO.FolderExists(FileSpec) = True And GO3 = False Then
            GO3 = True
        End If
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
        Shell "Explorer.exe /e, " + FileSpec, vbNormalFocus
        Me.WindowState = vbMinimized
    End If

End Sub

Private Sub MNU_CLIPBOARD_TEST_Click()

FlagTestClipBoardRoutine = True
MsgBox "A Flag Has Been Set So Next Clipboarded Object You Do Should Also See a Message Of Working", vbMsgBoxSetForeground

End Sub

Private Sub MNU_COMPARE_Click()
    FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\#COMPARE 1 DUPE CHECKER.Txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
    FILENAME_COMPARE_1 = FILENAME_IN_USE_CHECK
    FR1 = FreeFile
    On Error Resume Next
    Open FILENAME_IN_USE_CHECK For Output As #FR1
        Print #FR1, COMPARE_EXE_1;
    Close #FR1

    FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\#COMPARE 2 DUPE CHECKER.Txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
    FILENAME_COMPARE_2 = FILENAME_IN_USE_CHECK
    FR1 = FreeFile
    On Error Resume Next
    Open FILENAME_IN_USE_CHECK For Output As #FR1
        Print #FR1, COMPARE_EXE_2;
    Close #FR1

    Beep
    Shell "C:\Program Files (x86)\WinMerge\WinMergeU.exe" + " """ + FILENAME_COMPARE_1 + """ """ + FILENAME_COMPARE_2 + """"
    Me.WindowState = vbMinimized

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

Public Sub MNU_COPY_Click()

    '----
    'Copy rich text box's text - Xtreme Visual Basic Talk
    'http://www.xtremevbtalk.com/general/15668-copy-rich-text-boxs-text.html
    '----
    
    MNU_COPY.Caption = "<*** COPY --" + Str(Text1.SelLength) + " ***>"
    If Text1.SelLength = 0 Then
        CLIPBOARD_RETURN_TIMER_ERROR_2 = 0
        Beep
        Exit Sub
    End If
    
    EXECUTE_TIMER_ENABLED = False
    
    On Error Resume Next
    
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then
        ' PUBLIC HERE -- NAME OF SUB WHOLE
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_COPY_Click"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
        Exit Sub
    End If
    Err.Clear
    SendMessage Text1.hWnd, WM_COPY, 0&, 0& 'COPY THE SELECT TEXT
    If Err.Number > 0 Then
        ' PUBLIC HERE -- NAME OF SUB WHOLE
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_COPY_Click"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
        Exit Sub
    End If
    
    EXECUTE_TIMER_ENABLED = True
    
    AD$ = Mid(Text1.Text, Text1.SelStart + 1, Text1.SelLength)

    Me.WindowState = vbMinimized
    Call COMPARE_FOR_EXE
    AD_DATE = 0
    
    Exit Sub
    
    ' HERE NOT REQUIRE ANY MORE WAS OLD ROUTINE FOR THROW CLIPBOARD ANOTHER
    ' MANY GOT HERE
    ' Timer_MNU_COPY_2.Enabled = True

End Sub


Public Sub MNU_COPY_2_Click()

    '----
    'Copy rich text box's text - Xtreme Visual Basic Talk
    'http://www.xtremevbtalk.com/general/15668-copy-rich-text-boxs-text.html
    '----
    
    MNU_COPY.Caption = "<*** COPY --" + Str(Text1.SelLength) + " ***>"
    
    If Text1.SelLength = 0 Then
        Beep
        Exit Sub
    End If

    Me.WindowState = vbMinimized
    
    EXECUTE_TIMER_ENABLED = False
    
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then
        ' PUBLIC HERE -- NAME OF SUB WHOLE
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_COPY_2_Click"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
        Exit Sub
    End If
    
    Err.Clear
    SendMessage Text1.hWnd, WM_COPY, 0&, 0& ' COPY SELECT TEXT
    If Err.Number > 0 Then
        ' PUBLIC HERE -- NAME OF SUB WHOLE
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_COPY_2_Click"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
        Exit Sub
    End If
    
    EXECUTE_TIMER_ENABLED = True
    
    AD$ = Mid(Text1.Text, Text1.SelStart + 1, Text1.SelLength)

    Call COMPARE_FOR_EXE
    AD_DATE = 0

    ' HERE NOT REQUIRE ANY MORE WAS OLD ROUTINE FOR THROW CLIPBOARD ANOTHER
    ' MANY GOT HERE
    ' Timer_MNU_COPY_2.Enabled = True

End Sub

Private Sub TIMER_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_Timer()
    TIMER_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER.Enabled = False
    Exit Sub
    
    Dim FILE_INFO_2 As String
    Dim VAR_4 As String
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
    
    On Error Resume Next
    FR1 = FreeFile
    VAR_1 = App.Path + "\# DATA\GETCOMPUTERNAME_&_GETUSERNAME_SET__"
    VAR_2 = VAR_1 + GetComputerName + "__" + GetUserName + "__" + "#NFS.TXT"
    
    File2.Path = App.Path + "\# DATA\"
    File2.Pattern = "GETCOMPUTERNAME*.TXT"
    ' ONCE RUN ALL FILE DATE POPULATE -- VARIABLE NOT DO AGAIN
    ' NOT OWN COMPUTER
    ' --------------------------------------------------------
    If ONCE_RUN_ALL_FILE_DATE_POPULATE__VARIABLE_NOT_DO_AGAIN__ = False Then
        For R_LOOP1 = 0 To File2.ListCount - 1
            ' MsgBox File2.List(R_LOOP1)
            FILE_INFO_2 = File2.Path + "\" + File2.List(R_LOOP1)
            Set F = FSO.GetFile(FILE_INFO_2)
            ADATE1 = F.DateLastModified
            If ADATE1 > Now - 10 Then
                FILE_INFO_44 = FILE_INFO_2
                FILE_INFO_44 = Replace(FILE_INFO_44, "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\GETCOMPUTERNAME_&_GETUSERNAME_SET__", "")
                FILE_INFO_44 = Replace(FILE_INFO_44, "__#NFS.TXT", "")
                FILE_INFO_44 = Replace(FILE_INFO_44, "__", "_")
                VAR_4 = "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\"
                VAR_4 = VAR_4 + FILE_INFO_44 + "\#OutClipChunck.Txt"
                ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_DATE(R_LOOP2) = ADATE1
                ' END UP WITH HERE FOR VAR4
                ' D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\4-ASUS-GL522VW_MATT 01\#OutClipChunck.Txt
                ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_NAME(R_LOOP2) = FILE_INFO_2
                CRC32_HEX = Hex(m_CRC.CalculateFile(VAR_4))
                ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_CRC32(R_LOOP2) = CRC32_HEX
            End If
        Next
    End If
    ONCE_RUN_ALL_FILE_DATE_POPULATE__VARIABLE_NOT_DO_AGAIN__ = True
    If Date <> "22/10/2020" Then
    If VALUE_START_STOP = False And IsIDE = True Then
        VALUE_START_STOP = True
        Stop ' WORK AROUND HERE
    End If
    End If
    
    ' CHECK FIND INFO IN ARRAY FROM FILE CONTROL FORM
    ' -----------------------------------------------
    For R_LOOP1 = 0 To File2.ListCount - 1
        ' MsgBox File2.List(R_LOOP1)
        FILE_INFO_2 = File2.Path + "\" + File2.List(R_LOOP1)
        Set F = FSO.GetFile(FILE_INFO_2)
        ADATE1 = F.DateLastModified
        ' MsgBox ADATE1
        If FILE_INFO_2 <> VAR_2 Then   ' NOT INCLUDE OWN
        If ADATE1 > Now - 10 Then
            ' ------------------------------------------------
            SET_OPERATION = False
            VAR_FIND_DATA = False
            For R_LOOP2 = 0 To 100
                INFO_2 = ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_NAME(R_LOOP2)
                ' IF INFO2 GOT NOTHNG TO DO WITH FILE SCRIPT ALTERNATE
                If INFO_2 <> "" Then
                If INFO_2 = FILE_INFO_2 Then
                    ' REDO STORE IT IF FIND AND CHECK BEFORE ON VALUE MATCH
                    ' -----------------------------------------
                    If ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_DATE(R_LOOP2) <> ADATE1 Then
                        ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_DATE(R_LOOP2) = ADATE1
                        ' GOT NEW DATE & ALSO HAVE TO CHECK CRC32
                        ' ---------------------------------------
                        FILE_INFO_44 = FILE_INFO_2
                        FILE_INFO_44 = Replace(FILE_INFO_44, "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\GETCOMPUTERNAME_&_GETUSERNAME_SET__", "")
                        FILE_INFO_44 = Replace(FILE_INFO_44, "__#NFS.TXT", "")
                        FILE_INFO_44 = Replace(FILE_INFO_44, "__", "_")
                        VAR_4 = "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\"
                        VAR_4 = VAR_4 + FILE_INFO_44 + "\#OutClipChunck.Txt"
                        ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_DATE(R_LOOP2) = ADATE1
                        ' END UP WITH HERE FOR VAR4
                        ' D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\4-ASUS-GL522VW_MATT 01\#OutClipChunck.Txt
                        ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_NAME(R_LOOP2) = FILE_INFO_2
                        CRC32_HEX = Hex(m_CRC.CalculateFile(VAR_4))
                        If ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_CRC32(R_LOOP2) <> CRC32_HEX Then
                            ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_CRC32(R_LOOP2) = CRC32_HEX
                            SET_OPERATION = True
                            SET_OPERATION_ARRAY_NUMERIC = R_LOOP2
                            Open VAR_2 For Output As #FR1
                                Print #FR1, "TIME DATE BE HOLD IN FILE SYSTEM AND USER TO DETECT THE RECENT COMPUTER FOR PURPOSE CLIPBOARD CHUNK SEND GATHER FROM OTHER COMPUTER"
                            Close #FR1
                        End If
                    End If
                    VAR_FIND_DATA = True
                End If
                End If
            Next
        End If
        End If
        ' ------------------------------------------------
        ' ---------------------------------------------------
        ' CAME THROUGH ALL ARRAY SEARCH
        ' SEARCH FILE AND GOT INFO DATE NEVER BEEN PUT ARRAY YET
        ' AND HERE
        ' ---------------------------------------------------
        If VAR_FIND_DATA = False Then
            For R_LOOP2 = 0 To 100
                If ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_NAME(R_LOOP2) = "" Then
                    ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_DATE(R_LOOP2) = ADATE1
                    ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_NAME(R_LOOP2) = FILE_INFO_2
                    ' ----
                    FILE_INFO_44 = FILE_INFO_2
                    FILE_INFO_44 = Replace(FILE_INFO_44, "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\GETCOMPUTERNAME_&_GETUSERNAME_SET__", "")
                    FILE_INFO_44 = Replace(FILE_INFO_44, "__#NFS.TXT", "")
                    FILE_INFO_44 = Replace(FILE_INFO_44, "__", "_")
                    VAR_4 = "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\"
                    VAR_4 = VAR_4 + FILE_INFO_44 + "\#OutClipChunck.Txt"
                    ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_DATE(R_LOOP2) = ADATE1
                    ' END UP WITH HERE FOR VAR4
                    ' D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\4-ASUS-GL522VW_MATT 01\#OutClipChunck.Txt
                    ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_NAME(R_LOOP2) = FILE_INFO_2
                    CRC32_HEX = Hex(m_CRC.CalculateFile(VAR_4))
                    If ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_CRC32(R_LOOP2) <> CRC32_HEX Then
                        ARRAY_GATHER_CLIPBOARD_FROM_OTHER_VALUE_CRC32(R_LOOP2) = CRC32_HEX
                        SET_OPERATION = True
                        SET_OPERATION_ARRAY_NUMERIC = R_LOOP2
                        Open VAR_2 For Output As #FR1
                            Print #FR1, "TIME DATE BE HOLD IN FILE SYSTEM AND USER TO DETECT THE RECENT COMPUTER FOR PURPOSE CLIPBOARD CHUNK SEND GATHER FROM OTHER COMPUTER"
                        Close #FR1
                    End If
                    Exit For
                End If
            Next
        End If
    Next
    
    
    '-----------------------------------------------
    ' AND FINALLY -- SET_OPERATION = True
    If FILE_INFO_2 <> VAR_2 Then   ' NOT INCLUDE OWN
    If SET_OPERATION = True Then
        FILE_INFO_2 = ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_NAME(SET_OPERATION_ARRAY_NUMERIC)
        ' IT READ THE FILE FOR HEX CRC32 AND BETTER DO IN ONE
        FR1 = FreeFile
        Open VAR_2 For Binary As #FR1
            AD$ = Input(LOF(FR1), FR1)
        Close #FR1
        ENTER_LARGE_IN_LOGGER = True
        Call Timer1_Timer
    End If
    End If


    ' DETECT ALL BUT OWN
    ' TIMER
    ' TIMER_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER -- DEFAULT ENABLE -- FALSE
    ' MENU MNU _BUTTON GO
    ' MNU_ARRAY_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER
    ' TIMER
    ' "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\4-ASUS-GL522VW_MATT 01\#OutClipChunck.Txt"
    ' Thu 28-May-2020 17:35:05
    ' Thu 28-May-2020 19:04:00 -- 1 HOUR 28 MINUTE
    On Error GoTo 0

End Sub

Private Sub TIMER_GATHER_CLIPBOARD_FROM_OTHER_COMPUTER_Timer()

End Sub

Private Sub Timer_MNU_COPY_2_Timer()

On Error Resume Next
OO_1 = Clipboard.GetFormat(vbCFText)
OO_2 = Clipboard.GetFormat(vbCFBitmap)

Clipboard_GetFormat_vbCFText_OO_1 = OO_1
Clipboard_GetFormat_vbCFBitmap_OO_2 = OO_2

If Err.Number > 0 Then
    GoTo EXIT_RETRY_AGAIN_ERROR
End If

Timer_MNU_COPY_2.Enabled = False


If CLIPBOARD_RETURN_TIMER_ERROR_1 = 1 Then
    CLIPBOARD_RETURN_TIMER_ERROR_1 = 0
    Call MNU_COPY_2_Click
End If

If CLIPBOARD_RETURN_TIMER_ERROR_2 = 1 Then
    CLIPBOARD_RETURN_TIMER_ERROR_2 = 0
    Call MNU_COPY_Click
End If

If CLIPBOARD_RETURN_TIMER_ERROR_3 = 1 Then
    CLIPBOARD_RETURN_TIMER_ERROR_3 = 0
    Call COMPARE_FOR_EXE
End If

If CLIPBOARD_RETURN_TIMER_ERROR_4 = 1 Then
    CLIPBOARD_RETURN_TIMER_ERROR_4 = 0
    Call Timer1_Timer
End If

If CLIPBOARD_RETURN_TIMER_ERROR_5 = 1 Then
    CLIPBOARD_RETURN_TIMER_ERROR_5 = 0
    Call CALC_DO_THE_ADDING
End If
If CLIPBOARD_RETURN_TIMER_ERROR_7 = 1 Then
    CLIPBOARD_RETURN_TIMER_ERROR_7 = 0
    Call FORM_LOADER_GET_CLIPBOARD_INFO
End If
If CLIPBOARD_RETURN_TIMER_ERROR_8 = 1 Then
    ' DO IN ROUTINE
    ' CLIPBOARD_RETURN_TIMER_ERROR_8 = 0
    Call FORM_LOADER_GET_CLIPBOARD_INFO
End If
If CLIPBOARD_RETURN_TIMER_ERROR_9 = 1 Then
    CLIPBOARD_RETURN_TIMER_ERROR_9 = 0
    Call FORM_LOADER_GET_CLIPBOARD_INFO
End If

If CLIPBOARD_RETURN_TIMER_ERROR_10 = 1 Then
    CLIPBOARD_RETURN_TIMER_ERROR_10 = 0
    Call NEXT_LOADING_FORM_LOADER_GET_CLIPBOARD_INFO
End If
If CLIPBOARD_RETURN_TIMER_ERROR_11 = 1 Then
    CLIPBOARD_RETURN_TIMER_ERROR_11 = 0
End If
If CLIPBOARD_RETURN_TIMER_ERROR_12 = 1 Then
    CLIPBOARD_RETURN_TIMER_ERROR_12 = 0
End If
If CLIPBOARD_RETURN_TIMER_ERROR_12 = 1 Then
    CLIPBOARD_RETURN_TIMER_ERROR_12 = 0
End If


' GetFormat_And_Display
' GET THIS ONE DONE

EXIT_RETRY_AGAIN_ERROR:


End Sub

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
    XPOS = InStr(LCase(URL_PATH), "http")
    
    
    
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
    
    
    'http://cpc.farnell.com/powerpax/sw3516-vi/ac-adaptor-3-3v-1-5a-regulated/dp/PW02342?output-voltage-output-1=3.3v&pf=111782151&anyFilterApplied=true&ddkey=http%3Aen-CPC%2FCPC_United_Kingdom%2Fw%2Fc%2Felectrical-lighting%2Fbatteries-power-supplies%2Fpower-supplies%2Fac-to-dc-converters%2Fac-adaptors
    CPC_OKAY_URL_PATH = False
    If InStr(LCase(URL_PATH), "http://cpc.farnell.com/") = 0 Then
        CPC_OKAY_URL_PATH = True
    End If
    If InStr(LCase(URL_PATH), "https://cpc.farnell.com/") = 0 Then
        CPC_OKAY_URL_PATH = True
    End If
    
    If InStr(LCase(URL_PATH), "cpc.farnell.com/webapp/wcs/stores/servlet/ProductDisplay") > 0 Then
        MsgBox "Wrong Type of CPC URL" + vbCrLf + vbCrLf + "Put the Order-Code in the Search Box Press Enter", vbOKOnly + vbMsgBoxSetForeground
        Exit Sub
    End If
    
    If CPC_OKAY_URL_PATH = False Then
        MsgBox "MUST BE A CPC URL PATH " + vbCrLf + "http(s)://cpc.farnell.com/" + vbCrLf + "GIVE RESULT WAS" + URL_PATH, vbMsgBoxSetForeground
        Exit Sub
    End If

    
    
    
    'http://cpc.farnell.com/pro-power/ppc100/foam-cleaner-400ml-aerosol/dp/SA01881
    
    'URL_PATH = "http://cpc.farnell.com/webapp/wcs/stores/servlet/ProductDisplay?catalogId=15002&langId=69&urlRequestType=Base&partNumber=HK01161&storeId=10180"
    
    'http://cpc.farnell.com/powerpax/sw3516-vi/ac-adaptor-3-3v-1-5a-regulated/dp/PW02342
    If InStr(LCase(URL_PATH), LCase("&storeId=")) > 0 Then
        URL_PATH = Mid(URL_PATH, 1, InStr(LCase(URL_PATH), LCase("&storeId=")) - 1)
    End If
    If InStr(LCase(URL_PATH), "?") > 0 Then
        URL_PATH = Mid(URL_PATH, 1, InStr(LCase(URL_PATH), "?") - 1)
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
    
        MNU_CPC_TOP_MENU.Caption = "CPC -- 00 40 80 90 -- " + Format(RCPC, "00")
    
        If Me.WindowState <> vbMinimized Then
            RESUME_GO_CPC = False
            Do
                Sleep 100
                Me.WindowState = vbMinimized
            
            Loop Until RESUME_GO_CPC = True Or Me.WindowState = vbMinimized
        End If
    
        IRCPC = Mid(Format(RCPC, "00"), 1, 1)
        ' 00 40 80 90
        If InStr("*0489", IRCPC) > 0 Or 1 = 2 Then
            ShellExecute Me.hWnd, "open", vFile + Format(RCPC, "00"), vbNullString, vbNullString, conSwNormal
        
            Sleep 1000
            URL_COUNT = URL_COUNT + 1
            
        
            If URL_COUNT > 20 Then
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

'MNU_404_CPC_PAGES.Visible = True
VAR_MNU_404_HITT_ONCE = True

End Sub


Private Sub MNU_404_CPC_PAGES_Click()

On Error Resume Next
EXECUTE_FILE_NAME = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-Autohotkeys Set AutoRun.ahk"
If FSO.FileExists(EXECUTE_FILE_NAME) Then
    Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + EXECUTE_FILE_NAME + """", WSCRIPT_DontShowWindow, DontWaitUntilFinished
    Set WSHShell = Nothing
End If
On Error GoTo 0
Beep
Me.WindowState = vbMinimized

Exit Sub

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
    ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
    



End Sub

Private Sub MNU_EBAY_FILTER_2222WILDCARD2222_Click()
      
    Exit Sub
      
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
    ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
    




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

'called at - ---Timer_SCREEN_SHOT_AUTO_Timer


' Shell "Explorer.exe " + FOLDERNAME2, vbNormalFocus

Call LAST_IMAGE("EXPLORER", FOLDERNAME_AUTO)

'Shell "EXPLORER.EXE /SELECT, " + FOLDERNAME_AUTO, vbNormalFocus

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
' HERE NOT MENU OPTION ANYMORE
' ---------------------------------------------

'AUTO SCREEN IMAGE CAPTURE

'called at - ---Timer_SCREEN_SHOT_AUTO_Timer

Call LAST_IMAGE("EXPLORER", FOLDERNAME_AUTO)

'Shell "Explorer.exe " + FOLDERNAME1, vbNormalFocus

End Sub

Private Sub Mnu_IVIEW_Screen_Capture_Click()

End Sub

Sub LAST_IMAGE(VAR1, VAR2)

'XGO = False
'If GetComputerName = "1-ASUS-EEE" Then XGO = True
'If GetComputerName = "1-ASUS-X5DIJ" Then XGO = True


'GIVE UP ON THIS A BIT

' VAR2 = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\AUTO_Form_Shot\"

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

If FSO.FileExists(VAR2) = True Then
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
'Set F = FSO.getfile((Filespec1))
'ADATE1 = F.datelastmodified
'
'ScanPath.lblCount7 = ""
'ScanPath.ListView1.ListItems.Clear
'
'ScanPath.TxtPath = "D:\0 00 Art Loggers\# APP AND SCREEN\"
'Call ScanPath.cmdScan_Click


'Filespec2 = ScanPath.lblCount7
'If Filespec2 <> "" Then
'    Set F = FSO.getfile((Filespec2))
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

'called at - ---Timer_SCREEN_SHOT_AUTO_Timer

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

'Call LAST_IMAGE("IVIEW", "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\Hot-Key-App-Shots\") ' = I_VIEW
'Call LAST_IMAGE("IVIEW", "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName)  ' = I_VIEW


Call RECURSIVE_SEARCH_LAST_CLIPBOARD_ART_FOLDER

    If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
    MsgBox "File None Found Last Art Clipboard Search Var _ LAST_CLIPBOARD_FILE"
    Exit Sub
End If

If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
    MsgBox "File Found Not Valid Last Art Clipboard Search Var _ LAST_CLIPBOARD_FILE"
    Exit Sub
End If

'    Shell "Explorer.exe /SELECT, """ + LAST_CLIPBOARD_FILE + """", vbMaximizedFocus

If IRFANVIEW_PATH <> "" Then
    Shell IRFANVIEW_PATH + " """ + LAST_CLIPBOARD_FILE + """ /fs /silent", vbMaximizedFocus
Else
    Me.WindowState = vbNormal
    MsgBox "IRFANVIEW_PATH VAR -- PATH NOT FOUND FOR FILE" + vbCrLf + "NOT INSTALED AT EXPECTED LOCATION " + IRFANVIEW_PATH3 + vbCrLf + "OR" + vbCrLf + IRFANVIEW_PATH2, vbMsgBoxSetForeground
End If

Me.WindowState = vbMinimized


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
'Set F = FSO.getfile((Filespec1))
'ADATE1 = F.datelastmodified
'
'ScanPath.lblCount7 = ""
'ScanPath.ListView1.ListItems.Clear
'
'ScanPath.TxtPath = "D:\0 00 Art Loggers\# APP AND SCREEN\"
'Call ScanPath.cmdScan_Click


'Filespec2 = ScanPath.lblCount7
'If Filespec2 <> "" Then
'    Set F = FSO.getfile((Filespec2))
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



Private Sub Mnu_LAST_CLIP_TIME_Click()
'Mnu_LAST_CLIP_TIME
End Sub

Private Sub Mnu_Last_Clipboard_Timer_Click()
'Mnu_Last_Clipboard_Timer
End Sub

Private Sub MNU_LAST_GRAB_CAPS_TOP_BAR_Click()

    Call MNU_LAST_GRAB_CAPS_Click

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
    
    If (FSO.FileExists(A1 + B1) And FSO.FileExists(FOLDER_SENDTO_SYSTEM + B1)) = False Then
        FSO.CopyFile A1 + B1, FOLDER_SENDTO_SYSTEM + B1
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

Private Sub MNU_MULTI_MENU_CLICK()
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
            
    If MMControl1.Filename = "" Then
        Call RESET_SETUP_SOUND_FILE("NOTSOUND")
    End If
    MMControl1.Command = "prev"
    MMControl1.Command = "Play"
            
    If Mnu_SoundOn.Checked = False Then
        MsgBox "Play Sound in Program Running Set of Off", vbMsgBoxSetForeground
    End If
End Sub

Private Sub Mnu_Play_Sound_2_Click()
    If MMControl2.Filename = "" Then
        Call RESET_SETUP_SOUND_FILE("NOTSOUND")
    End If
    MMControl2.Command = "prev"
    MMControl2.Command = "Play"
            
    If Mnu_SoundOn.Checked = False Then
        MsgBox "Play Sound in Program Running Set of Off", vbMsgBoxSetForeground
    End If
End Sub

Private Sub MNU_PROCAPS_TOP_BAR_Click()

Call Mnu_Fix1st_Letters_Click

Me.WindowState = vbMinimized

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
                        If FSO.FileExists(PATH_SHORT) = True Then
                            On Error Resume Next
                            Err.Clear
                            FSO.DeleteFile PATH_SHORT
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

'THIS GET RUN HERE AT RESET
'AND AT ONE CLICK SELECT

'-----------------------------
'MODIFY_SOUND_SELECTION = True
'Call SETUP_SOUND_FILE("")
'-----------------------------
Call RESET_SETUP_SOUND_FILE("")

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
    frm_MSGBOX.Show
End Sub

Sub RECURSIVE_SEARCH_LAST_CLIPBOARD_ART_FOLDER()

    Dim FS, F, F1, fc, S
    Dim STORE_DATE
    Dim objFiles As Files, objFile As File
    Dim objFolders As Folders, objFolder As FOLDER
    Dim fld As FOLDER

    Call RecursiveSearch_Most_Recent_01("D:\0 00 Art Loggers\# APP AND SCREEN\" + GetComputerName)
    LAST_CLIPBOARD_FILE = RecursiveSearch_Most_Recent_VAR

End Sub

Private Sub MNU_SHOW_IMAGE_1_Click()
    
    'Beep
    'Call MNU_LAST_ART_PIC_IVIEW_Click
    
        If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
        Call RECURSIVE_SEARCH_LAST_CLIPBOARD_ART_FOLDER
    End If
    
        If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
        MsgBox "File None Found Last Art Clipboard Search Var _ LAST_CLIPBOARD_FILE"
        Exit Sub
    End If
    
    If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
        MsgBox "File Found Not Valid Last Art Clipboard Search Var _ LAST_CLIPBOARD_FILE"
        Exit Sub
    End If
    
    Clipboard.Clear
    Clipboard.SetText LAST_CLIPBOARD_FILE
    
    'Call LAST_IMAGE("IVIEW", "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName)  ' = I_VIEW
    
    If IRFANVIEW_PATH <> "" Then
        Shell IRFANVIEW_PATH + " """ + LAST_CLIPBOARD_FILE + """ /fs /silent", vbMaximizedFocus
'        Shell "Explorer.exe /SELECT, """ + LAST_CLIPBOARD_FILE + """", vbMaximizedFocus
'    Beep
    Else
        Me.WindowState = vbNormal
        MsgBox "IRFANVIEW_PATH VAR -- PATH NOT FOUND FOR FILE" + vbCrLf + "NOT INSTALED AT EXPECTED LOCATION " + IRFANVIEW_PATH3 + vbCrLf + "OR" + vbCrLf + IRFANVIEW_PATH2, vbMsgBoxSetForeground
    End If

    Me.WindowState = vbMinimized

End Sub

Sub RecursiveSearch_Most_Recent_01(objStartFolder)
    RecursiveSearch_Most_Recent_STORE_DATE = 0
    Set FOLDER = FSO.GetFolder(objStartFolder)

    Call RecursiveSearch_Most_Recent_02(FOLDER)

End Sub


Sub RecursiveSearch_Most_Recent_02(FOLDER)
    Dim process
    Dim EXTENSION

    Dim SubFolder
    For Each SubFolder In FOLDER.SubFolders
        RecursiveSearch_Most_Recent_02 SubFolder
    Next
    Dim File
    For Each File In FOLDER.Files
            ' Operate on each file
            process = True
            If File.Name = "Thumbs.db" Then process = False
            EXTENSION = " "
            If InStrRev(File.Name, ".") > 0 Then
                EXTENSION = UCase(Mid(File.Name, InStrRev(File.Name, ".")))
            End If
            If InStr(EXTENSION, ".TXT") > 0 Then process = False
            If InStr(File.Path, "\_gsdata_") > 0 Then process = False
            If process = True Then
                If File.DateLastModified > RecursiveSearch_Most_Recent_STORE_DATE Then
                    RecursiveSearch_Most_Recent_STORE_DATE = File.DateLastModified
                    RecursiveSearch_Most_Recent_VAR = File.Path
                End If
            End If
    Next

'REFERENCE
'----
'vb6 fso get sub folder and files - Google Search
'https://www.google.co.uk/search?q=vb6+fso+get+sub+folder+and+files&oq=vb6+fso+get+sub+folder+and+files&aqs=chrome..69i57.11696j0j7&sourceid=chrome&ie=UTF-8
'--------
'Loop through folders/subfolders-VBForums
'http://www.vbforums.com/showthread.php?613400-Loop-through-folders-subfolders
'----
End Sub

Private Sub MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC_Click()
    
    'GetSpecialfolder_DEBUG
    
    STRING_VAR = "DESKTOP PUBLIC -- "
    '------------
    FOLDER_SYSTEM = GetSpecialfolder(25) + "\"
    
'    MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC.Caption = STRING_VAR + FOLDER_SYSTEM
    Call SET_MNU_SPECIAL_FOLDER_DESKTOP_PUBLIC(STRING_VAR + FOLDER_SYSTEM)
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
    'MNU_SPECIAL_FOLDER_DESKTOP_USER.Caption = STRING_VAR + FOLDER_SYSTEM
    Call SET_MNU_SPECIAL_FOLDER_DESKTOP_USER(STRING_VAR + FOLDER_SYSTEM)
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
    ' MNU_SPECIAL_FOLDER_PIN_TO_START_MENU.Caption = "START MENU -- PIN TO -- " + FOLDER_SYSTEM
    Call SET_MNU_SPECIAL_FOLDER_PIN_TO_START_MENU(FOLDER_SYSTEM)
    
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

'MNU_SPECIAL_FOLDER_PIN_TO_START_MENU.Caption = "EXPLORER -- PIN TO START MENU -- " + FOLDER_SYSTEM
Call SET_MNU_SPECIAL_FOLDER_PIN_TO_START_MENU(FOLDER_SYSTEM)

'MNU_PIN_TO_START_MENU
'C:\Users\MATT 01\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu
End Sub

       
        



Sub SPECIAL_FOLDER_SEND_TO()

FOLDER_SENDTO_SYSTEM = GetSpecialfolder(CSIDL_SENDTO) + "\"
'--------
'01 OF 02
'--------

Call SET_MNU_SEND_TO_SYSTEM_FOLDER(FOLDER_SENDTO_SYSTEM)
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
Call SET_MNU_SEND_TO_FAT32_FOLDER(FOLDER_SENDTO_FAT32_STORE)

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
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Caption = "SCROLL BACK TO BOTTOM ON TIMER.ENABLED=FALSE"
    MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM_PRESS.Caption = "SCROLL BACK TO BOTTOM ON TIMER.ENABLED=FALSE"
    MNU_SCROLL_BOTTOM_OFF__WANT_ON_ASK_PRESS.Caption = "SCROLL BOTTOM=ON"
Else
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
    
    ' STOP_RESIZE_WHILE_MOUSE_OR_KEY_DOWN = True
    
    
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

'If IsIDE = True Then
'    If KEYBOARD_OR_MOUSE_ACTIVE_3_MIN > Now Then
'        Timer_API_OKAY_COLOUR.Interval = 4000
'    Else
'        Timer_API_OKAY_COLOUR.Interval = 1
'    End If
'End If

TIMER_INTERVAL = 20
If Timer_API_OKAY_COLOUR.Interval <> TIMER_INTERVAL Then Timer_API_OKAY_COLOUR.Interval = TIMER_INTERVAL





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

WTrue2 = WTrue2 + TW2
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




Private Sub TIMER_Clipboard_Clear_Timer()
On Error Resume Next
Clipboard_Clear_SUPPRESS_MESSAGE = True
Do
    Err.Clear
    Clipboard.Clear
Loop Until Err.Number = 0
TIMER_Clipboard_Clear_VAR = True
TIMER_Clipboard_Clear.Enabled = False
End Sub

Private Sub TIMER_Clipboard_Set_Text_Timer()
If TIMER_Clipboard_Set_Text_VAR = "" Then
    TIMER_Clipboard_Set_Text.Enabled = False
    Exit Sub
End If

X_COUNT_VAR = 0
Do
    On Error Resume Next
    Err.Clear
    DoEvents
    Clipboard.SetText TIMER_Clipboard_Set_Text_VAR
    DoEvents
    XZ = Err.Number
    If XZ > 0 Then Sleep 100
    On Error GoTo 0
    X_COUNT_VAR = X_COUNT_VAR + 1
Loop Until XZ = 0 Or X_COUNT_VAR > 100

If X_COUNT_VAR > 100 Then
    'MsgBox "ERROR UNABLE TO SET CLIPBOARD TO" + vbCrLf + vbCrLf + TIMER_Clipboard_Set_Text_VAR, vbMsgBoxSetForeground
    MSGBOX2 = "ERROR UNABLE TO SET CLIPBOARD TO" + vbCrLf + vbCrLf + TIMER_Clipboard_Set_Text_VAR
    frm_MSGBOX.Timer1.Enabled = True
End If

TIMER_Clipboard_Set_Text_VAR = ""
TIMER_Clipboard_Set_Text.Enabled = False
End Sub

Private Sub Timer_GET_KEY_ESC_MINIMIZE_Timer()
If GetForegroundWindow <> Me.hWnd Then
    Exit Sub
End If

If GetAsyncKeyState(27) < 0 Then Me.WindowState = vbMinimized

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
    
    ITXT = "Last Clipboard " + Trim(ClipFormatDesc2) + " -- "
    ITXT = ITXT + Format(DateDiff("n", TIME_LAST_CLIPBOARD, Now), "0")
    Test_Min_Var = " Minutes"
    If DateDiff("s", TIME_LAST_CLIPBOARD, Now) < 60 Then
        ITXT = ITXT + Format(DateDiff("s", TIME_LAST_CLIPBOARD, Now), "00")
        Test_Min_Var = " Seconds"
    End If
    LAST_CLIP_TIME_INFO = ITXT + Test_Min_Var
    
    'Mnu_LAST_CLIP_TIME_Second.Caption = LAST_CLIP_TIME_INFO + " Away"
    LAST_CLIP_TIME_INFO = Replace(LAST_CLIP_TIME_INFO, "Last Clipboard", "")
    TT_VAR_LAST_CLIP_TIME_Second = "Last Clipper" + LAST_CLIP_TIME_INFO + " Away"
    Call SET_LABEL3

End If

If TIME_LAST_CLIPBOARD = 0 And TIME_LAST_CLIPBOARD_Timer1 > 0 Then
    
    ITXT = "Last Clipboard " + Trim(ClipFormatDesc2) + " -- "
    ITXT = ITXT + Format(DateDiff("n", TIME_LAST_CLIPBOARD_Timer1, Now) + "00")
    Test_Min_Var = " Minutes"
    If TIME_LAST_CLIPBOARD = 0 Then TIME_LAST_CLIPBOARD = Now
    If DateDiff("s", TIME_LAST_CLIPBOARD, Now) < 60 Then
        ITXT = ITXT + Format(DateDiff("s", TIME_LAST_CLIPBOARD_Timer1, Now), "00")
        Test_Min_Var = " Seconds"
    End If
    LAST_CLIP_TIME_INFO = ITXT + Test_Min_Var
    
    
    'Mnu_LAST_CLIP_TIME_Second.Caption = "Last Clipper" + Str(DateDiff("s", TIME_LAST_CLIPBOARD_Timer1, Now)) + " Start-up In Buffer"
     LAST_CLIP_TIME_INFO = Replace(LAST_CLIP_TIME_INFO, "Last Clipboard", "")
    TT_VAR_LAST_CLIP_TIME_Second = "Last Clipper" + LAST_CLIP_TIME_INFO + " Start-up In Buffer"
    Call SET_LABEL3

End If

'----------------------
If TIME_LAST_CLIPBOARD_ERROR_MSG > 0 Then
    
'    ITXT = "Show Error Script && Last Error = "
'    TIME_SET = Trim(Str(DateDiff("s", TIME_LAST_CLIPBOARD_ERROR_MSG, Now)))
'    TIME_SET = TIME_SET + " Second"
'    If DateDiff("s", TIME_LAST_CLIPBOARD_ERROR_MSG, Now) > 60 Then
'        TIME_SET = ""
'        TIME_SET = Trim(Str(DateDiff("n", TIME_LAST_CLIPBOARD_ERROR_MSG, Now)))
'        TIME_SET = TIME_SET + " Minute"
'    End If
'    ITXT = ITXT + TIME_SET
'    LAST_CLIP_TIME_INFO = ITXT
'    If Mnu_Show_Error_Script_Frm_Msgbox.Caption <> LAST_CLIP_TIME_INFO Then
        'Mnu_Show_Error_Script_Frm_Msgbox.Caption = LAST_CLIP_TIME_INFO
        Call frm_MSGBOX.UPDATE_MNU_ERROR
'    End If
    
'Else
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
    OMenu_Height = Menu_Height

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
            On Error Resume Next
            frm_MSGBOX.Timer1.Enabled = True
            
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
            frm_MSGBOX.Timer1.Enabled = True

            'FORM_STAY_UP_WITH_MSGBOX = False
                
        End If
    End If
End If

End Sub

Private Sub TIMER_OutClipChunck_Txt_Timer()

    '10MS TIMER

    FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\#OutClipChunck.Txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
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

Private Sub Timer_RESET_API_CLIPPER_Timer()

    If FRMCLIPTEST01_ENABLE = False Then Exit Sub

    Exit Sub
    If API_LOAD = True Then
        API_LOAD = False
        Unload FRMCLIPTEST01
    End If
    DoEvents
    API_LOAD = True
    Load FRMCLIPTEST01
    Timer_RESET_API_CLIPPER.Enabled = False

    'FRMCLIPTEST01.ctlClipboard1.EndClipboardViewer
    'FRMCLIPTEST01.ctlClipboard1.StartClipboardViewer
    'FRMCLIPTEST01.Timer_RESET_API_CLIPPER

End Sub

Private Sub Timer_SET_LABEL3_INFO_BAR_Timer()
SET_LABEL3_INFO_BAR_READY_TO_GO_VAR = True
End Sub

Private Sub TIMER_VB_PROJECT_DATE_Timer()

Exit Sub

' HERE CALL FROM FORM_LOAD ONLY USE ONCE
' PROBERBLY FIND ANOTHER ONE IN PROJECT CHECK DATE FORM
'
TIMER_VB_PROJECT_DATE.Enabled = False

Debug.Print "CALL PROJECT_CHECK_DATE.VB_PROJECT_CHECKDATE(""FORM LOAD"")"

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
If IsIDE = False Or 1 = 1 Then
    'Call Project_Check_Date.VB_PROJECT_CHECKDATE("FORM LOAD")
    Load Project_Check_Date
End If

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Debug.Print "CALL PROJECT_CHECK_DATE"

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
    
    If STOP_RESIZE_WHILE_MOUSE_OR_KEY_DOWN = True Then Exit Sub

    If MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = False Then Exit Sub

    Text1.Enabled = False
    Text1.SelStart = 1
    Text1.Enabled = True
    Text1.SelStart = Len(Text1)
    
    Timer_FORM_RESIZE.Enabled = False

End Sub


Private Sub Form_Resize()

'Exit Sub

If DONT_RESIZE_WHILE_SETUP = True Then Exit Sub
If STOP_RESIZE_WHILE_MOUSE_OR_KEY_DOWN = True Then Exit Sub

If FORM_LOAD_SIZE_NOT_SET_YET = True Then
    If Me.WindowState <> vbMinimized Then
        FORM_LOAD_SIZE_NOT_SET_YET = False
        DONT_RESIZE_WHILE_SETUP = True
        
        Me.WindowState = vbMaximized
        
        DoEvents
        DONT_RESIZE_WHILE_SETUP = False
    End If
End If

Timer_MENU_HEIGHT_CHANGED.Enabled = True

'MAYBE SOMETIMES WINDOW IS MIN WHILE TAKE A CHANGE
'AT CLIPBOARD
'THINK NOT
'If Me.WindowState = vbMinimized Then RRR = Now - TimeSerial(0, 0, 3)

HHH = Now + TimeSerial(0, 2, 0) ' vbMinimized WINDOW

GMCL_I = GET_MENU_CONTROL_LENGTH(I)
If GMCL_I + Menu_Height + Me.Top + Me.Left + Me.Width + Me.Height = OLTLWH_3 Then Exit Sub
OLTLWH_3 = Menu_Height + Me.Top + Me.Left + Me.Width + Me.Height + GMCL_I
O_GMCL_I = GMCL_I

Call SETUP_DIMENSIONS_INNER_FORM

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
    'i = SetForegroundWindow(Me.hWnd)
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


Function GET_PARENT(I)
   'I = GetForegroundWindow
    GET_PARENT = I
    Do While I <> 0
         GET_PARENT = I
         I = GetParent(I)
      Loop
End Function


Sub SETUP_DIMENSIONS_INNER_FORM()



If Menu_Height = O_Menu_Height_1 Then Exit Sub
O_Menu_Height_1 = Menu_Height

FindCursor Nx, Ny
mWnd = WindowFromPoint(Nx, Ny)
If GET_PARENT(mWnd) = Me.hWnd Then
        FLAG_WW_VAR = True
    Else
        FLAG_WW_VAR = False
End If
If FLAG_WW_VAR <> O_FLAG_WW_VAR Then OLTLWH_1 = 0
O_FLAG_WW_VAR = FLAG_WW_VAR

If Menu_Height + Me.Top + Me.Left + Me.Width + Me.Height = OLTLWH_1 Then Exit Sub
OLTLWH_1 = Menu_Height + Me.Top + Me.Left + Me.Width + Me.Height

If STOP_RESIZE_WHILE_MOUSE_OR_KEY_DOWN = True Then Exit Sub


'DoEvents   ' Yield for other processing.
'Line Method Example

'Line1.BorderWidth = 40
'Border Line On the Edge

Label1.Top = 0
Label1.Width = Form1.Width - 70
Call HEIGHT_BAR_TOP_ROUTINE


Text1.Left = 0 - 18

Text1.Top = Label3.Top + Label3.Height

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
' Call HEIGHT_BAR_TOP_ROUTINE


'MAKE FORM TALLER OR TEXT BOX
'FORM IS PRIOITY OVER TEXT BOX

XXDD = Form1.Height - (Menu_Height * Screen.TwipsPerPixelY) - 500 - (Label3.Top + Label3.Height)
If XXDD > 0 Then
    If XXDD - HEIGHT_ADJUST > 0 Then
        Text1.Height = XXDD - HEIGHT_ADJUST
    End If
End If


If Menu_Height + Me.Top + Me.Left + Me.Width + Me.Height = OLTLWH_2 Then Exit Sub
OLTLWH_2 = Menu_Height + Me.Top + Me.Left + Me.Width + Me.Height

'SCROLL TO BOTTOM PROCEDURE
'--------------------------
Timer_FORM_RESIZE.Enabled = False
Timer_FORM_RESIZE.Interval = 1
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
        Form1.Width = SCREEN_WIDTH_SPACE '- 200
        Me.Width = Form1.Width
        Form1.Height = SCREEN_HEIGHT_SPACE - 900
        Form1.Height = SCREEN_HEIGHT_SPACE '- 200
        Me.Height = Form1.Height
        'THIS THE FORM HEIGHT
        'THE BOX INSIDE IS ADJUST AFTER
        
        'Form1.Left = (Screen.Width - Me.Width) / 2
        Form1.Left = X4 + (SCREEN_WIDTH_SPACE - Me.Width) / 2
        Me.Left = Form1.Left
        Form1.Top = Y2 + (SCREEN_HEIGHT_SPACE - Me.Height) / 2
        Me.Top = Form1.Top
    
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

'On Error Resume Next
'For Each Form In Forms
'    Err.Clear
'    If Form.EXIT_TRUE = True Then
'        'Form.Name
'        If Err.Number = 0 Then
'            Me.EXIT_TRUE = True
'        End If
'    End If
'Next Form
'On Error GoTo 0



If IsIDE = False Then
    If Me.WindowState <> vbMinimized And Me.EXIT_TRUE = False Then
        Me.WindowState = vbMinimized
        Cancel = True
        Exit Sub
    End If
End If
'Unload FrmJoy
'Unload frmSender

'----------------------------
MMControl2.Command = "close"
MMControl1.Command = "close"
'----------------------------

If TIMER_OutClipChunck_Txt.Enabled = True Then
    FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\#OutClipChunck.Txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
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

End

If Me.EXIT_TRUE = True Then Cancel = False

'For Each Form In Forms
'    Unload Form
'    Set Form = Nothing
'Next Form

Unload Form2_ANY_SPECIAL_FOLDER
Unload Form3_Extra
Unload Form4_Reload_Form
Unload frm_MSGBOX
'Unload FRMCLIPTEST01
'Unload FRMCLIPTEST02
Unload frmJoy
Unload frmListMenu
Unload frmSender
Unload Project_Check_Date
Unload ScanPath
Unload SCREEN_CAP
Unload UNLOAD_FORM_SAFE


If API_LOAD = True Then
    API_LOAD = False
    Unload FRMCLIPTEST01
    API_LOAD = False
End If

If API_LOAD = True Then
    End
End If

'EXIT_TRUE = True
'Dim Control
''SET ALL TIMERS IN ALL FORMS ENABLED=FALSE
'On Error Resume Next
'    For i = 0 To Forms.Count - 1
'        For Each Control In Forms(i).Controls
'            If InStr(UCase(Control.Name), "TIMER") > 0 Then
'                'Debug.Print Control.Name
'                Control.Enabled = False
'            End If
'        Next
'    Next i
'On Error GoTo 0
'
'On Error Resume Next
'For Each Form In Forms
'    Err.Clear
'    If Form.Name = "Form1" Then
'        If Form.EXIT_TRUE = True Then
'            If Err.Number = 0 Then EXIT_TRUE = True
'        End If
'    End If
'Next Form
'On Error GoTo 0
'If EXIT_TRUE = True Then Cancel = False
'
'For Each Form In Forms
'    Unload Form
'    Set Form = Nothing
'Next Form
'
'Unload Form1
'
''---------------------------------------------------------
''On Error Resume Next
''ERROR BECAUSE -- NOT CONTROLED HERE IF DRIVE DON'T EXIST
''---------------------------------------------------------
'
''UNLOAD_FORM_SAFE.Timer1.Enabled = True
'
''Cancel = True
End Sub

Sub SETUP_SOUND_FILE(VARSOUND)
' VARSOUND -- MAKE SOUND WHEN FIND %

' --------------------------------------------------------------------
' TIDY UP HERE WORK
' SETUP_SOUND_FILE(VARSOUND)
' AND AT ROUTINE --
' Private Sub MNU_SOUND_OPTION_Click(Index As Integer)
' --------------------------------------------------------------------
' Tue 02-Jun-2020 21:24:05
' Wed 03-Jun-2020 01:01:16 -- 3 HOUR 37 MINUTE
' --------------------------------------------------------------------

vPathSOUND2 = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\"
DATA_INFO_SOUND_SELECTOR = vPathSOUND2 + "#0__DATA_INFO_SOUND_FILE_SELECTOR.TXT"

Do
    If ScanPath.ListView1.ListItems.Count > 0 Then
        Sleep 1
    End If
    If EXIT_TRUE = True Then Exit Sub
Loop Until ScanPath.ListView1.ListItems.Count = 0

' Call SETUP_SOUND_FILE("NOTSOUND")
' VARSOUND = "NOTSOUND"
Dim XCOUNT1, XCOUNT2, XCHECKED

For Each Control In MNU_SOUND_OPTION
    If Control.Caption <> "MNU_SOUND_OPTION" Then
        XCOUNT1 = XCOUNT1 + 1
    End If
Next

'-----------------------------------
'ONLY USE IF WANT EXTRA DATE INFO ADDED
'-----------------------------------
ScanPath.chkCopyMemory.Value = vbUnchecked
' ---------------------------------------------------------
' \Sound_Wav's--2\ AND \Sound_Wav's\ REQUIRE SWAP OR RENAME
' ---------------------------------------------------------
ScanPath.TxtPath = App.Path + "\Sound_Wav's--2\" + GetComputerName
ScanPath.cboMask = "*.WAV"
ScanPath.chkSubFolders = vbUnchecked

ARCHIVE_Path_Of_Sound_File = Path_Of_Sound_File
Path_Of_Sound_File = ""

ScanPath.ListView1.ListItems.Clear
'WE WANT THEM SORTED ALPHABETICALY
Call ScanPath.cmdScan_Click
'LAST BREAK POINT WHILE WORK WAS SET HERE

ReDim Path_Of_Sound(ScanPath.ListView1.ListItems.Count)
' MSGBOX GONE OVER LIMIT
MAX_ITEM = 6 ' LEAVE ONE FOR SOUND 2
XCOUNT1 = ScanPath.ListView1.ListItems.Count
For WE = 1 To ScanPath.ListView1.ListItems.Count
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    If WE > MAX_ITEM Then MsgBox "MAXIMUM OF" + Str(MAX_ITEM) + " OF \SOUND_WAV'S--1\ IN MENU - DELETE TO DO": Exit For
    If WE > MAX_ITEM Then
        ' MsgBox "MAXIMUM OF 2 OF \SOUND_WAV'S--2\ IN MENU - DELETE TO DO"
        Exit For
    End If
    Path_Of_Sound(WE) = A1 + B1
    MNU_SOUND_OPTION(WE).Caption = "SOUND OPTION - 1 -" + Str(WE) + " OF" + Str(XCOUNT1) + " - \" + B1
    MNU_SOUND_OPTION(WE).Visible = True
    MNU_SOUND_OPTION(WE).Enabled = True
Next

ScanPath.TxtPath = App.Path + "\Sound_Wav's\"
ScanPath.cboMask = "*.WAV"
ScanPath.chkSubFolders = vbUnchecked
ScanPath.ListView1.ListItems.Clear
Call ScanPath.cmdScan_Click
XCOUNT2 = XCOUNT1
XCOUNT1 = ScanPath.ListView1.ListItems.Count
XCOUNT3 = XCOUNT1
MAX_ITEM = 12
If XCOUNT3 > MAX_ITEM Then XCOUNT3 = MAX_ITEM
For WE = 1 To ScanPath.ListView1.ListItems.Count
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    If WE > MAX_ITEM Then
        ' MsgBox "MAXIMUM OF 2 OF \SOUND_WAV'S--2\ IN MENU - DELETE TO DO"
        Exit For
    End If
    On Error Resume Next
    ' SORT TO DO MAX ABLE GET FILL THE MENU OPTION
    Err.Clear
    MNU_SOUND_OPTION(WE + XCOUNT2).Caption = "TEST"
    If Err.Number > 0 Then
        Exit For
    End If
    ' SORT TO DO MAX ABLE GET FILL THE MENU OPTION
    On Error GoTo 0
    
    MNU_SOUND_OPTION(WE + XCOUNT2).Caption = "SOUND OPTION - 2 -" + Str(WE) + " OF" + Str(XCOUNT3) + " - \" + B1
    MNU_SOUND_OPTION(WE + XCOUNT2).Visible = True
    MNU_SOUND_OPTION(WE + XCOUNT2).Enabled = True
Next


ScanPath.ListView1.ListItems.Clear
ScanPath.Hide

'GOT TO RUN TWICE - THINK
'SEARCHES FOR THE TEXT CAPTIONS BEEN PUT IN
'--------
'DONT RUN AGAIN WHEN PROGRAM IS GOING SETTINGS ONLY SAVED ON EXIT
'AND YES DO RUN TWICE ON FIRST LOAD

' -----------------------------------------------
' -----------------------------------------------
If LATCH_RUN_ONCE = False Then Call zzLoad_Checks
' -----------------------------------------------
' -----------------------------------------------

' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' FILENAME SAVE SUPPOSE TO BE DO BY SAVE CHECKER
' OPEN FILE WAY INSTEAD
If RUN_ONCE_AT_START = False Then
    RUN_ONCE_AT_START = True
    ' READ FILE --------------------------------------------------------
    ' HERE BE BETER MOVE TO FIRST PATH OF SOUND OR DATA FOLDER
    ' "\SOUND_WAV'S"
    ' vPathSOUND2 = App.Path + "\Sound_Wav's\" ' + GetComputerName + "\"
    If Dir(DATA_INFO_SOUND_SELECTOR) <> "" Then
        On Error Resume Next
        FR1 = FreeFile
        Open DATA_INFO_SOUND_SELECTOR For Input As #FR1
            Line Input #FR1, MMCONTROL1__FILENAME
            Line Input #FR1, MMCONTROL2__FILENAME
        Close #FR1
        On Error GoTo 0
    End If
    If MMCONTROL1__FILENAME <> "" Then
        If Dir(MMCONTROL1__FILENAME) = "" Then MMCONTROL1__FILENAME = ""
    End If
    If MMCONTROL2__FILENAME <> "" Then
        If Dir(MMCONTROL2__FILENAME) = "" Then MMCONTROL2__FILENAME = ""
    End If
    ' DO ---- MMCONTROL1__FILENAME
    If MMCONTROL1__FILENAME <> "" Then
    If Dir(MMCONTROL1__FILENAME) <> "" Then
    IMM = MMCONTROL1__FILENAME
    IMM = Mid(IMM, InStrRev(IMM, "\"))
    If IMM <> "" Then
        For Each Control In MNU_SOUND_OPTION
            If InStr(Control.Caption, IMM) > 0 Then
                ' Path_Of_Sound_File = App.Path + "\Sound_Wav's" + Mid(XCHECKED, 5)
                Path_Of_Sound_File = MMCONTROL1__FILENAME
                Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
                Control.Checked = True
                Control.Enabled = True
                MNU_SOUND_WAV_1_CHECKER_INDEX = Control.Caption
                Exit For
            End If
        Next
    End If
    End If
    End If
    
    ' DO ---- MMCONTROL2__FILENAME
    ' -----------------------------------------------
    ' HERE NOT DO ANY YET
    ' AS ABOVE ONE FOR
    ' MMCONTROL1__FILENAME
    ' BUT HERE NOT POPULATE AS ONLY ONE AT THE MOMENT
    ' RUN ROUTINE NOT PICK UP ANY -- DEMO
    ' POISTION CODE ABLE MOVE LOWER
    ' -----------------------------------------------
    ' DO ---- MMCONTROL2__FILENAME
    If MMCONTROL2__FILENAME <> "" Then
    If Dir(MMCONTROL2__FILENAME) <> "" Then
        IMM = MMCONTROL2__FILENAME
        IMM = Mid(IMM, InStrRev(IMM, "\"))
        If MMCONTROL2__FILENAME <> "" Then
            For Each Control In MNU_SOUND_OPTION
                If InStr(Control.Caption, IMM) > 0 Then
                    ' Path_Of_Sound_File = App.Path + "\Sound_Wav's" + Mid(XCHECKED, 5)
                    vFileSOUND2 = MMCONTROL2__FILENAME
                    Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
                    Control.Checked = True
                    Control.Enabled = True
                    MNU_SOUND_WAV_2_CHECKER_INDEX = Control.Caption
                    Exit For
                End If
            Next
        End If
    End If
    End If
End If

' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' IF MORE THAN 1 SELECTOR THEN UNSELECT ALL <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' FIND THE CHECKER ONE SELECTOR   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' AND THEN -- UNSELECT ALL -- RESTORE BACK
' IF NONE SELECT PICK FIRST
' -----------------------------------------------------------------------
' IF MORE THAN 1 SELECTOR THEN UNSELECT ALL <<<< SOUND OPTION - 1
COUNT_CONTROL_CHECKED = 0
For Each Control In MNU_SOUND_OPTION
    If InStr(Control.Caption, "SOUND OPTION - 1") > 0 Then
    If Control.Checked = True Then
    COUNT_CONTROL_CHECKED = COUNT_CONTROL_CHECKED + 1
    End If
    End If
Next
If COUNT_CONTROL_CHECKED > 1 Then
For Each Control In MNU_SOUND_OPTION
    If InStr(Control.Caption, "SOUND OPTION - 1") > 0 Then
        Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
        Control.Checked = False
    End If
Next
End If
' TRY HERE ROUTINE SNIPPET AGAIN
' DO ---- MMCONTROL1__FILENAME
IMM = MMCONTROL1__FILENAME
If IMM <> "" Then
    IMM = Mid(IMM, InStrRev(IMM, "\"))
End If
If MMCONTROL1__FILENAME <> "" Then
If Dir(MMCONTROL1__FILENAME) <> "" Then
If IMM <> "" Then
    For Each Control In MNU_SOUND_OPTION
        If InStr(Control.Caption, IMM) > 0 Then
            ' Path_Of_Sound_File = App.Path + "\Sound_Wav's" + Mid(XCHECKED, 5)
            Path_Of_Sound_File = MMCONTROL1__FILENAME
            Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
            Control.Checked = True
            Control.Enabled = True
            MNU_SOUND_WAV_1_CHECKER_INDEX = Control.Caption
            Exit For
        End If
    Next
End If
End If
End If
' -----------------------------------------------------------------------
' IF MORE THAN 1 SELECTOR THEN UNSELECT ALL <<<< SOUND OPTION - 2
COUNT_CONTROL_CHECKED = 0
For Each Control In MNU_SOUND_OPTION
    If InStr(Control.Caption, "SOUND OPTION - 2") > 0 Then
    If Control.Checked = True Then
    COUNT_CONTROL_CHECKED = COUNT_CONTROL_CHECKED + 1
    End If
    End If
Next
If COUNT_CONTROL_CHECKED > 1 Then
For Each Control In MNU_SOUND_OPTION
    If InStr(Control.Caption, "SOUND OPTION - 2") > 0 Then
        Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
        Control.Checked = False
    End If
Next
End If
' TRY HERE ROUTINE SNIPPET AGAIN
' DO ---- MMCONTROL2__FILENAME
IMM = MMCONTROL2__FILENAME
If InStrRev(IMM, "\") > 0 Then
    IMM = Mid(IMM, InStrRev(IMM, "\"))
End If
If MMCONTROL2__FILENAME <> "" Then
    If Dir(MMCONTROL2__FILENAME) <> "" Then
        For Each Control In MNU_SOUND_OPTION
            If InStr(Control.Caption, IMM) > 0 Then
                ' Path_Of_Sound_File = App.Path + "\Sound_Wav's" + Mid(XCHECKED, 5)
                vFileSOUND2 = MMCONTROL2__FILENAME
                Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
                Control.Checked = True
                Control.Enabled = True
                MNU_SOUND_WAV_2_CHECKER_INDEX = Control.Caption
                Exit For
            End If
        Next
    End If
End If









' -----------------------------------------------------------------------
' IF FILE NOT EXIST SELECT 1ST <<<< SOUND OPTION - 1
If MMCONTROL1__FILENAME <> "" Then
If Dir(MMCONTROL1__FILENAME) = "" Then
    COUNT_CONTROL_CHECKED = 0
    For Each Control In MNU_SOUND_OPTION
        If InStr(Control.Caption, "SOUND OPTION - 1") > 0 Then
        If Control.Checked = True Then
            COUNT_CONTROL_CHECKED = COUNT_CONTROL_CHECKED + 1
        End If
        End If
    Next
    If COUNT_CONTROL_CHECKED > 1 Then
    For Each Control In MNU_SOUND_OPTION
        If InStr(Control.Caption, "SOUND OPTION - 1") > 0 Then
            Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
            Control.Checked = False
        End If
    Next
    End If
    For Each Control In MNU_SOUND_OPTION
        If InStr(Control.Caption, "SOUND OPTION - 1") > 0 Then
            Path_Of_Sound_File = MMCONTROL1__FILENAME
            Path_Of_Sound_File = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\" + Mid(Control.Caption, InStrRev(Control.Caption, "\") + 1)
            MMCONTROL1__FILENAME = Path_Of_Sound_File
            Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
            Control.Checked = True
            Control.Enabled = True
            MNU_SOUND_WAV_1_CHECKER_INDEX = Control.Caption
            Exit For
        End If
    Next
End If
End If
' -----------------------------------------------------------------------
' IF FILE NOT EXIST SELECT 1ST <<<< SOUND OPTION - 2
If MMCONTROL2__FILENAME <> "" Then
If Dir(MMCONTROL2__FILENAME) = "" Then
    COUNT_CONTROL_CHECKED = 0
    For Each Control In MNU_SOUND_OPTION
        If InStr(Control.Caption, "SOUND OPTION - 2") > 0 Then
        If Control.Checked = True Then
            COUNT_CONTROL_CHECKED = COUNT_CONTROL_CHECKED + 1
        End If
        End If
    Next
    If COUNT_CONTROL_CHECKED > 1 Then
    For Each Control In MNU_SOUND_OPTION
        If InStr(Control.Caption, "SOUND OPTION - 2") > 0 Then
            Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
            Control.Checked = False
        End If
    Next
    End If
    For Each Control In MNU_SOUND_OPTION
        If InStr(Control.Caption, "SOUND OPTION - 2") > 0 Then
            vFileSOUND2 = App.Path + "\Sound_Wav's\" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
            MMCONTROL2__FILENAME = vFileSOUND2
            Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
            Control.Checked = True
            Control.Enabled = True
            MNU_SOUND_WAV_2_CHECKER_INDEX = Control.Caption
            Exit For
        End If
    Next
End If
End If







' NOT SURE IF REQUIRE HERE -- CODE LEFT FROM BEFORE AS WORK
' IF NONE SELECT FOR ---- MNU_SOUND_WAV_1_CHECKER_INDEX ---- GET 1ST
If MNU_SOUND_WAV_1_CHECKER_INDEX = "" Then
For Each Control In MNU_SOUND_OPTION
    If InStr(Control.Caption, "SOUND OPTION - 1") > 0 Then
    If Control.Checked = True Then
        MNU_SOUND_WAV_1_CHECKER_INDEX = Control.Caption
        Exit For
    End If
    End If
Next
End If
' IF NONE SELECT FOR ---- MNU_SOUND_WAV_2_CHECKER_INDEX ---- GET 1ST
If MNU_SOUND_WAV_2_CHECKER_INDEX = "" Then
For Each Control In MNU_SOUND_OPTION
    If InStr(Control.Caption, "SOUND OPTION - 2") > 0 Then
    If Control.Checked = True Then
        MNU_SOUND_WAV_2_CHECKER_INDEX = Control.Caption
        Exit For
    End If
    End If
Next
End If
' -----------------------------------------------------------------------
' FIND THE CHECKER ONE SELECTOR
' IF NONE -- VARIABLE STORE FIRST     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' AND THEN -- UNSELECT ALL -- RESTORE ONE
' -----------------------------------------------------------------------
If MNU_SOUND_WAV_1_CHECKER_INDEX = "" Then
    For Each Control In MNU_SOUND_OPTION
        If InStr(Control.Caption, "SOUND OPTION - 1") > 0 Then
            ' Path_Of_Sound_File = App.Path + "\Sound_Wav's" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
            Path_Of_Sound_File = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\" + Mid(Control.Caption, InStrRev(Control.Caption, "\") + 1)
'            Path_Of_Sound_File = App.Path + "\Sound_Wav's\" + GetComputerName + "\" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))

            MNU_SOUND_WAV_1_CHECKER_INDEX = Control.Caption
            Exit For
        End If
    Next
End If
If MNU_SOUND_WAV_2_CHECKER_INDEX = "" Then
    For Each Control In MNU_SOUND_OPTION
        If InStr(Control.Caption, "SOUND OPTION - 2") > 0 Then
            ' vFileSOUND2 = App.Path + "\Sound_Wav's" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
            ' vPathSOUND2 = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
            vFileSOUND2 = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
            vFileSOUND2 = App.Path + "\Sound_Wav's\" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
            MNU_SOUND_WAV_2_CHECKER_INDEX = Control.Caption
            Exit For
        End If
    Next
End If

' -----------------------------------------------------------------------
' FIND THE CHECKER ONE SELECTOR
' IF NONE -- VARIABLE STORE FIRST
' TESTER FOR CHANGE -- IF DO -- UNSELECT ALL -- RESTORE ONE <<<<<<<<<<<<<
' -----------------------------------------------------------------------
If OLD_MNU_SOUND_WAV_1_CHECKER_INDEX <> MNU_SOUND_WAV_1_CHECKER_INDEX Then
' CLEAR ALL THE CONTROL.CHECKER -- FOR SOUND OPTION - 1
For Each Control In MNU_SOUND_OPTION
    If InStr(Control.Caption, "SOUND OPTION - 1") > 0 Then
        Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
        Control.Checked = False
        Control.Enabled = True
        Control.Visible = True 'AND SAFE MESSURE
    End If
Next
' SET THE CONTROL.CHECKER -- FOR SOUND OPTION - 1
For Each Control In MNU_SOUND_OPTION
    If InStr(Control.Caption, "SOUND OPTION - 1") > 0 Then
    If InStr(Control.Caption, MNU_SOUND_WAV_1_CHECKER_INDEX) > 0 Then
        MNU_SOUND_WAV_1_CHECKER_INDEX = Control.Caption
        Path_Of_Sound_File = App.Path + "\Sound_Wav's" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
        Path_Of_Sound_File = App.Path + "\Sound_Wav's\" + GetComputerName + "\" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
        Path_Of_Sound_File = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\" + Mid(Control.Caption, InStrRev(Control.Caption, "\") + 1)

        MMCONTROL1__FILENAME = App.Path + "\Sound_Wav's" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
        Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
        Control.Checked = True
        Control.Enabled = True
        Exit For
    End If
    End If
Next
End If
OLD_MNU_SOUND_WAV_1_CHECKER_INDEX = MNU_SOUND_WAV_1_CHECKER_INDEX
' ------------------------------------------------------------------------
' ---- NUMBER -- SOUND_WAV'S--2 ------------------------------------------
' ------------------------------------------------------------------------
If OLD_MNU_SOUND_WAV_2_CHECKER_INDEX <> MNU_SOUND_WAV_2_CHECKER_INDEX Then
' CLEAR ALL THE CONTROL.CHECKER -- FOR SOUND OPTION - 2
For Each Control In MNU_SOUND_OPTION
    If InStr(Control.Caption, "SOUND OPTION - 2") > 0 Then
        Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
        Control.Checked = False
        Control.Enabled = True
        Control.Visible = True 'AND SAFE MESSURE
    End If
Next
' SET THE CONTROL.CHECKER -- FOR SOUND OPTION - 2
For Each Control In MNU_SOUND_OPTION
    If InStr(Control.Caption, "SOUND OPTION - 2") > 0 Then
    If InStr(Control.Caption, MNU_SOUND_WAV_2_CHECKER_INDEX) > 0 Then
        MNU_SOUND_WAV_2_CHECKER_INDEX = Control.Caption
        ' vFileSOUND2 = App.Path + "\Sound_Wav's" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
        vFileSOUND2 = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
        vFileSOUND2 = App.Path + "\Sound_Wav's\" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))

        MMCONTROL2__FILENAME = App.Path + "\Sound_Wav's" + Mid(Control.Caption, InStrRev(Control.Caption, "\"))
        Control.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
        Control.Checked = True
        Control.Enabled = True
        Exit For
    End If
    End If
Next
End If
OLD_MNU_SOUND_WAV_2_CHECKER_INDEX = MNU_SOUND_WAV_2_CHECKER_INDEX

'PICK THE FIRST ONE - PICK ME UP - GIVE US A LIFT - DON'T MIND IF I DO - HELP MYSELF - INDULGE MYSELF
'TAKE THE FIRST ONE

'MNU_SOUND_OPTION(1).Enabled = False ' DON'T DO EXTRA CLICK'S TO THE ROUTINE
'MNU_SOUND_OPTION(1).Checked = True
'MNU_SOUND_OPTION(1).Visible = True
'MNU_SOUND_OPTION(1).Enabled = True ' SAFE

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

If Path_Of_Sound_File <> "" Then
    MMControl1.Command = "Stop"
    MMControl1.Command = "Close"
    MMControl1.Notify = True
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.Filename = Path_Of_Sound_File
    'Debug.Print Path_Of_Sound_File
    'Path_Of_Sound(1) = App.Path + "\Camera1a_2_8kHz.wav"
    'MMControl1.fileName = App.Path + "\Camera1a_2_8kHz.wav"
    'MMControl1.fileName = App.Path + "\Shot-12 Mix to Shot-18.wav"
    'MMControl1.fileName = App.Path + "\01 Woarble Tone 5 Mins.wav"
    MMControl1.Command = "Open"
End If

If vFileSOUND2 <> "" Then
    IMM = vFileSOUND2
    IMM = Mid(IMM, InStrRev(IMM, "\"))
    ' MNU_SOUND_2.Caption = "SOUND OPTION - 2 ------ \" + IMM
    MMControl2.Command = "Stop"
    MMControl2.Command = "Close"
    '---- Set properties needed by MCI to open.
    MMControl2.Notify = True
    MMControl2.Wait = True
    MMControl2.Shareable = False
    MMControl2.DeviceType = "WaveAudio"
    MMControl2.Filename = vFileSOUND2
    ' Open the MCI WaveAudio device.
    MMControl2.Command = "Open"
    ' MMControl2.Command = "prev" :  MMControl2.Command = "Play"
End If

If Mnu_SoundOn.Checked = True And VARSOUND <> "NOTSOUND" Then
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
End If

' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' -----------------------------------------------------------------------
' FILENAME SAVE SUPPOSE TO BE DO BY SAVE CHECKER
' OPEN FILE WAY INSTEAD

On Error Resume Next
FR1 = FreeFile
Open DATA_INFO_SOUND_SELECTOR For Output As #FR1
    Print #FR1, MMControl1.Filename
    Print #FR1, MMControl2.Filename
Close #FR1
On Error GoTo 0

ScanPath.ListView1.ListItems.Clear

'WORK TO DO HERE
LATCH_RUN_ONCE = True   ' Call zzLoad_Checks
MODIFY_SOUND_SELECTION = False
DUPE_CLIPPER_AT_LOAD_FORM = False

End Sub




Sub RESET_SETUP_SOUND_FILE(VARSOUND)
'MMControl2.fileName = vFileSOUND2
'Debug.Print Path_Of_Sound_File

'If VARSOUND <> "" Then Stop

On Error GoTo EXIT_END


'compare sort
'If Dir(Path_Of_Sound_File) = "" Then
'    Me.WindowState = vbNormal
'    MsgBox "SOUND FILE 01# IS MISSING", vbMsgBoxSetForeground
'End If

'If Dir(vFileSOUND2) = "" Then
'    Me.WindowState = vbNormal
'    MsgBox "SOUND FILE 02# IS MISSING", vbMsgBoxSetForeground
'End If









'
''Path_Of_Sound_File = App.Path + "\Sound_Wav's\"
'vPathSOUND1 = App.Path + "\Sound_Wav's\"
'vFileSOUND1 = Dir(vPathSOUND1 + "*.WAV")
'Path_Of_Sound_File = vPathSOUND1 + vFileSOUND1
'
'If Path_Of_Sound_File <> "" And Dir(Path_Of_Sound_File) <> "" Then
'
'    MMControl1.Command = "Stop"
'    MMControl1.Command = "Close"
'    MMControl1.Notify = True
'    MMControl1.Wait = True
'    MMControl1.Shareable = False
'    MMControl1.DeviceType = "WaveAudio"
'    MMControl1.fileName = Path_Of_Sound_File
''    Debug.Print Path_Of_Sound_File
'
'    'Path_Of_Sound(1) = App.Path + "\Camera1a_2_8kHz.wav"
'    'MMControl1.fileName = App.Path + "\Camera1a_2_8kHz.wav"
'    'MMControl1.fileName = App.Path + "\Shot-12 Mix to Shot-18.wav"
'    'MMControl1.fileName = App.Path + "\01 Woarble Tone 5 Mins.wav"
'    MMControl1.Command = "Open"
'    MNU_AUDIO_01_MISSING.Visible = False
'
'End If
'
'
'vPathSOUND2 = App.Path + "\Sound_Wav's--2\" + GetComputerName + "\"
'vFileSOUND2 = vPathSOUND2 + Dir(vPathSOUND2 + vPathSOUND2 + "*.WAV")
'
'If Dir(vFileSOUND2) <> "" Then
'    IMM = vFileSOUND2
'    IMM = Mid(IMM, InStrRev(IMM, "\"))
'    MNU_SOUND_2.Caption = "SOUND OPTION - 2 ------ \" + IMM
'
'    MMControl2.Command = "Stop"
'    MMControl2.Command = "Close"
'    '---- Set properties needed by MCI to open.
'    MMControl2.Notify = True
'    MMControl2.Wait = True
'    MMControl2.Shareable = False
'    MMControl2.DeviceType = "WaveAudio"
'    MMControl2.fileName = vFileSOUND2
'    ' Open the MCI WaveAudio device.
'    MMControl2.Command = "Open"
'
'    'MMControl2.Command = "prev"
'    'MMControl2.Command = "Play"
'
'    MNU_AUDIO_01_MISSING.Visible = False
'End If
'
'If DUPE_CLIPPER_AT_LOAD_FORM = False Then

'If Mnu_SoundOn.Checked = True And VARSOUND <> "NOTSOUND" Then

If STARTUP_RUN = False Then
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
End If

If Dir(Path_Of_Sound_File) = "" Then
    Me.WindowState = vbNormal
    'DEBUG_HERE
    'MsgBox "SOUND FILE 01# IS MISSING", vbMsgBoxSetForeground
    MNU_AUDIO_01_MISSING.Visible = True
Else
    MNU_AUDIO_01_MISSING.Visible = False
End If

If Dir(vFileSOUND2) = "" Then
    Me.WindowState = vbNormal
    Stop
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
    
    Dim ITXT As String
    
    If MNU_LAST_GRAB_ALL_CAPS.Checked = False Then
        MNU_LAST_GRAB_ALL_CAPS.Checked = True
        ' ---------------------------------------
        ' CAN'T CHECK HERE AS MENU NOT SUB ITEM
        ' ---------------------------------------
        ' MNU_TEXT_CAPITAL_MODE_ON.Checked = True
        ' ---------------------------------------
        ITXT = MNU_TEXT_CAPITAL_MODE_ON.Caption
        ITXT = Replace(ITXT, "_OFF", "_ON")
        MNU_TEXT_CAPITAL_MODE_ON.Caption = ITXT
        GoTo EXIT1
    End If

    If MNU_LAST_GRAB_ALL_CAPS.Checked = True Then
        MNU_LAST_GRAB_ALL_CAPS.Checked = False
        ' ----------------------------------------
        ' CAN'T CHECK HERE AS MENU NOT SUB ITEM
        ' ----------------------------------------
        ' MNU_TEXT_CAPITAL_MODE_ON.Checked = False
        ' ----------------------------------------
        MNU_TEXT_CAPITAL_MODE_ON.Checked = False
        ITXT = MNU_TEXT_CAPITAL_MODE_ON.Caption
        ITXT = Replace(I, "_ON", "_OFF")
        MNU_TEXT_CAPITAL_MODE_ON.Caption = ITXT
    End If

    If MNU_LAST_GRAB_ALL_CAPS.Checked = False Then Exit Sub

EXIT1:
    
    AD$ = UCase(AD$)
    EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        Clipboard.SetText AD$
    EXECUTE_TIMER_ENABLED = True
    AD_DATE = 0
    
    Me.WindowState = 1
    
    ' Solvability
              
End Sub

Private Sub LAST_GRAB_PRO_CAPS_002_Click()
    
    If MNU_LAST_GRAB_PRO_CAPS.Checked = False Then Exit Sub
    
    If Clipboard.GetFormat(vbCFText) = False Then Exit Sub
    
    Call Mnu_Fix1st_Letters_Click

End Sub




Private Sub LAST_GRAB_ALL_CAPS_002_Click()
    
    If MNU_LAST_GRAB_ALL_CAPS.Checked = False Then Exit Sub
    
    CHANGE_AD = AD$
    AD$ = UCase(AD$)
    If AD$ = CHANGE_AD Then Exit Sub
    
    RT = Now + TimeSerial(0, 0, 20)
    On Error Resume Next
    ERROR_HITTER = False
    For R2 = 1 To 10000
        Err.Clear
        EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        If Err.Number = 0 Then
            Clipboard.SetText AD$
        End If
        If Err.Number > 0 Then
            Debug.Print Err.Description + "_TRY _ " + Str(R2)
            ERROR_HITTER = True
            Sleep 100
            DoEvents
        End If
        If RT < Now Or Err.Number = 0 Then
            Exit For
        End If
        EXECUTE_TIMER_ENABLED = True
    Next
    If IsIDE = True Then
        If ERROR_HITTER = True And R2 > 40 Then Stop
    End If
    
    O_2_AD = AD_CAP
    AD_DATE = 0
    
    Me.WindowState = 1
    
    ' SOLVABILITY

End Sub




Private Sub LAST_GRAB_ALL_CAPS_002_Click_OLD()
    
    If MNU_LAST_GRAB_ALL_CAPS.Checked = False Then Exit Sub
    
    If Clipboard.GetFormat(vbCFText) = False Then Exit Sub
    
    ' --------------------------------------------
    ' MAYBE A FAULT BUT ONLY WANT TEXT THAT CHANGE
    ' OR USE THE API
    ' MORE CODE TO DO MAKE EFFECIENT -- Sat 14-Sep-2019 20:42:40
    ' --------------------------------------------
    
    Clipboard.GetText AD2
    
    AD_CAP = UCase(AD$)
    If AD_CAP = AD$ Then
        API_CLIPBOARD_TEXT_CHANGE_FOR_FORMAT_ALL_CAPITAL_MODE = False
        O_2_AD = AD$
        Exit Sub
    End If
    
    If API_CLIPBOARD_TEXT_CHANGE_FOR_FORMAT_ALL_CAPITAL_MODE = True Then
        API_CLIPBOARD_TEXT_CHANGE_FOR_FORMAT_ALL_CAPITAL_MODE = False
        O_2_AD = Str(Now) + Str(Timer)
    End If
    
    If O_2_AD = AD_CAP Then Exit Sub
    
    RT = Now + TimeSerial(0, 0, 20)
    On Error Resume Next
    ERROR_HITTER = False
    For R2 = 1 To 10000
        Err.Clear
        EXECUTE_TIMER_ENABLED = False
        Clipboard.Clear
        If Err.Number = 0 Then
            Clipboard.SetText AD_CAP
        End If
        If Err.Number > 0 Then
            Debug.Print Err.Description + "_TRY _ " + Str(R2)
            ERROR_HITTER = True
            Sleep 400
            DoEvents
        End If
        If RT < Now Or Err.Number = 0 Then
            Exit For
        End If
        EXECUTE_TIMER_ENABLED = True
    Next
    If IsIDE = True Then
        If ERROR_HITTER = True And R2 > 40 Then Stop
    End If
    O_2_AD = AD_CAP
    
    Me.WindowState = 1
    AD_DATE = 0
    
    ' SOLVABILITY

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
    EXPLORER_PATH = Path_Of_Sound_File
    Me.WindowState = vbMinimized
    If Dir(EXPLORER_PATH) <> "" Then
        Shell "Explorer.exe /select," + EXPLORER_PATH, vbNormalFocus
        Exit Sub
    End If
    EXPLORER_PATH = Mid(EXPLORER_PATH, 1, InStrRev(EXPLORER_PATH, "\"))
    Shell "Explorer.exe """ + EXPLORER_PATH + """", vbNormalFocus
End Sub

Private Sub Mnu_Explorer_Sound_2_Click()
    EXPLORER_PATH = vFileSOUND2
    Me.WindowState = vbMinimized
    If Dir(EXPLORER_PATH) <> "" Then
        Shell "Explorer.exe /select," + EXPLORER_PATH, vbNormalFocus
        Exit Sub
    End If
    EXPLORER_PATH = Mid(EXPLORER_PATH, 1, InStrRev(EXPLORER_PATH, "\"))
    Shell "Explorer.exe """ + EXPLORER_PATH + """", vbNormalFocus
End Sub


Private Sub Mnu_Find_New_Sounds_Click()
    
    'THIS GET RUN HERE FIND NEW SOUND
    'AND AT ONE CLICK
    'AND AT MENU RESET AUDIO
    'FIRST RUN FORM LOAD
    '------------------------------
    'RELOAD OR MODIFY THE SELECTION
    '------------------------------
    
    MODIFY_SOUND_SELECTION = True
    Call SETUP_SOUND_FILE("")

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

Mid$(EE, 1, 1) = UCase(Mid$(EE, 1, 1))

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

ii = GetShortName(PATH_TO_CLIPBOARD_APP_DATA_1 + "\")
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

'Set F = FSO.getfile((Filespec1))
'ADATE1 = F.datelastmodified

'ScanPath.lblCount7 = ""
'ScanPath.ListView1.ListItems.Clear

'ScanPath.TxtPath = "D:\0 00 Art Loggers - WEBCAM\"
'Call ScanPath.cmdScan_Click


'Filespec2 = ScanPath.lblCount7
'If Filespec2 <> "" Then
'    Set F = FSO.getfile((Filespec2))
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
    EE = Replace(EE, "(" + LCase(Chr(r)), "(" + UCase(Chr(r)))
    EE = Replace(EE, "." + LCase(Chr(r)), "." + UCase(Chr(r)))
    EE = Replace(EE, "\" + LCase(Chr(r)), "\" + UCase(Chr(r)))
    EE = Replace(EE, "_" + LCase(Chr(r)), "_" + UCase(Chr(r)))
    EE = Replace(EE, """" + LCase(Chr(r)), """" + UCase(Chr(r)))
Next

'HERE IS -- PROCAPS

Mid$(EE, 1, 1) = UCase(Mid$(EE, 1, 1))

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

EXIT_TRUE = True

Unload Form1

End Sub

Private Sub MNU_INFO_RAPID_Click()

'ii = " /nologo /DirD:\#*MY*DOCS\01*My*Documents\03*NotePad*Files\00*TOP\  /FileClued*Up.txt"


ii = " /nologo /Dir" + App.Path + "\#*DATA\" + GetComputerName + "\  /File00_ClipBoard_Total--TRIM.txt"
'ii = Replace(ii, "ClipBoard Logger", "Fast*Clipboard")
ii = Replace(ii, UCase("Clipboard Logger"), UCase("Clipboard*Logger"))
ii = Replace(ii, "\# Data", "\#*Data")

Shell "C:\Program Files\seRapid\seRapid.exe " + ii, vbNormalFocus

Me.WindowState = 1

End Sub

Private Sub MNU_INFO_RAPID_TOTAL_Click()

'ii = " /nologo /DirD:\#*MY*DOCS\01*My*Documents\03*NotePad*Files\00*TOP\  /FileClued*Up.txt"

' "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\4-ASUS-GL522VW_MATT 01\00_ClipBoard_Total.txt"
PATH_STRING_TO_WILD_CARD = "/Dir" + App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\  /File00_ClipBoard_Total.txt"
PATH_STRING_TO_WILD_CARD = Replace(PATH_STRING_TO_WILD_CARD, " ", "*")
PATH_STRING_TO_WILD_CARD = Replace(PATH_STRING_TO_WILD_CARD, "\**/", "\  /")
ii = " /nologo " + PATH_STRING_TO_WILD_CARD

Shell "C:\Program Files\seRapid\seRapid.exe " + ii, vbNormalFocus

Me.WindowState = 1

End Sub
Private Sub MNU_INFO_RAPID_ALL_Click()

'ii = " /nologo /WithSubdirectoriesYES /Dir" + App.Path + "\#*DATA\" + GetComputerName + "\  /File/*.TXT"
ii = " /nologo /WithSubdirectoriesYES /Dir" + App.Path + "\#*DATA\ /File*.TXT"
ii = Replace(ii, UCase("Clipboard Logger"), UCase("Clipboard*Logger"))
ii = Replace(ii, "\# Data", "\#*Data")

Shell "C:\Program Files\seRapid\seRapid.exe " + ii, vbNormalFocus
Me.WindowState = 1


'ClipBoard-

End Sub

Private Sub MNU_INFO_RAPID_ALL_SMALL_FILES_Click()


Timer_INFORAPID_MSGBOX.Enabled = True

ii = " /nologo /WithSubdirectoriesYES /Dir" + App.Path + "\#*DATA\ /FileClipBoard-*.TXT"
ii = Replace(ii, UCase("Clipboard Logger"), UCase("Clipboard*Logger"))
ii = Replace(ii, "\# Data", "\#*Data")

xd_1 = "C:\Program Files\seRapid\seRapid.exe"
xd_2 = "C:\Program Files (x86)\seRapid\seRapid.exe"
xd = xd_1
If Dir(xd) = "" Then
    xd = xd_2
End If
If Dir(xd) = "" Then
    MsgBox "Not Find " + vbCrLf + xd_1 + vbCrLf + "or" + vbCrLf + xd_2
    Exit Sub
End If


Shell xd + " " ' + ii, vbNormalFocus
Me.WindowState = 1


'ClipBoard-


End Sub

Private Sub MNU_LAST_ART_PIC_Click()

    If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
        Call RECURSIVE_SEARCH_LAST_CLIPBOARD_ART_FOLDER
    End If
    
    If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
        MsgBox "File None Found Last Art Clipboard Search Var _ LAST_CLIPBOARD_FILE"
        Exit Sub
    End If
    
    If FSO.FileExists(LAST_CLIPBOARD_FILE) = False Then
        MsgBox "File Found Not Valid Last Art Clipboard Search Var _ LAST_CLIPBOARD_FILE"
        Exit Sub
    End If
    
    
    Me.WindowState = vbMinimized
    Shell "Explorer.exe /SELECT," + LAST_CLIPBOARD_FILE, vbMaximizedFocus
    Beep


'Call LAST_IMAGE("EXPLORER", "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\Hot-Key-App-Shots\") '=EXPLORER


End Sub

Private Sub MNU_LOGGEXPLORER_Click(Index As Integer)

Dim M()
ReDim M(MNU_LOGGEXPLORER.Count)
I = 0
I = I + 1: M(I) = "OPEN LOGG EXPLORER"
I = I + 1: M(I) = "----"
I = I + 1: M(I) = "INFO RAPID ALL FOLDERS"
I = I + 1: M(I) = "INFO RAPID ALL FOLDERS AND SMALL FILES BEGIN CLIPBOARD-*.TXT"
I = I + 1: M(I) = "INFO RAPID MY USER"
I = I + 1: M(I) = "INFO RAPID TRIM LOGG"
I = I + 1: M(I) = "INFO RAPID TOTAL LOGG" ' __ BETTER DATE ORDER SORT __ MIGHT SOME MISS AH DUE TO OLD DAY AND CRASH APP"
I = I + 1: M(I) = "----"
I = I + 1: M(I) = "TEXTCRAWLER ALL FOLDERS"
I = I + 1: M(I) = "TEXTCRAWLER TRIM LOGG"
I = I + 1: M(I) = "TEXTCRAWLER STRIP LOGG"
I = I + 1: M(I) = "----"
I = I + 1: M(I) = "FILELOCATOR ALL FOLDERS"
I = I + 1: M(I) = "FILELOCATOR IMAGE SET"
I = I + 1: M(I) = "FILELOCATOR IMAGE SET AUTO"
I = I + 1: M(I) = "FILELOCATOR IMAGE SET AUTO ALL COMPUTER"
I = I + 1: M(I) = "----"
I = I + 1: M(I) = "HEX OPEN RECENT TRIM LOGG"
I = I + 1: M(I) = "----"
I = I + 1: M(I) = "EDIT RECENT TRIM LOGG"
I = I + 1: M(I) = "EDIT THIS LOGG"
I = I + 1: M(I) = "EDIT TOTAL LOGG"
I = I + 1: M(I) = "EDIT STRIP LOGG"
I = I + 1: M(I) = "----"
I = I + 1: M(I) = "TEXT VIEW RECENT TRIM LOGG"
I = I + 1: M(I) = "TEXT VIEW THIS LOGG"
I = I + 1: M(I) = "TEXT VIEW TOTAL LOGG"
I = I + 1: M(I) = "TEXT VIEW STRIP LOGG"
I = I + 1: M(I) = "----"
I = I + 1: M(I) = "SHELL T -- THIS LOGG"
I = I + 1: M(I) = "SHELL T -- TOTAL LOGG"

For r = 1 To MNU_LOGGEXPLORER.Count
    If MNU_LOGGEXPLORER(r).Caption <> M(r) Then
        MNU_LOGGEXPLORER(r).Caption = M(r)
    End If
    If MNU_LOGGEXPLORER(r).Caption = "" Then
        MNU_LOGGEXPLORER(r).Visible = False
    End If

Next

Select Case M(Index)

Case "OPEN LOGG EXPLORER"
    ' Call MNU_LOGGEXPLORER_Click
    Me.WindowState = vbMinimized
    Shell "Explorer.exe /e," + PATH_TO_CLIPBOARD_APP_DATA_1 + """", vbNormalFocus

Case "INFO RAPID ALL FOLDERS"
Call MNU_INFO_RAPID_ALL_Click

Case "INFO RAPID ALL FOLDERS AND SMALL FILES BEGIN CLIPBOARD-*.TXT"
Call MNU_INFO_RAPID_ALL_SMALL_FILES_Click

Case "INFO RAPID MY USER"
Call MNU_INFO_RAPID_MYUSER_Click
    
Case "INFO RAPID TRIM LOGG"
Call MNU_INFO_RAPID_Click

Case "INFO RAPID TOTAL LOGG"
Call MNU_INFO_RAPID_TOTAL_Click

Case "TEXTCRAWLER ALL FOLDERS"
Me.WindowState = 1
iTX = "/I """ + App.Path + "\# DATA\"
Shell "C:\Program Files (x86)\TextCrawler Pro\TextCrawler.exe " + iTX, vbMaximizedFocus

Case "TEXTCRAWLER TRIM LOGG"
Me.WindowState = 1
iTX = "/FN """ + App.Path + "\# DATA\" + GetComputerName + "\00_ClipBoard_Total--TRIM.txt"
Shell "C:\Program Files (x86)\TextCrawler Pro\TextCrawler.exe " + iTX, vbMaximizedFocus

Case "TEXTCRAWLER STRIP LOGG"
Me.WindowState = 1
iTX = "/FN """ + PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Tot_Strip.txt"
Shell "C:\Program Files (x86)\TextCrawler Pro\TextCrawler.exe " + iTX, vbMaximizedFocus


Case "FILELOCATOR ALL FOLDERS"
Me.WindowState = 1
iTX = "-d """ + App.Path + "\# DATA" + """ -fed -f ""ClipBoard-*.TXT"" -c"""
Shell "C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe " + iTX, vbMaximizedFocus
'-d
'Directory(s) to search (i.e. the Look In field), use -dw for current working directory
'-f
'File names to search for
'-fe?
'File name expr type: -fed=DOS, -fex=Regular Expression, -feb=Boolean, -fee=Plain Text, -feh=Boolean RegEx, -few=Whole Word, -fez = Fuzzy
'-fm
'Match case when comparing file names
' -----------------------------------------------------------
' -c -- NONE CONTENT IF FIELD EXIST FROM BEFORE CLEAR IT GONE


Case "FILELOCATOR IMAGE SET"
Me.WindowState = 1
ART_FOLDER = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName
ART_FOLDER = "D:\0 00 ART LOGGERS\# APP AND SCREEN"
iTX = "-d """ + ART_FOLDER + """ -fed -f ""*.JPG"" -c """
Shell "C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe " + iTX, vbMaximizedFocus
' -c -- NONE CONTENT IF FIELD EXIST FROM BEFORE CLEAR IT GONE

Case "FILELOCATOR IMAGE SET AUTO"
Me.WindowState = 1
ART_FOLDER = "D:\0 00 Art Loggers\# APP AND SCREEN AUTO\4-ASUS-GL522VW_AUTO"
iTX = "-d """ + ART_FOLDER + """ -fed -f ""*.JPG"" -c """
Shell "C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe " + iTX, vbMaximizedFocus
' -c -- NONE CONTENT IF FIELD EXIST FROM BEFORE CLEAR IT GONE

Case "FILELOCATOR IMAGE SET AUTO ALL COMPUTER"
Me.WindowState = 1
ART_FOLDER = "D:\0 00 Art Loggers\# APP AND SCREEN AUTO"
iTX = "-d """ + ART_FOLDER + """ -fed -f ""*.JPG"" -c """
Shell "C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe " + iTX, vbMaximizedFocus


Case "HEX OPEN RECENT TRIM LOGG"
Call Mnu_Open_Recent_Hex_Click

Case "EDIT RECENT TRIM LOGG"
Call Mnu_Open_Recent_Click

Case "EDIT THIS LOGG"
Call Mnu_Open_Logg_Click

Case "EDIT TOTAL LOGG"
Call Mnu_Open_Total_Click

Case "EDIT STRIP LOGG"
Call Mnu_StripLogg_Click

Case "TEXT VIEW RECENT TRIM LOGG"
Call Mnu_TEXT_Open_Recent_Click

Case "TEXT VIEW THIS LOGG"
Call Mnu_TEXT_Open_Logg_Click

Case "TEXT VIEW TOTAL LOGG"
Call Mnu_TEXT_Open_Total_Click

Case "TEXT VIEW STRIP LOGG"
Call Mnu_TEXT_StripLogg_Click

Case "SHELL T -- THIS LOGG"
Call Mnu_ShellT_Todays_Click

Case "SHELL T -- TOTAL LOGG"
Call Mnu_Shell_T_Click

Case "----"

End Select


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
    
    If GetMenuBarInfo(Me.hWnd, OBJID_MENU, 0, MenuInfo) Then
   
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

'Sub SET_MENU_PADD_WORK()
'
'Dim i_Menu_Count, i_Form_Counter
'Dim i_Menu_Not_Visa_Count
'
'Dim Control As Control, LABEL_44, LABEL_48
'
'Dim R_NEXT
'
'Dim Text_Checker_Form_Menu As String
'
'Dim MENU_ITEM_VAR
'
'For i = 0 To Forms.Count - 1
'
'    For Each Control In Forms(i).Controls
'        If InStr(UCase(Control.Name), "MNU_") > 0 Then
'            If Control.Visible = True Then
'                i_Menu_Count = i_Menu_Count + 1
'            End If
'            i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count + 1
'
'        End If
'    Next
'Next
'
'i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count - i_Menu_Count
'
'Mnu_Menu_Item_Count.Caption = "Menu Item Count = " + Str(i_Menu_Count) + " &&" + Str(i_Menu_Not_Visa_Count) + " Not Visible"
''Mnu_Form_Count.Caption = "Form Counter = " + Str(Forms.Count - 1) '  + " Really 7"
''Mnu_Form_Count.Visible = False
'
'i_Menu_Count = 0
'
'
'
'For i = 0 To Forms.Count - 1
'    Text_Checker_Form_Menu = ""
'    FRMLISTMENU01.GetMenuInfo_Not_Indented GetMenu(Forms(i).hWnd), 0, "", Text_Checker_Form_Menu
'    Text_Checker_Form_Menu = UCase(Text_Checker_Form_Menu)
'    For Each Control In Forms(i).Controls
'        If InStr(UCase(Control.Name), "MNU_") > 0 Then
'            MENU_ITEM_VAR = Replace(Control.Caption, "[__ ", "")
'            MENU_ITEM_VAR = Replace(MENU_ITEM_VAR, " __]", "")
'            MENU_ITEM_VAR = UCase(Trim(MENU_ITEM_VAR))
'            If InStr(Text_Checker_Form_Menu, "SUB MENU ----" + MENU_ITEM_VAR) = 0 Then
'
'                'i_Menu_Count = i_Menu_Count + 1
'                If InStr(Trim(Control.Caption), "[__ ") = 0 Then
'                    LABEL_44 = Trim(Control.Caption)
'                    'LABEL_48 = Replace(LABEL_44, " ", "_")
'                    LABEL_48 = LABEL_44
'                    LABEL_48 = Replace(LABEL_48, "___", "__")
'                    LABEL_48 = "[__ " + LABEL_48 + " __]"
'                    LABEL_48 = Replace(LABEL_48, "[__ [__ ", "[__ ")
'                    LABEL_48 = Replace(LABEL_48, " __] __]", " __]")
'                    If LABEL_48 <> LABEL_44 Then
'                        Control.Caption = LABEL_48
'                    End If
'                End If
'            End If
'        End If
'    Next
'Next
'
''Stop
'
'''MNU_BRing_Front
'''---------------
''i_Form_Counter = Forms.Count
''i_Form_Counter = 0
'''for each f
'''For i = 0 To Forms.Count - 1
'''    Load Forms(i)
'''    Show Forms(i)
'''Next
''
''Dim Form As Form
''i_Form_Counter = 0
''For Each Form In Forms
''    i_Form_Counter = i_Form_Counter + 1
''    Load Form
''    Form.Show
''    Show Form
''    'Set Form = Nothing
''Next Form
''
''i_Form_Counter = 0
''For Each Form In Forms
''    i_Form_Counter = i_Form_Counter + 1
''    Load Form
''    Form.Show
''    'Set Form = Nothing
''Next Form
'
'i_Form_Counter = Forms.Count - 1
'Me.Refresh
'End Sub
'
'
'

Private Sub Mnu_Open_Logg_Click()
'Shell "notepad " + PATH_TO_CLIPBOARD_TEXT_INFO_APP_DAY_DATA+"\ClipBoard-" + Format$(Start, "YYYY-MM-DD") + ".Txt", vbNormalFocus


'WON'T DO THIS
'Shell "EXPLORER " + PATH_TO_CLIPBOARD_TEXT_INFO_APP_DAY_DATA+"\ClipBoard-" + Format$(Start, "YYYY-MM-DD") + ".Txt", vbNormalFocus
'OR THIS
'Shell "CMD /C ""START /MAX """ + PATH_TO_CLIPBOARD_TEXT_INFO_APP_DAY_DATA+"\ClipBoard-" + Format$(Start, "YYYY-MM-DD") + ".Txt""""", vbNormalFocus

'ANSWER
vFile = PATH_TO_CLIPBOARD_APP_DATA_1 + "\Day-Data\ClipBoard-" + Format$(start, "YYYY-MM-DD") + ".Txt"

Me.WindowState = vbMinimized

ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal

End Sub

Private Sub Mnu_Open_Recent_Click()
Me.WindowState = vbMinimized

'Shell "notepad " + App.Path + "\# DATA\"+GetComputerName + "\Data\00_ClipBoard_Total--TRIM.txt", vbNormalFocus
vFile = PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total--TRIM.txt"
ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal

End Sub

Private Sub Mnu_Open_Recent_Hex_Click()
Me.WindowState = vbMinimized

'C:\Program Files\XVI32\XVI32.exe
Shell "C:\Program Files\XVI32\XVI32.exe """ + PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total--TRIM.txt""", vbNormalFocus


End Sub

Private Sub Mnu_Open_Total_Click()
Me.WindowState = vbMinimized

'Shell "notepad " + App.Path + "\# DATA\"+GetComputerName + "\Data\00_ClipBoard_Total.txt", vbNormalFocus

vFile = PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total.txt"
ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal

End Sub




Private Sub MNU_SCREEN_SHOT_Click()

'Beep

Me.WindowState = vbMinimized

ART_FOLDER = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName

Shell "Explorer.exe /e, " + """" + ART_FOLDER + """", vbMaximizedFocus

'If Dir("M:\0 00 Art Loggers\", vbDirectory) <> "" And 1 = 2 Then
    'Shell "Explorer.exe /e, M:\0 00 Art Loggers\", vbNormalFocus
'End If

'Me.WindowState = vbMinimized

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

EE$ = GetShortName(PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total.txt")
Shell "D:\Utils\T.com " + EE$, vbNormalFocus
End Sub

Private Sub Mnu_ShellT_Todays_Click()

If Dir("D:\Utils\T.com") = "" Then
    MsgBox "D:\UTILS\T.COM" + vbCrLf + "PROGRAM DON'T EXIST" + vbCrLf + " I SET THIS MENU OPTION TO USE " + vbCrLf + "LIST Enhanced v2.4y1 (c) 1990-2005 Vernon D. Buerg - All rights reserved" + vbCrLf + "Matthew Lancaster, Single-User, s/n ######" + vbCrLf + "Unauthorized duplication prohibited." + vbCrLf + "A FREE VERSION YOU CAN GET - ITS A TEXT BASED VIEWER FROM DOS6.22 DAYS AND HANDLES LONG FILE NAMES" + vbCrLf + "I CALL IT RENAMED FROM LIST.COM TO T.COM -- T FOR TEXT -- IN DIRECTORY D:\UTILS\T.COM", vbMsgBoxSetForeground
    Exit Sub
End If

vFile = PATH_TO_CLIPBOARD_APP_DATA_1 + "\Day-Data\ClipBoard-" + Format$(start, "YYYY-MM-DD") + ".Txt"

'EE$ = GetShortName(App.Path + "\# DATA\"+GetComputerName + "\Data\#ClipBoard.Txt")

EE$ = GetShortName(vFile)
Shell "D:\Utils\T.com " + EE$, vbNormalFocus

End Sub


Private Sub MNU_SOUND_OPTION_Click(Index As Integer)

'THIS GET RUN HERE AT CLICK
'AND AT MENU RESET AUDIO
'FIRST RUN
'AND FIND NEW SOUND

If Index > 0 Then
    'If MNU_SOUND_OPTION(Index).Checked = True Then Exit Sub
    If MNU_SOUND_OPTION(Index).Enabled = False Then Exit Sub
    If MNU_SOUND_OPTION(Index).Visible = False Then Exit Sub
End If

For Each Control In MNU_SOUND_OPTION
    If InStr(MNU_SOUND_OPTION(Index).Caption, "SOUND OPTION - 1") > 0 Then
        For Each Control_2 In MNU_SOUND_OPTION
            If InStr(Control_2.Caption, "SOUND OPTION - 1") > 0 Then
                Control_2.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
                Control_2.Checked = False
                Control_2.Enabled = True
            End If
        Next
    End If
    If InStr(MNU_SOUND_OPTION(Index).Caption, "SOUND OPTION - 2") > 0 Then
        For Each Control_2 In MNU_SOUND_OPTION
            If InStr(Control_2.Caption, "SOUND OPTION - 2") > 0 Then
                Control_2.Enabled = False ' ----------- SO IT DOESNT ACCIDENTLY CLICK IT
                Control_2.Checked = False
                Control_2.Enabled = True
            End If
        Next
    End If
Next
' MNU_SOUND_OPTION(Index).Caption = ""
If InStr(MNU_SOUND_OPTION(Index).Caption, "SOUND OPTION - 1") > 0 Then
    MNU_SOUND_OPTION(Index).Checked = True
    MMCONTROL1__FILENAME = App.Path + "\Sound_Wav's" + Mid(MNU_SOUND_OPTION(Index).Caption, InStrRev(MNU_SOUND_OPTION(Index).Caption, "\"))
End If
If InStr(MNU_SOUND_OPTION(Index).Caption, "SOUND OPTION - 2") > 0 Then
    MNU_SOUND_OPTION(Index).Checked = True
    MMCONTROL2__FILENAME = App.Path + "\Sound_Wav's" + Mid(MNU_SOUND_OPTION(Index).Caption, InStrRev(MNU_SOUND_OPTION(Index).Caption, "\"))
End If

MODIFY_SOUND_SELECTION = True

'THIS GET RUN HERE AT CLICK MENU
'AND AT RESET AUDIO
Call SETUP_SOUND_FILE("")

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
    
    vFile = PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Tot_Strip.txt"
    ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
    
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

vFile = PATH_TO_CLIPBOARD_APP_DATA_1 + "\Day-Data\ClipBoard-" + Format$(start, "YYYY-MM-DD") + ".Txt"

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

Me.WindowState = vbMinimized

vFile = PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total--TRIM.txt"
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

vFile = PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total.txt"
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
vFile = PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Tot_Strip.txt"
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
    ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
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
    If InStr(XPOS, URL_PATH, vbCr) > 0 Then
        URL_PATH = Mid(URL_PATH, XPOS, InStr(XPOS, URL_PATH, vbCr) - XPOS)
    Else
        URL_PATH = Mid(URL_PATH, XPOS)
    End If
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
    Call Menu_Clipboard_size(Len(AD$))
    Me.WindowState = vbMinimized
End Sub
Private Sub Mnu_UrlLoad_Click()
    If Mid$(LastURL, 2, 2) = ":\" Then
        Shell "Explorer.exe /e," + LastURL, vbNormalFocus
    Else
        vLaunch (LastURL)
    End If
End Sub
Public Sub vLaunch(vFile As String)
    Dim Tri
    Tri = 1
    If Mid$(vFile, 1, 5) = "Http:" Then Tri = 1
    If Mid$(vFile, 2, 2) = ":\" Then Tri = 2
    If Right$(vFile, 1) = "\" Then Tri = 3
    Select Case Tri
    Case 1
        ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
    Case 2
        vFile = "file:\\" + vFile
        ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
        'Shell "C:\Program Files\Internet Explorer\iexplore.exe " + vFile, vbNormalFocus
    Case 3
        Shell "Explorer.exe /e," + vFile, vbNormalFocus
    End Select
End Sub
Private Sub VB_RUN_NOT_WHEN_IDE_AND_THEN_SHOW()
    ' -----------------------------------------------------------------
    ' IS SHORT CUT NOT SET ADMINISTATOR FOR THIS PARTICALUR CODE IN EXE
    ' MODE IT WILL NOT RUN LOAD ITSELF ON THE VB COMPILER
    ' BUT AND WILL RUN OTHER PROGRAM
    ' AS TALK OTHER PROGRAM ARE OKAY BUT THEN MAYBE SET AS ADMIN
    ' AMSWER GIVEN
    ' AUG 28 SUN
    ' -----------------------------------------------------------------
    If 1 = 1 And IsIDE = True Then
        MNU_VB.Enabled = False
        Beep
        MsgBox "NOT WHEN ISIDE = TRUE"
        Exit Sub
    End If
    Dim ReturnHwnd As Long
    Dim I
    'VB ONLY WANTS THE 1ST OF THE 2 HWND
    'ReturnHwnd = FindWindowPartVB("ClipBoard Logger - Microsoft Visual Basic[")
    '------------------------------------------------
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
    '------------------------------------------------
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
        If MsgBox("ReturnHwnd" + Str(ReturnHwnd) + " -- " + "MeHwnd" + Str(Me.hWnd) + vbCrLf + "VB Code " + vbCrLf + WIN_SPY_NAME + vbCrLf + " already Running - Sure Want to Run VB IDE", vbYesNo) = vbYes Then
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
    
'    Dim CODER_VBP_FILE_NAME_2
'    Dim VB_1, VB_2, VB_3
'    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
'    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
'    If Dir(VB_1) <> "" Then VB_3 = VB_1
'    If Dir(VB_2) <> "" Then VB_3 = VB_2
'    '------------------------------------------------------
'    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
'    '------------------------------------------------------
'    Dim OBJSHELL
'    Set OBJSHELL = CreateObject("Wscript.Shell")
'    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
'    OBJSHELL.RUN """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
'    Set OBJSHELL = Nothing
'    EXIT_TRUE = True
'    Unload Me
End Sub
Private Sub Mnu_VB_ME_Click()
    Call VB_RUN_NOT_WHEN_IDE_AND_THEN_SHOW
    Dim CODER_VBP_FILE_NAME_2
    Dim VB_1, VB_2, VB_3
    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VB_1) <> "" Then VB_3 = VB_1
    If Dir(VB_2) <> "" Then VB_3 = VB_2
    '------------------------------------------------------
    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
    '------------------------------------------------------
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    Call CHECK_INTEGRITY_OF_VISUAL_BASIC_PROJECT_VBP(CODER_VBP_FILE_NAME_2)
    Dim OBJSHELL
    Set OBJSHELL = CreateObject("WSCRIPT.SHELL")
    OBJSHELL.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set OBJSHELL = Nothing
    EXIT_TRUE = True
    Unload Me
    Exit Sub
End Sub
Private Sub MNU_VB_FOLDER_Click()
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub
Sub CHECK_INTEGRITY_OF_VISUAL_BASIC_PROJECT_VBP(CODER_VBP_FILE_NAME_2)
    ' --------------------------------------------------------------------
    ' FIND IF THE IS NOT CORRECT VERSION AND SIMPLE PUT CORRECT
    ' ONE MY COMPUTER SEEM PROBLEM WITH BASIC
    ' TALK HAS WRONG VERISON ONCE EVERY FEW DAY
    ' AND EDIT IT TO WHAT NORM VERISON SUPPOSED TO BE AND FINE
    ' DISCOVERY -- IT SEEM TO HAPPEN JUST AFTER A COMPILE SOMETIME_
    ' [ Monday 15:37:30 Pm_17 June 2019 ]
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
    Dim r, FR
    Dim VAR_STRING As String
    FR = FreeFile
    Open CODER_VBP_FILE_NAME_2 For Binary As FR
        VAR_STRING = Space(LOF(FR))
        Get #FR, , VAR_STRING
    Close FR
    ' --------------------------------------------------------------
    ' # .VBP # .VBP" -- # MSCOMC -- # OCX -- # MSCOMCTL # CTL # .CTL
    ' --------------------------------------------------------------
    ' FOUND TO BE ERROR
    ' --------------------------------------------------------------
    ' Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0; MSCOMCTL.OCX
    ' --------------------------------------------------------------
    ' CORRECT VALUE
    ' --------------------------------------------------------------
    ' Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0; MSCOMCTL.OCX
    ' --------------------------------------------------------------
    XR1 = InStr(UCase(VAR_STRING), UCase("A1}#2.1; MSCOMCTL.OCX"))
    XR2 = InStr(UCase(VAR_STRING), UCase("A1}#2.#")) ' ---- "A1}#2.#0; mscomctl.OCX"
    XR3 = InStr(UCase(VAR_STRING), UCase("#0; MSCOMCTL.OCX"))
    If XR1 = 0 And XR2 > 0 And XR3 > 0 Then
        GET_01 = InStr(UCase(VAR_STRING), UCase("MSCOMCTL.OCX")) - XR2
        Mid(VAR_STRING, XR2, GET_01 + Len("MSCOMCTL.OCX")) = "A1}#2.1#0; MSCOMCTL.OCX"
        ' MsgBox "MSCOMCTL.OCX" + vbCrLf + "WRONG VERSION -- CHANGE TO" + vbCrLf + "#2.1# -- MSCOMCTL.OCX"), vbMsgBoxSetForeground
        GO_NEXT_IN = True
    End If
    If GO_NEXT_IN = True Then
        If Dir(CODER_VBP_FILE_NAME_2) <> "" Then
            Kill CODER_VBP_FILE_NAME_2
        End If
        FR = FreeFile
        Open CODER_VBP_FILE_NAME_2 For Binary As FR
            Put #FR, , VAR_STRING
        Close FR
    End If
End Sub
Private Sub MNU_FOLDER_DIGITAL_STILL_CAMERA_Click()
    ' Beep
    Me.WindowState = vbMinimized
    Dim FOLDER_Path
    FOLDER_Path = "D:\DSC\2015+SONY\2020 CyberShot HX60V\JPG"
    Dir1.Path = FOLDER_Path
    Dim TD
    For r = 0 To Dir1.ListCount - 1
        FIND_PATH_1 = Dir1.List(r)
        FIND_PATH_2 = Mid(FIND_PATH_1, InStrRev(FIND_PATH_1, "\"))
        If Mid(FIND_PATH_2, 1, 3) = "\20" Then
            TD = Dir1.List(r)
        End If
    Next
    FOLDER_Path = TD
    Shell "EXPLORER /SELECT, """ + FOLDER_Path + """", vbNormalFocus
    EXIT_TRUE = True
    ' Unload Me
    'End
End Sub
Private Sub MNU_CAMERA_VIDEO_FOLDER_4G_Click()
    Me.WindowState = vbMinimized
    Shell "EXPLORER \\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY_MP4\2020 CyberShot HX60V __ MP4", vbMaximizedFocus
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    ' Beep
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
End Sub
'-------------------------------------------
'-------------------------------------------
'---------------------------------------------------
'Public Type SHITEMID
'    cb As Long
'    abID As Byte
'End Type
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'Function GetSpecialfolder(ByVal CSIDL As Long) As String
'    Dim R As Long
'    Dim IDL As ITEMIDLIST
'    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
'    If R = NOERROR Then
'        Path$ = Space$(512)
'        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
'        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
'        Exit Function
'    End If
'    GetSpecialfolder = ""
'End Function
'---------------------------------------------------
'-------------------------------------------
'-------------------------------------------
Public Function GetSpecialFolder_Show_Script_Debug(CSIDL As Long) As String
Dim r As Long
On Error Resume Next
For r = 0 To 120
    If Trim(GetSpecialfolder(r)) <> "" Then
        'Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
        'AAX = GetSpecialfolder(R)
    End If
Next
End
End Function
Public Function GetSpecialfolder(CSIDL As Long) As String
    '##############################################################################################
    'Returns the Path to a "Special" Folder (i.e. Internet History)
    '##############################################################################################
    Dim IDL As ITEMIDLIST
    Dim lResult As Long
    Dim sPath As String
    lResult = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If lResult = 0 Then
        sPath = Space$(512)
        lResult = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        GetSpecialfolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End Function













Private Sub Text1_DblClick()

Call Mnu_Max_Click

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

'F3
If KeyCode = 114 And Shift = 0 Then
    SEARCH_F3 = "F3_NEXT"
    Call Txt_Search_Change
End If

'F3 AND SHIFT
If KeyCode = 114 And Shift = 1 Then
    SEARCH_F3 = "F3_REVERSE"
    Call Txt_Search_Change
End If

' --------------------------------
' SHIFT AND PAGE DOWN -- INITAL
If KeyCode = 16 And Shift = 1 Then
    Exit Sub
End If
' --------------------------------
' SHIFT AND PAGE DOWN -- ANOTHER
If KeyCode = 35 And Shift = 1 Then
    Exit Sub
End If
' --------------------------------


RRR = Now + TimeSerial(0, 0, 3) 'STOPS THE ENTRY IN LOG 'SCROLL IF ABOVE 0 OR LOWER > TIME

If IsIDE = True And Date = "02/06/2020" Then
    If KeyCode = 27 Then Beep: Unload Me: Exit Sub
End If

' Text1.ToolTipText = Str(KeyCode) + "--" + Str(Shift)


' MsgBox Shift

' 2 = CTRL ---- 67 = C
If KeyCode = 67 And Shift = 2 Then
    Me.WindowState = vbMinimized
End If


'Text1.ToolTipText = Str(KeyCode) + "--" + Str(Shift)



' EDITOR FUNCTIONS BEGIN
' ----------------------

SET_GO = False
If KeyCode > 32 And KeyCode < 127 Then SET_GO = True

' ARROW KEYS

If KeyCode = 37 Then Exit Sub ' LEFT
If KeyCode = 38 Then Exit Sub ' UP
If KeyCode = 40 Then Exit Sub ' DOWN
If KeyCode = 39 Then Exit Sub ' RIGHT
If KeyCode = 114 Then Exit Sub ' F3 NEXT SEARCHER

If Shift = 2 Then Exit Sub


If KeyCode = 8 Then
    If Len(Txt_Search.Text) > 0 Then
        If Txt_Search.SelStart > 0 Then
            TEXT_TEMP = Txt_Search.Text
            iPos = Txt_Search.SelStart
            Txt_Search.Text = Mid(TEXT_TEMP, 1, iPos - 1) + Mid(TEXT_TEMP, iPos + 1)
            Txt_Search.SelStart = iPos - 1
        End If
    End If
    Txt_Search.SetFocus
End If

' MsgBox Str(KeyCode) + "--" + Str(Shift)

'DEL KEY
If KeyCode = 46 Then
    SET_GO = False
    If Len(Txt_Search.Text) > 0 Then
        If Txt_Search.SelStart < Len(Txt_Search.Text) Then
            TEXT_TEMP = Txt_Search.Text
            iPos = Txt_Search.SelStart
            Txt_Search.Text = Mid(TEXT_TEMP, 1, iPos) + Mid(TEXT_TEMP, iPos + 2)
            Txt_Search.SelStart = iPos
        End If
    End If
    Txt_Search.SetFocus
End If


' TAB
If KeyCode = 9 Then
    If SEARCH_BOX_NEVER_BEFORE_FOCUS = True Then
        SEARCH_BOX_NEVER_BEFORE_FOCUS = False
        Txt_Search = ""
    End If
End If

If SET_GO = True Then
    iPos = Txt_Search.SelStart
    Txt_Search.Text = Txt_Search.Text + Chr(KeyCode)
    Txt_Search.SetFocus
    Txt_Search.SelStart = iPos + 1
End If

End Sub



Private Sub TEMPOR()

Exit Sub


Dim SI As SCROLLINFO


FlatSB_SetScrollPos Text1.hWnd, SB_VERT, 60, False
SI.cbSize = Len(SI)
SI.fMask = SIF_ALL
FlatSB_GetScrollInfo Text1.hWnd, SB_VERT, SI
SI.nPos = SI.nPos - 10
SI.nPos = 50
FlatSB_SetScrollInfo Text1.hWnd, SB_VERT, SI, True

'Text1.AutoRedraw = True


End Sub

Sub TEXT_SELECTOR_VALUE()

    If O_Text1_SelLength = Text1.SelLength Then Exit Sub

    If MNU_COPY.Caption = "<*** COPY --" + Str(Text1.SelLength) + " ***>" Then Exit Sub
    
    MNU_COPY.Caption = "<*** COPY --" + Str(Text1.SelLength) + " ***>"
    O_Text1_SelLength = Text1.SelLength

End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    STOP_RESIZE_WHILE_MOUSE_OR_KEY_DOWN = True
    
    Call TEXT_SELECTOR_VALUE
    
    If Button <> 1 Then Exit Sub ' Right Click
    RRR = Now + TimeSerial(0, 0, 3) 'STOPS THE ENTRY IN LOG 'SCROLL IF ABOVE 0 OR LOWER > TIME
    

    HHH = Now + TimeSerial(0, 2, 0) ' vbMinimized WINDOW
    
    XXMOUSEMOVE = Now + TimeSerial(0, 2, 0) ' RESET THE SCROLL TO BOTTOM OF TEXT BOX
        
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    STOP_RESIZE_WHILE_MOUSE_OR_KEY_DOWN = False
    
    Call TEXT_SELECTOR_VALUE

    If InStr(MNU_SELECTOR_WHOLE_LINE_MODE.Caption, "RUN_FULL_LINE_SELECT=TRUE") > 0 Then
        AI_1 = Text1.SelStart
        AI_2 = InStrRev(Text1.Text, vbCrLf, AI_1)
        AI_3 = InStr(AI_2 + 1, Text1.Text, vbCrLf)
        Text1.SelStart = AI_2 + 1
        Text1.SelLength = AI_3 - AI_2 - 1
        'Text1.SE
        
        If Text1.SelLength > 2 Then
            Call CLIPBOARD_TEXT_OF_SELECTED_AUTO

            If InStr(MNU_AUTO_MINIMIZE_OFF_ON.Caption, "AUTO MINIMIZE=TRUE") Then
                Me.WindowState = vbMinimized
            End If
        End If
    End If

End Sub


Public Sub CLIPBOARD_TEXT_OF_SELECTED_AUTO()
    '----
    'Copy rich text box's text - Xtreme Visual Basic Talk
    'http://www.xtremevbtalk.com/general/15668-copy-rich-text-boxs-text.html
    '----
    MNU_COPY.Caption = "<*** COPY --" + Str(Text1.SelLength) + " ***>"
    If Text1.SelLength = 0 Then Exit Sub
    EXECUTE_TIMER_ENABLED = False
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then
        ' PUBLIC HERE -- NAME OF SUB WHOLE
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "CLIPBOARD_TEXT_OF_SELECTED_AUTO"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
        Exit Sub
    End If
    Err.Clear
    ' LIKE -- CLIPBOARD.SETTEXT
    SendMessage Text1.hWnd, WM_COPY, 0&, 0& 'Copy

    If Err.Number > 0 Then
        ' PUBLIC HERE -- NAME OF SUB WHOLE
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "CLIPBOARD_TEXT_OF_SELECTED_AUTO"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
        Exit Sub
    End If
    EXECUTE_TIMER_ENABLED = True
    
    AD$ = Mid(Text1.Text, Text1.SelStart + 1, Text1.SelLength)
    
    If MMControl1.Filename = "" Then
        Call RESET_SETUP_SOUND_FILE("NOTSOUND")
    End If

    ' MMControl1.Command = "prev"
    ' MMControl1.Command = "Play"

    Call COMPARE_FOR_EXE
End Sub


Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Exit Sub

    RRR = Now + TimeSerial(0, 0, 3) 'STOPS THE ENTRY IN LOG 'SCROLL IF ABOVE 0 OR LOWER > TIME
    HHH = Now + TimeSerial(0, 2, 0) ' vbMinimized WINDOW
    
    XXMOUSEMOVE = Now + TimeSerial(0, 2, 0) ' RESET THE SCROLL TO BOTTOM OF TEXT BOX
    
    
    Call TEXT_SELECTOR_VALUE
    TRIPPING = " Byte's Selected"
    If Text1.SelLength <> OLD_VAR_TRIPPING Then
        OLD_VAR_TRIPPING = Text1.SelLength
        
        If MNU_SELECTED.Caption <> Trim(Str(Text1.SelLength)) + TRIPPING Then
            MNU_SELECTED.Caption = Trim(Str(Text1.SelLength)) + TRIPPING
        End If
    End If
    
    
    'MNU_SELECTED


End Sub


Sub GetFormat_And_Display()

'If FirstRun = False Then
'    Mnu_Run_Status.Visible = False
'    Mnu_Clip_Status.Visible = False
'End If

On Error Resume Next

I = False
ClipFormatDescription = ""
ClipFormatDesc2 = ""

Do
    Err.Clear
    If I = False Then
        I = Clipboard.GetFormat(vbCFText)
        '"Clip Format:- "+"Text (.txt file)"
        If I = True And ClipFormatDescription = "" Then
            ClipFormatDescription = "Text (.txt file)"
            ClipFormatDesc2 = "Text"
        End If
    End If
Loop Until Err.Number = 0
' Check for Text Before Rich Text Better Option - But Check if You Can Grab Rich Text into Plain Text

' DON'T REQUIRE THIS IF AS IF I=FALSE DO THAT WORK
If ClipFormatDesc2 <> "Text" Then
    ' Two Most Important First
    Do
        Err.Clear
        If I = False Then
            I = Clipboard.GetFormat(vbCFBitmap)
            If I = True And ClipFormatDescription = "" Then
                ClipFormatDescription = "Bitmap (.bmp file)"
                ClipFormatDesc2 = "Image"
            End If
        End If
    Loop Until Err.Number = 0
End If

Do
    Err.Clear
    If I = False Then
        I = Clipboard.GetFormat(vbCFRTF)
        If I = True And ClipFormatDescription = "" Then
            ClipFormatDescription = "Rich Text Format (.rtf file)"
            ClipFormatDesc2 = "Rich Text Format File"
        End If
    End If
Loop Until Err.Number = 0

Do
    Err.Clear
    If I = False Then
        I = Clipboard.GetFormat(vbCFLink)
        If I = True And ClipFormatDescription = "" Then
            ClipFormatDescription = "DDE Conversation information"
            ClipFormatDesc2 = "DDE Conversation info"
        End If
    End If
Loop Until Err.Number = 0

Do
    Err.Clear
    If I = False Then
        I = Clipboard.GetFormat(vbCFMetafile)
        If I = True Then
            ClipFormatDescription = "Metafile (.wmf file)"
            ClipFormatDesc2 = "Metafile .WMF"
        End If
    End If
Loop Until Err.Number = 0

Do
    Err.Clear
    If I = False Then
        I = Clipboard.GetFormat(vbCFEMetafile)
        If I = True And ClipFormatDescription = "" Then
            ClipFormatDescription = "Device-independent bitmap - E Type"
            ClipFormatDesc2 = "Device-independent bitmap - E Type"
        End If
    End If
Loop Until Err.Number = 0

Do
    Err.Clear
    If I = False Then
        I = Clipboard.GetFormat(vbCFDIB)
        If I = True And ClipFormatDescription = "" Then
            ClipFormatDescription = "Device-independent bitmap"
            ClipFormatDesc2 = "Device-independent bitmap"
        End If
    End If
Loop Until Err.Number = 0

Do
    Err.Clear
    If I = False Then
        I = Clipboard.GetFormat(vbCFPalette)
        If I = True And ClipFormatDescription = "" Then
            ClipFormatDescription = "Color palette"
            ClipFormatDesc2 = "Color Palette"
        End If
    End If
Loop Until Err.Number = 0

Do
    Err.Clear
    If I = False Then
        I = Clipboard.GetFormat(vbCFFiles)
        If I = True And ClipFormatDescription = "" Then
            ClipFormatDescription = "File list from Windows Explorer"
            ClipFormatDesc2 = "File list from Windows Explorer"
        End If
    End If
Loop Until Err.Number = 0

Do
    Err.Clear
    If I = False Then
        I = Clipboard.GetFormat(-15694) 'Pasted Objects on VB IDE Form Designer
        If I = True And ClipFormatDescription = "" Then
            ClipFormatDescription = "VB IDE (Form Designer) Object -15694"
            ClipFormatDesc2 = "VB IDE (Form Designer) Object -15694"
        End If
    End If
Loop Until Err.Number = 0

'Check the Most Common 1st

'And Then Check All - Don't Let Any Escape

iindex = -32000
If I = False Then
    For r = -30000 To 32000
        If I = False Then
            Do
                Err.Clear
                I = Clipboard.GetFormat(r)
                iindex = r
            Loop Until Err.Number = 0
        Else
            Exit For
        End If
    Next

    If I = True And iindex <> -32000 And ClipFormatDescription = "" Then
        ClipFormatDescription = "Format Unknown #" + Str(iindex) + " of -- Between -30000 and 32000"
        ClipFormatDesc2 = "Format Unknown #" + Str(iindex) + " of -- Between -30000 and 32000"
    End If
End If

If I = False Then
    ClipFormatDescription = "Empty Clipboard"
    ClipFormatDesc2 = "Empty Clipboard"
End If

On Error GoTo EXIT_OUT
Sleep 100
'If Clipboard.GetFormat(vbCFText) Then
If ClipFormatDesc2 = "Text" Then
    For r = 2000 To 0 Step -1
        On Error Resume Next
        Err.Clear
        If Len(Clipboard.GetText) = 0 Then
            ClipFormatDescription = "Text And Len(0) Empty Clipboard"
            ClipFormatDesc2 = "Text And Len(0) Empty Clipboard"
        End If
        If Err.Number > 0 Then Sleep 100
        If Err.Number = 0 Then Exit For
    Next
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

Exit Sub
EXIT_OUT:



End Sub


Sub COMPARE_FOR_EXE()
    ' COMPARE LAST TWO CLIPBOARD FOR EXE PROGRAM TO RUN WINCOMPARE
    '-------------------------------------------------------------
    
    On Error GoTo EXIT_OUT
        
    If Clipboard.GetFormat(vbCFText) = False Then Exit Sub
    
    Do
        Err.Clear
        COMPARE_EXE_3 = Clipboard.GetText
        If Err.Number > 0 Then GoTo EXIT_OUT
    Loop Until Err.Number = 0

    On Error GoTo 0
            
    COMPARE_EXE_2 = COMPARE_EXE_1
    
    COMPARE_EXE_1 = COMPARE_EXE_3
    
    MNU_COMPARE.Caption = "COMPARE" + Str(Len(COMPARE_EXE_1)) + " " + Str(Len(COMPARE_EXE_2))
    
    COMPARE_EXE_3 = ""
    
    If COMPARE_EXE_1 = COMPARE_EXE_2 Then
        MNU_COMPARE.Enabled = False
    Else
        MNU_COMPARE.Enabled = True
    End If
    
    
Exit Sub


EXIT_OUT:
    
    CLIPBOARD_RETURN_TIMER_ERROR_3 = 1
    Timer_MNU_COPY_2.Enabled = True
    
End Sub

Sub ICON_NOT_WANTING()

'D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\# ICON NOT WANTING ON CLIPBOARD

File1.Path = App.Path + "\# DATA\# ICON NOT WANTING ON CLIPBOARD"
ReDim ARRAY_ICON_NOT_WANT(1)
COUNT_HIGH_UP = 0
For R_LOOP_8 = 1 To 2
    If R_LOOP_8 = 1 Then File1.Pattern = "*.BMP"
    If R_LOOP_8 = 2 Then File1.Pattern = "*.JPG"

    
    For r = 0 To File1.ListCount - 1
        COUNT_HIGH_UP = COUNT_HIGH_UP + 1
        ReDim Preserve ARRAY_ICON_NOT_WANT(COUNT_HIGH_UP)
        FR1 = FreeFile
        Open File1.Path + "\" + File1.List(r) For Binary As #FR1
            ARRAY_ICON_NOT_WANT(COUNT_HIGH_UP) = Input(50, FR1)
        Close #FR1
    Next
Next

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------

ReDim ARRAY_ICON_DE_DUPE_VAR_(2)
ReDim ARRAY_ICON_DE_DUPE_FILE(2)
'01 OF 03 LOAD IN
COUNT_HIGH_UP = 0

For R_LOOP_3 = 1 To 2
    If R_LOOP_3 = 1 Then
        File1.Path = App.Path + "\# DATA\# ICON NOT WANTING ON CLIPBOARD"
    End If
    If R_LOOP_3 = 2 Then
        File1.Path = App.Path + "\# DATA\# VBIcon"
    End If
    
    For R_LOOP_8 = 1 To 2
        If R_LOOP_8 = 1 Then File1.Pattern = "*.BMP"
        If R_LOOP_8 = 2 Then File1.Pattern = "*.JPG"

        For R_LOOP = 0 To File1.ListCount - 1
            COUNT_HIGH_UP = COUNT_HIGH_UP + 1
            ReDim Preserve ARRAY_ICON_DE_DUPE_VAR_(COUNT_HIGH_UP)
            ReDim Preserve ARRAY_ICON_DE_DUPE_FILE(COUNT_HIGH_UP)
            PATH_BMP = File1.Path + "\" + File1.List(R_LOOP)
            ARRAY_ICON_DE_DUPE_FILE(COUNT_HIGH_UP) = PATH_BMP
            FR1 = FreeFile
            Open File1.Path + "\" + File1.List(R_LOOP) For Binary As #FR1
                ARRAY_ICON_DE_DUPE_VAR_(COUNT_HIGH_UP) = Input(LOF(FR1), FR1)
            Close #FR1
            
        Next
    Next
Next

'02 OF 03 COMPARE
For R_LOOP_1 = UBound(ARRAY_ICON_DE_DUPE_VAR_) To 0 Step -1
For R_LOOP_2 = UBound(ARRAY_ICON_DE_DUPE_VAR_) To 0 Step -1
    PATH_BMP = ARRAY_ICON_DE_DUPE_FILE(R_LOOP_1)
    If PATH_BMP <> "" Then
        If R_LOOP_1 <> R_LOOP_2 Then
            If ARRAY_ICON_DE_DUPE_VAR_(R_LOOP_1) <> "" And ARRAY_ICON_DE_DUPE_VAR_(R_LOOP_2) <> "" Then
                If ARRAY_ICON_DE_DUPE_VAR_(R_LOOP_1) = ARRAY_ICON_DE_DUPE_VAR_(R_LOOP_2) Then
                    KILLCOUNTER = KILLCOUNTER + 1
                    
                    ' Debug.Print KILLCOUNTER
                    
                    Kill PATH_BMP
                    ARRAY_ICON_DE_DUPE_FILE(R_LOOP_1) = ""
                    ARRAY_ICON_DE_DUPE_VAR_(R_LOOP_1) = ""
                End If
            End If
        End If
    End If
Next
Next
'03 OF 03 CLEAR
For R_LOOP_2 = 0 To UBound(ARRAY_ICON_DE_DUPE_VAR_)
    ARRAY_ICON_DE_DUPE_VAR_(R_LOOP_2) = ""
Next
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------



End Sub


Private Sub Timer1_PLUS_Timer()

'5 MILLISEC TIMER

If Timer1_PLUS.Interval <> 1 Then Timer1_PLUS.Interval = 1

Call SETUP_DIMENSIONS_INNER_FORM

Call MNU_LOGGEXPLORER_Click(0)

End Sub

Private Sub TIMER1_SPOT_CHECK_CLIPBOPARD_CORRECT_LENGHT_Timer()
Dim CLIP_INFO_04 As String
TIMER1_SPOT_CHECK_CLIPBOPARD_CORRECT_LENGHT.Interval = 10000
On Error Resume Next
    CLIP_INFO_01 = Trim(Str(1)) + "-ARE-" + Trim(UCase(Str(True))) + "--"
    If Clipboard.GetFormat(1) = False Then
        For R1 = 1 To 20
            CLIP_INFO_01 = CLIP_INFO_01 + Trim(Str(R1)) + "-ARE-" + Trim(UCase(Str(Clipboard.GetFormat(R1)))) + "--"
        Next
    End If
    
    ' ---------------------------------------------------
    '    CONSTANT VALUE DESCRIPTION
    '    vbCFLink     &HBF00 DDE CONVERSATION INFORMATION
    '    vbCFText     1 TEXT
    '    vbCFBitmap   2 BITMAP (.BMP FILES)
    '    vbCFMetafile 3 METAFILE (.WMF FILES)
    '    vbCFDIB      8 DEVICE-INDEPENDENT BITMAP (DIB)
    '    vbCFPalette  9 COLOR PALETTE
    ' ---------------------------------------------------
    ' CLIP_INFO_02 = "LEN OF CLIPBOARD.GETDATA -- " + Trim(Str(Len(Clipboard.GetData)))
    ' Clipboard.GetData -- NOT WORK GOT TO BE HOLD IN VARIABLE AND
    ' PICTURE IMAGE MOST THE TIME
    ' NEW METHOD AFTER
    ' PROJECT MENU >>
    ' COMPONENTS >>>>
    ' MICROSOFT FORMS2.0 OBJECT LIBRARY
    ' C:\Windows\SysWOW64\FM20.DLL
    ' MSForms.DataObject
    ' Dim CLIP As New MSForms.DataObject
    ' ----------------------------------------------------------------------------------------
    If Clipboard.GetFormat(1) = False Then ' ---- CHECK FOR TEXT
        Dim CLIP As New MSForms.DataObject
        CLIP.GetFromClipboard
        ' REF SOURCE INFO END OF PROCEDURE
        CLIP_INFO_02 = "LEN OF CLIPBOARD.GETDATA -- " + Trim(Str(Len(CLIP.GetText)))
    End If
    CLIP_INFO_03 = "LEN OF CLIPBOARD.GETTEXT -- " + Trim(Str(Len(Clipboard.GetText)))
    
    If Err.Number > 0 Then
        Exit Sub
    End If
On Error GoTo 0

CLIP_INFO_04 = CLIP_INFO_01 + CLIP_INFO_02 + CLIP_INFO_03

If OLD_CLIP_INFO_04 <> CLIP_INFO_04 Then
    ' CLIPBOARD HAS CHANGE
    ' CHECK TIME STAMP
    TIME_STAMP_CLIPBOARD = Now
End If
If MISSING_PULSE_PULSE_TRAIN <> 0 Then
    X_TIME = Abs(DateDiff("s", TIME_STAMP_CLIPBOARD, Timer_API_Test_Logick_Var1))
    X_TIME = Abs(DateDiff("s", TIME_STAMP_CLIPBOARD, MISSING_PULSE_PULSE_TRAIN))
    Debug.Print X_TIME
    If X_TIME > 120 Then
        If FIRST_RUN_SPOT_CHECK_CLIPBOARD = False Then
            ' REPORT ERROR CLIPBOARD NOT PICKUP
            ' 1ST TRY ALL RESET
            ' REQUIRE VBSCRIPT TO RELOAD BETTER
            ' CLOSE AND RELOAD
            Call VBSCRIPT_CLOSE_AND_RELOAD
        End If
        FIRST_RUN_SPOT_CHECK_CLIPBOARD = False
    End If

'    Debug.Print X_TIME
End If
OLD_CLIP_INFO_04 = CLIP_INFO_04

' -------------------------------------------------------------------------
' EXCEL - COPY CLIPBOARD DATA TO ARRAY - STACK OVERFLOW
' https://stackoverflow.com/questions/33156317/copy-clipboard-data-to-array
' -------------------------------------------------------------------------
' SAT 03-OCT-2020 15:10:30
' -------------------------------------------------------------------------
End Sub

Sub VBSCRIPT_CLOSE_AND_RELOAD()
    
    If IsIDE = True Then Exit Sub
    
    EXIT_TRUE = True
    On Error Resume Next
    Me.WindowState = VBMINIMIZE
    On Error GoTo 0
    VBS_LAUNCHER_NAME = "D:\VB6\VB-NT\00_BEST_VB_01\Clipboard Logger\VBS - RELOAD RESTART.VBS"
    Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.Run """" + VBS_LAUNCHER_NAME + """", ShowWindow_2, DontWaitUntilFinished
    Set WSHShell = Nothing
    
    Unload Me
End Sub

Public Sub Timer1_Timer()

Form1.Timer1.Enabled = False

XX_FAULT_ = DateDiff("s", TIMER_MISSING_PULSE_CLIPBOARD_01, TIMER_MISSING_PULSE_CLIPBOARD_02)
' NOT USE ANYWHERE
MISSING_PULSE_PULSE_TRAIN = Now



Dim GetForegroundWindow_XXRT, XXR5, XXR7
Dim ASH2$
Dim ASH$

'DateTimerCheckIntegrityOfProgram = Now
TIME_LAST_CLIPBOARD_Timer1 = Now
CALC_NEW_INPUT = True

TEXT_TO_ENTER_IN_LOGGER = True

VAR_INFO_LOGG_CLIP = ""

TIMER1_LAST_DATE = Now

If FlagTestClipBoardRoutine = True Then
    FlagTestClipBoardRoutine = False
    MsgBox "Clipboarded Trigger Flag Has Found to Be Working Upto Here The Entry to Routine", vbMsgBoxSetForeground
End If

ADTEST_BEFORE$ = ""

'EXECUTE_TIMER_ENABLED = True


'STOPS THE ENTRY IN LOG 'SCROLL IF ABOVE 0 OR LOWER > TIME
'DON'T ACT ON WHAT IS COPIED FROM CLIPBOARD TEXT BOX
'RUN IF TIME > NOW -- OR = MY FORM -- AND ALSO ONLY IF PICTURE IMAGE = FALSE

'CHANGE TO FIRST IS AND

' --------------------------------------------------------------------------------
If True = False Then
    ' ----------------------------------------------------------------------------
    ' AN ERROR TRAP FOR NEXT FOLLOW ON CODE SET
    ' GET CLIPBOARD REPEAT QUICK EVEN WITH DOEVENT AND SLEEP
    ' IS PRESSURE ON THE SYSTEM
    ' ----------------------------------------------------------------------------
EXIT_RETRY_AGAIN_ERROR:
    CLIPBOARD_RETURN_TIMER_ERROR_4 = 1
    Timer_MNU_COPY_2.Enabled = True
    Exit Sub

End If
' --------------------------------------------------------------------------------

OO = False
OO_1 = False
OO_2 = False
On Error GoTo EXIT_RETRY_AGAIN_ERROR

Err.Clear
OO_1 = Clipboard.GetFormat(vbCFText)
If OO_1 = False Then
    OO_2 = Clipboard.GetFormat(vbCFBitmap)
End If

Clipboard_GetFormat_vbCFText_OO_1 = OO_1
Clipboard_GetFormat_vbCFBitmap_OO_2 = OO_2

If Err.Number > 0 Then
    GoTo EXIT_RETRY_AGAIN_ERROR
End If

If OO_1 = True Or OO_2 = True Then OO = True

On Error GoTo 0


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
If Clipboard_Clear_SUPPRESS_MESSAGE = True Then
    If ClipFormatDescription = "Empty Clipboard" Then
        Exit Sub
    End If
End If

If OO_1 = True Then
    Call CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY(False)
End If

'WHEN GRAB TEXT FROM OWN WINDOW ClipBoard Logger
'TREAT IT AS A SAME COPY AND NOT WANTED AND SOUND EFFECT AS SAME #2
'RRR = + TIME WHEN SET
'THIS DONT SEEM TO ENTER AND RUN AND IS CALLED LATER
'YES WHEN WE DON'T WANT IT IT IS RUN

SET_GO = False
If OO_1 = True Or OO_2 = True Then SET_GO = True
If (RRR > Now And GetForegroundWindow = Me.hWnd) = True Then SET_GO = False
If ENTER_TEXT_IN_LOGGER_FOREGROUND_OVERRIDE = True Then SET_GO = True
ENTER_TEXT_IN_LOGGER_FOREGROUND_OVERRIDE = False

If SET_GO = False Then
    On Error Resume Next
    COUNTER_ERROR = 0
    Do
        Err.Clear
        If OO_1 = True Then
            AD$ = Clipboard.GetText
        End If
        If Err.Number > 0 Then Sleep 10
        COUNTER_ERROR = COUNTER_ERROR + 1
        'If COUNTER_ERROR Mod 200 = 0 Then DoEvents
        If COUNTER_ERROR > 1000 Then
            Exit Sub
        End If
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    If OO = True Then
        AD_DATE = Now
        Call Menu_Clipboard_size(Len(AD$))
    End If
    RRR = 0
    If Mnu_SoundOn.Checked = True Then
        On Error Resume Next
        If MMControl2.Filename = "" Then
            Call RESET_SETUP_SOUND_FILE("NOTSOUND")
        End If

'        If FindWindow("TV_CClientWindowClass", vbNullString) > 0 Then
'            On Error Resume Next
'            FR1 = FreeFile
'            Open App.Path + "\# DATA\" + GetComputerName + "\TEAMVIEWER.TXT" For Output As #FR1
'                Print #FR1, Format(Now, "DD/MM/YYYY HH:MM:SS")
'            Close #FR1
'            Set WSHShell = CreateObject("WScript.Shell")
'            WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 24-TEAMVIEWER COPY FILE OVER NETWORK.ahk"""
'            Set WSHShell = Nothing
'        End If
        
        
'        MMControl2_GFW_Var = GetForegroundWindow
'        If MMControl2_GFW_Var <> O_MMControl2_GFW_Var Then
'            MMControl2_Counting_Var = 0
'        End If
'        O_MMControl2_GFW_Var = MMControl2_GFW_Var
        

        'DUPE SOUND
        MMControl2.Command = "prev"
        MMControl2.Command = "Play"
                
        On Error GoTo 0
    End If
    
    Exit Sub

End If


If SET_GO = True Then
    
    On Error Resume Next
    COUNTER_ERROR = 0
    Do
        Err.Clear
        If OO_1 = True Then
            AD$ = Clipboard.GetText
        End If
        If Err.Number > 0 Then Sleep 10
        COUNTER_ERROR = COUNTER_ERROR + 1
        'If COUNTER_ERROR Mod 200 = 0 Then DoEvents
        If COUNTER_ERROR > 1000 Then
            Exit Sub
        End If
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    If OO = True Then
        AD_DATE = Now
        Call Menu_Clipboard_size(Len(AD$))
        

        
        
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
    
    
    If OO_1 = True Then
        If MNU_LAST_GRAB_ALL_CAPS.Checked = True Then
            Call LAST_GRAB_ALL_CAPS_002_Click
        End If
        
        ' MORE WORKER TO DO HERE -- Sat 14-Sep-2019 21:18:00
        ' --------------------------------------------------
        If MNU_LAST_GRAB_PRO_CAPS.Checked = True Then
            ' Call LAST_GRAB_ALL_CAPS_002_Click
        End If
    End If
    
    
    
'-----------------------------------------------------
    
    '-------------------------------------------------------------
    'NOT HERE WE WILL IF
    'DUPE BUT NOT AS TAKE TEXT OFF SCREEN FORM
    'MAYBE
    'BUT ADD
    'DO THE MAKE ALL CAPS RUN
    '------------------------
    '-------------------------------------------------------------
'    If Mnu_SoundOn.Checked = True Then
'        On Error Resume Next
'        MMControl2.Command = "prev"
'        MMControl2.Command = "Play"
'        On Error GoTo 0
'    End If

'-----------------------------------------------------
    
    'VAR_INFO_LOGG_CLIP = "Clip in Own TextBox"
    'TEXT_TO_ENTER_IN_LOGGER = False
    'MNU_ERROR_INFO.Caption = VAR_INFO_LOGG_CLIP
    'MNU_ERROR_INFO.Visible = True
    
    'Exit Sub
    
    
    'Exit Sub
'-----------------------------------------------------

End If


'--------------------------------------------------------
'TEXT IS A DUPE TRUE FALSE
'--------------------------------------------------------
'DUPE AS LAST TIME THAT HAS BEEN LOGGED SEE FURTHER AHEAD
'--------------------------------------------------------
If OO_1 = True Then
    If AD$ = GG$ Then
    '-----------------------------------------------------
        If AD$ <> "" Then
            '------------------
            'PLAY TUNE FOR DUPE
            If Mnu_SoundOn.Checked = True Then
                '----------------------------------------
                'IN WIN 10 IS IF SOUND NOT ABLE TO PLAY HERE
                'AND THEN CAUSE CRASH HOLD FREEZE
                '----------------------------------------
                
                If MMControl2.Filename = "" Then
                    Call RESET_SETUP_SOUND_FILE("NOTSOUND")
                End If

                'SET IDLE TIME FOR TEAMVIEW KEEPER PLAYING WITH CLIPBOARD
                'SO IF AT CONTROL PLAY A SOUND
'                If FindWindow("TV_CClientWindowClass", vbNullString) > 0 Then
'                    MMControl2_Counting_Var = MMControl2_Counting_Var + 1
'                End If
'                MMControl2_GFW_Var = GetForegroundWindow
'                If MMControl2_GFW_Var <> O_MMControl2_GFW_Var Then
'                    MMControl2_Counting_Var = 0
'                End If
'                O_MMControl2_GFW_Var = MMControl2_GFW_Var
                
                'MMControl2_Counting_Var = MMControl2_Counting_Var + 1
                
                If OLD_FOREGROUND_WINDOW_TEAM_VIEWER <> GetForegroundWindow Then
                    'DONT_PLAY_SOUND_TEAM_VIEWER = False
                End If
                
                'SOUND FOR A DUPE
                If Mnu_SoundOn.Checked = True And DONT_PLAY_SOUND_TEAM_VIEWER = False Then
                'If A_TimeIdle + TimeSerial(0, 0, 10) > Now Then
                    MMControl2.Command = "prev"
                    MMControl2.Command = "Play"
                End If
            
                '-------------------------------------------------------------------------------------------
                'THIS WINDOW IS SHOWN ON BOTH REMOTE AND BASE MACHINE
                '-------------------------------------------------------------------------------------------
                If FindWindow("TV_ControlWin", "TeamViewer Panel") > 0 Then
                    DONT_PLAY_SOUND_TEAM_VIEWER = True
                Else
                    DONT_PLAY_SOUND_TEAM_VIEWER = False
                End If
    
            
            
            
            End If
        Else
            'ADD CLIPBOARD ON UNIQUE SOUND PLAYED OR DUPE ELSE WHERE
            'AND PICTURE IS HANDLE ELSE WHERE
            'AND HOPE TEST DUPE FOR THAT WORK
            'BUT TEXT ON PICTURE NEVER BE DUPE
            
        End If
    
    '-----------------------------------------------------
        VAR_INFO_LOGG_CLIP = "Dupe Text__Blocked"
        TEXT_TO_ENTER_IN_LOGGER = False
    '    Exit Sub
    '-----------------------------------------------------
    Else
        'MMControl2_Counting_Var = 0
        
        DONT_PLAY_SOUND_TEAM_VIEWER = False
        
        If Mnu_SoundOn.Checked = True Then
            If MMControl1.Filename = "" Then
                Call RESET_SETUP_SOUND_FILE("NOTSOUND")
            End If
            'SOUND NOT FOR A DUPE
            '--------------------
            If STARTUP_RUN = False Then
                A = MMControl1.Filename
                MMControl1.Command = "prev"
                MMControl1.Command = "Play"
            End If
        End If
        
    End If
End If

OLD_FOREGROUND_WINDOW_TEAM_VIEWER = GetForegroundWindow

'----------------------------------
'----------------------------------
'----------------------------------
GG$ = AD$



'----------------------------------------
'COMPARE TO CALL PROGRAM WINMERGE COMPARE
'----------------------------------------
If OO_1 = True Then Call COMPARE_FOR_EXE



'NOT A TEXT OR PICTURE
If OO = False Then
    TEXT_TO_ENTER_IN_LOGGER = False
    'Exit Sub
End If


'VBICON = False

'VBICON CHECKING

If OO_2 = True Then
    On Error Resume Next
    COUNTER_ERROR = 0
    Do
        Err.Clear
        Picture1.Picture = Clipboard.GetData(vbCFBitmap)
        If Err.Number > 0 Then Sleep 10
        COUNTER_ERROR = COUNTER_ERROR + 1
        'If COUNTER_ERROR Mod 200 = 0 Then DoEvents
        If COUNTER_ERROR > 1000 Then
            Exit Sub
        End If
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    Xx1 = Picture1.Picture.Width
    yy1 = Picture1.Picture.Height
    'hh1 = Picture1.Picture.hPal
    hh2 = Picture1.Picture.Type
    
    If Xx1 > 0 And yy1 > 0 And hh2 = 1 Then
        SET_GO = False
        If Xx1 = 1058 And yy1 = 1058 Then SET_GO = True
        If Xx1 <= 1058 And yy1 <= 1058 Then SET_GO = True
        If Xx1 < 847 And yy1 < 847 Then SET_GO = True
        If Xx1 < 900 And yy1 < 900 Then SET_GO = True
        If SET_GO = True Then
            
            'GOT TO SAVE IT -- IF SMALL
            PATH_BMP = App.Path + "\# DATA\# VBIcon\" + Format(Now, "YYYY-MM-DD--HH-MM-SS") + "_VBIcon.bmp"
            SavePicture Picture1.Picture, PATH_BMP
            
            'OPEN IT AGAIN - TO GET DATA
            FR1 = FreeFile
            Open PATH_BMP For Binary As #FR1
                Pic2$ = Input(50, FR1)
            Close #1
            
            'IF IT IS NOT THE SAME AS OUR RECORD
            'AND A VBICON PICTURE
            
            'SOMETIME THE VBICON CHANGE
            'BUT MOST THE TIME ONLY REQUIRE CHECK AGAINST ONE OR MAYBE LATER JOIN TWO OR THREE TO COMPARE
            
            'XLEN ---- 4000 TO 7000 FOR WIN 10
            VBICON_VAR = False
            For R_LOOP = 0 To UBound(ARRAY_ICON_NOT_WANT)
                If Pic2$ = ARRAY_ICON_NOT_WANT(R_LOOP) Then
                    VBICON_VAR = True
                End If
            Next
            
            If VBICON_VAR = True Then
                
                VB_ICON_TEXT_VAR = "VB IDE ICON CHANGED PROPERTY"
                
                MSGBOX2 = "INFO - " + VB_ICON_TEXT_VAR
                frm_MSGBOX.Timer1.Enabled = True
                
                VB_ICON_CHANGED = True
                VAR_INFO_LOGG_CLIP = VB_ICON_TEXT_VAR
            End If
            
            '--------------------
            'AND TO CHECK IF WANT TEXT PUT BACK WHEN VBICON TRUE
            '---------------------------------------------------
            If VBICON_VAR = True Then
                Do
                    If FindWindow("VBSplash", vbNullString) > 0 Then Sleep 100
                Loop Until FindWindow("VBSplash", vbNullString) = 0
                
                '--------------------------------------------
                'IT IS A VBICON AND WE WANT OUR TEXT PUT BACK
                '--------------------------------------------
                If AD$ <> "" Then
                    
                    'ctlClipboard1_ClipboardChanged
                    EXECUTE_TIMER_ENABLED = False
                    'REMMER
                    On Error Resume Next
                    Do
                        Err.Clear
                        Clipboard.Clear
                    Loop Until Err.Number = 0
                    Do
                        Err.Clear
                        Clipboard.SetText AD$
                    Loop Until Err.Number = 0
                    EXECUTE_TIMER_ENABLED = True
                    
                    
                    Call Menu_Clipboard_size(Len(AD$))
                    If FirstRun = True Then
                        Mnu_Run_Status.Caption = "1st Run"
                    End If

'-----------------------------------------------------
                    
                    VAR_INFO_LOGG_CLIP = "VB_ICON HAPPEN AND OUR TEXT WAS PUT BACK"
                    If VB_ICON_CHANGED = True Then
                        VAR_INFO_LOGG_CLIP = "VB_ICON HAPPEN AND IT CHANGED - AND OUR TEXT WAS PUT BACK"
                    End If
                    
                    TEXT_TO_ENTER_IN_LOGGER = False
                    'VBICON CHECKING
                    Exit Sub
'-----------------------------------------------------
                End If
            End If
        End If
    End If
End If



'GET FIRST RUN STATUS
If OO_1 = True Then
    If FirstRun = True Then
        Mnu_Run_Status.Caption = "1st Run"
    End If
    FirstRun = False
End If

'WHAT DO IF NOTHING
If OO_1 = True Then
    '-------------------------------------------------------
    'THE AD$ GET_CLIPBOARD WAS DONE BEFORE IN THIS PROCEDURE
    '-------------------------------------------------------
    If AD$ = "" Then
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
If OO_2 = True Then
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
            If MMControl2.Filename = "" Then
                Call RESET_SETUP_SOUND_FILE("NOTSOUND")
            End If
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
    
    
If OO_2 = True Then
    If DUPE_IMAGE_AT_FORM_LOAD_VAR2 = False Then
            
        '-------------------------
        '-------------------------
        '-------------------------
        'HERE GET THE PICTURE INFO
        '-------------------------
        '-------------------------
        '-------------------------
        '--------------------------------------------------------------------
        If Mnu_SoundOn.Checked = True Then
            If MMControl1.Filename = "" Then
                Call RESET_SETUP_SOUND_FILE("NOTSOUND")
            End If
            MMControl1.Command = "prev"
            MMControl1.Command = "Play"
        End If
        '--------------------------------------------------------------------

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
If XI1 = True And OO_1 = True Then

    '-------------------------------------------
    'SET EARLIER BECUASE OTHER EVENT WANT TO SEE
    '-------------------------------------------
    LimitClipSize = LimitClipSize
    '-------------------------------------------
    
    MNU_ENTER_LARGE.Caption = "ENTER LARGE TEXT IN LOGGER OR REPEAT ARE BLOCKED - Hardcoded ""BluetoothLogView"" Is Allow -- The Limit" + Str(Int(LimitClipSize / 1000)) + " K" + " and Current =" + Str(Len(AD$))
    If Len(AD$) > LimitClipSize Then
        MNU_ENTER_LARGE.Caption = MNU_ENTER_LARGE.Caption + " = Answer Not Okay"
        MNU_CB_SIZE_BYTE_OVERSIZE.Visible = True
        MNU_CB_SIZE_BYTE_OVERSIZE.Caption = "CLIP OVERSIZE LIMIT -" + Str(Val(VarSize)) + ">" + Trim(Str(Int(LimitClipSize / 1000))) + " K"
    Else
        MNU_ENTER_LARGE.Caption = MNU_ENTER_LARGE.Caption + " = Answer Okay"
    End If
    
    If ENTER_LARGE_IN_LOGGER = True Then
        If Len(AD$) > LimitClipSize Then
            VAR_INFO_LOGG_CLIP = "LARGE TEXT IN LOGGER - ACCEPTED IS CUSTOMER USER REQUESTED"
        End If
    End If
    
    If FindWindow(vbNullString, "BluetoothLogView") = GetForegroundWindow Then
        'If Len(AD$) > LimitClipSize = True Then
            VAR_INFO_LOGG_CLIP = "LARGE TEXT IN LOGGER - ACCEPTED IS BluetoothLogView"
        'End If
        ENTER_LARGE_IN_LOGGER = True
    End If
    
    If ENTER_LARGE_IN_LOGGER = False Then
        If Len(AD$) > LimitClipSize Then
            VAR_INFO_LOGG_CLIP = "BLOCKED LARGE TEXT IN LOGGER"
            AD$ = "": GGSize = True
            Exit Sub
        Else
            GGSize = False
        End If
        'CLEAR THE AD$ BECAUSE IF IT'S BIG WON'T MATTER TO
        'RELOAD AGAIN AND CHECK IF IT'S DUPE
    End If

    '----------------------------------------------------
    'RESET THIS IT WAS CALLED BY THE MENU PULL DOWN CLICK
    '----------------------------------------------------
    ENTER_LARGE_IN_LOGGER = False

End If






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
    If OO_2 = True Then XIO = True: PICTURE_IMG = True
    If OO_1 = True Then XIO = True

    If XIO = False Then
        VAR_INFO_LOGG_CLIP = "CLIP INFO --- " + ClipFormatDescription
        INFO_TO_ENTER_IN_LOGGER = True
    End If


    If VAR_INFO_LOGG_CLIP = "CLIPBOARD PICTURE" Then
        INFO_TO_ENTER_IN_LOGGER = True
        
    End If


    'this is correct when clipboard is a empty
    '-----------------------------------------------
    If INFO_TO_ENTER_IN_LOGGER = False Then
        Exit Sub
    End If
    '-----------------------------------------------

End If






'-------------------------------
'NORMAL TEXT ENTRY -- TRUE FALSE
'-----------------
If OO_1 = True Then

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


'If ClipFormatDescription = "Text And Len(0) Empty Clipboard" Then Stop


'wndclass_desked_gsk


'--------
'GOING IN
'--------
'--------------------------------
'GET SOME VALUES - REQUIRED
'--------------------------------

On Error GoTo 0
ASH$ = GetActiveWindowTitle(True)

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

If ASH2$ <> "" And ASH2$ <> ASH$ Then
    RR$ = RR$ + "Form Parent = " + VAR_VBCRLF + ASH2$ + VAR_VBCRLF
    RR$ = RR$ + "---------------------" + VAR_VBCRLF
End If

If ASH2$ <> "" And ASH2$ <> ASH$ Then
    RR$ = RR$ + "Form Parent = " + VAR_VBCRLF + ASH2$ + VAR_VBCRLF
    RR$ = RR$ + "---------------------" + VAR_VBCRLF
End If
    
RR$ = RR$ + "Form FindWindow ---" + VAR_VBCRLF + ASH$ + VAR_VBCRLF

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


If OO_1 = True Then
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

If VB_ICON_CHANGED = True Then
    Exit Sub
End If


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
    'DoEvents
    '-----------------------------------------------------
    MSGBOX2 = "ERROR Not Adding to CLip Board TEXT1.TEXT RTB Textbox" + vbCrLf + Err.Description
    '-----------------------------------------------------
    frm_MSGBOX.Timer1.Enabled = True
    '-----------------------------------------------------
    Exit Sub
    '-----------------------------------------------------
End If

On Error GoTo 0


' TIMER1 _ WHERE NEW CLIPBOARD COMES ON

If MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = True Then
    If Text1.SelLength = 0 Then
        'TEXTBOX1_SELECTION_CHANGE = Text1.SelStart
        
        Text1.Enabled = False
        Text1.SelStart = 1
        Text1.Enabled = True
        'Text1.SelStart = TEXTBOX1_SELECTION_CHANGE
        Text1.SelStart = Len(Text1)
    
    End If
End If

If MNU_TOGGLE_NOT_SCROLL_TO_BOTTOM.Checked = False Then
    
    'MAIN TEXT BOX IS TEXT_BOX NOT RTB
    'THINK
    'RTB IS CALLED TEXT1
    
    
    ' TAKE OUT WHY -- DON'T LOOK LIKE WANTER NOTHNG SCROLL BOTTOM WHEN THIS OPTION SET
    ' NOT EVEN NEW CLIPBOARD
    ' 2019 06 09 JUNE
    ' --------------------------------------------------------------------------------
    ' Text1.SelStart = Len(Text1.Text)
    
    'WT-HECK -- IN ANOTHER MODE THE TEXT SCROLL WONT UPDATE STAY AT LAST
    'If Form1.Width <> 2400 Then
    '    Text1.SelStart = OFH
    'End If
    
    '----------------------------------------------------
    'DONT MOVE CURSOR IN TEXBOX IF TIMER SET AT NEW PASTE
    '----------------------------------------------------
    ' TAKE OUT WHY -- DON'T LOOK LIKE WANTER NOTHNG SCROLL BOTTOM WHEN THIS OPTION SET
    ' NOT EVEN NEW CLIPBOARD
    ' 2019 06 09 JUNE
    ' --------------------------------------------------------------------------------
'    If XXMOUSEMOVE < Now Then
'        Text1.SelStart = Len(Text1.Text)
'    End If

End If

On Error GoTo ErrorTrap

'---------------------------
'OUTPUT
'---------------------------
FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\#Count.Txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Output As #FR1
    Print #FR1, CountR
Close #FR1

'---------------------------
'APPEND
'---------------------------
FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\Day-Data\ClipBoard-" + Format$(start, "YYYY-MM-DD") + ".Txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, RR$;
Close #FR1

FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\#ClipBoard.Txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, RR$;
Close #FR1

FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, RR$;
Close #FR1

FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Total--TRIM.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, RR$;
Close #FR1
    
FILENAME_IN_USE_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1 + "\00_ClipBoard_Tot_Strip.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, AD$
Close #FR1

FILENAME_IN_USE_CHECK = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\00_ClipBoard_Total--TRIM-" + GetComputerName + "-" + GetUserName + ".txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, RR$;
Close #FR1


'AD$ = GG$
AD_DATE = Now


'-----------------------------
Exit Sub
'-----------------------------

'-----------------------------
ErrorTrap:
    'SLEEP 10
    'DoEvents
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

Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim L As Long
   Dim S As String
   L = GetWindowTextLength(hWnd)
   S = Space(L + 1)
   GetWindowText hWnd, S, L + 1
   GetWindowTitle = Left$(S, L)
End Function

Sub SET_FOLDER_CLIPBOARD_LOGGER()

    On Error Resume Next
    
    PATH_TO_CLIPBOARD_APP_DATA_1 = App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName
    PATH_TO_CLIPBOARD_APP_DATA_2 = App.Path + "\# DATA\0 LAST CLIPBOARD"
    
'    Mnu_Open_Logg.Caption = "Edit This Logg - " + PATH_TO_CLIPBOARD_APP_DATA_1 + "\Day-Data\ClipBoard-" + Format$(start, "YYYY-MM-DD") + ".Txt"

    If Not DirExist(PATH_TO_CLIPBOARD_APP_DATA_1) Then
        Err.Clear
        MkDirNested PATH_TO_CLIPBOARD_APP_DATA_1
        If Err.Number <> 0 Then
            MSGBOX2 = ""
            MSGBOX2 = MSGBOX2 + "ERROR Clipboard Logger" + vbCrLf
            MSGBOX2 = MSGBOX2 + "UNABLE TO MAKE FOLDER" + vbCrLf
            MSGBOX2 = MSGBOX2 + PATH_CHECK + "ERR.DESCRIPTION=" + Err.Description
            frm_MSGBOX.Timer1.Enabled = True
        End If
    End If
    If Not DirExist(PATH_TO_CLIPBOARD_APP_DATA_2) Then
        Err.Clear
        MkDirNested PATH_TO_CLIPBOARD_APP_DATA_2
        If Err.Number <> 0 Then
            MSGBOX2 = ""
            MSGBOX2 = MSGBOX2 + "ERROR Clipboard Logger" + vbCrLf
            MSGBOX2 = MSGBOX2 + "UNABLE TO MAKE FOLDER" + vbCrLf
            MSGBOX2 = MSGBOX2 + PATH_CHECK + "ERR.DESCRIPTION=" + Err.Description
            frm_MSGBOX.Timer1.Enabled = True
        End If
    End If
    
    ART2$ = "D:\0 00 Art Loggers\# APP AND SCREEN\" + GetComputerName + "\"
    ART4$ = "D:\0 00 Art Loggers\# APP AND SCREEN AUTO\" + GetComputerName + "_AUTO\"
    
    FF$ = Format$(Now, "YYYY-MM-MMM")
    FF_22 = "HOT_KEY_FORM_SHOT\HOT_KEY_FORM_SHOT_" + FF$ + "\"
    
    FF_24 = "AUTO_SCREEN__SHOT\" 'AUTO_SCREEN__SHOT_" ' + FF$ + "\"
    FF_28 = "AUTO_FORM____SHOT\" 'AUTO_FORM____SHOT_" ' + FF$ + "\"
    
    PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY = ART2$ + FF_22
    
    PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO_SCR__ = ART4$ + FF_24
    PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO_FORM_ = ART4$ + FF_28
    
    PATH_CHECK = PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY
    
    If Not DirExist(PATH_CHECK) Then
        Err.Clear
        MkDirNested PATH_CHECK
        If Err.Number <> 0 Then
            MSGBOX2 = ""
            MSGBOX2 = MSGBOX2 + "ERROR Clipboard Text Data Not Getting Saved" + vbCrLf
            MSGBOX2 = MSGBOX2 + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf
            MSGBOX2 = MSGBOX2 + PATH_CHECK + "ERR.DESCRIPTION=" + Err.Description
            frm_MSGBOX.Timer1.Enabled = True
        End If
    End If
    PATH_CHECK = PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO_SCR__
    If (Not DirExist(PATH_CHECK)) Then
        Err.Clear
        MkDirNested PATH_CHECK
        If Err.Number <> 0 Then
            MSGBOX2 = ""
            MSGBOX2 = MSGBOX2 + "ERROR Clipboard Text Data Not Getting Saved" + vbCrLf
            MSGBOX2 = MSGBOX2 + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf
            MSGBOX2 = MSGBOX2 + PATH_CHECK + "ERR.DESCRIPTION=" + Err.Description
            frm_MSGBOX.Timer1.Enabled = True
        End If
    End If
    PATH_CHECK = PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO_FORM_
    If (Not DirExist(PATH_CHECK)) Then
        Err.Clear
        MkDirNested PATH_CHECK
        If Err.Number <> 0 Then
            MSGBOX2 = ""
            MSGBOX2 = MSGBOX2 + "ERROR Clipboard Text Data Not Getting Saved" + vbCrLf
            MSGBOX2 = MSGBOX2 + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf
            MSGBOX2 = MSGBOX2 + PATH_CHECK + "ERR.DESCRIPTION=" + Err.Description
            frm_MSGBOX.Timer1.Enabled = True
        End If
    End If
    PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP = App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\Hot Key Capture Text"
    PATH_CHECK = PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP
    If (Not DirExist(PATH_CHECK)) Then
        Err.Clear
        MkDirNested PATH_CHECK
        If Err.Number <> 0 Then
            MSGBOX2 = ""
            MSGBOX2 = MSGBOX2 + "ERROR Clipboard Text Data Not Getting Saved" + vbCrLf
            MSGBOX2 = MSGBOX2 + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf
            MSGBOX2 = MSGBOX2 + PATH_CHECK + "ERR.DESCRIPTION=" + Err.Description
            frm_MSGBOX.Timer1.Enabled = True
        End If
    End If
    PATH_TO_CLIPBOARD_TEXT_INFO_APP_DAY_DATA = App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\Day-Data"
    PATH_CHECK = PATH_TO_CLIPBOARD_TEXT_INFO_APP_DAY_DATA
    If (Not DirExist(PATH_CHECK)) Then
        Err.Clear
        MkDirNested PATH_CHECK
        If Err.Number <> 0 Then
            MSGBOX2 = ""
            MSGBOX2 = MSGBOX2 + "ERROR Clipboard Text Data Not Getting Saved" + vbCrLf
            MSGBOX2 = MSGBOX2 + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf
            MSGBOX2 = MSGBOX2 + PATH_CHECK + "ERR.DESCRIPTION=" + Err.Description
            frm_MSGBOX.Timer1.Enabled = True
        End If
    End If
    PATH_TO_CLIPBOARD_APP_DATA_1 = App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName
    PATH_CHECK = PATH_TO_CLIPBOARD_APP_DATA_1
    If (Not DirExist(PATH_CHECK)) Then
        Err.Clear
        MkDirNested PATH_CHECK
        If Err.Number <> 0 Then
            MSGBOX2 = ""
            MSGBOX2 = MSGBOX2 + "ERROR Clipboard Text Data Not Getting Saved" + vbCrLf
            MSGBOX2 = MSGBOX2 + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf
            MSGBOX2 = MSGBOX2 + PATH_CHECK + "ERR.DESCRIPTION=" + Err.Description
            frm_MSGBOX.Timer1.Enabled = True
        End If
    End If
    PATH_CHECK = App.Path + "\# DATA\# VBIcon"
    If (Not DirExist(PATH_CHECK)) Then
        Err.Clear
        MkDirNested PATH_CHECK
        If Err.Number <> 0 Then
            MSGBOX2 = ""
            MSGBOX2 = MSGBOX2 + "ERROR Clipboard Text Data Not Getting Saved" + vbCrLf
            MSGBOX2 = MSGBOX2 + "Can't Create Folder Not a D Drive Maybe OR ACCESS" + vbCrLf
            MSGBOX2 = MSGBOX2 + PATH_CHECK + "ERR.DESCRIPTION=" + Err.Description
            frm_MSGBOX.Timer1.Enabled = True
        End If
    End If

End Sub


Sub PictureClip()


Dim IP, TimeSet2
On Error Resume Next

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
    

FILENAME1 = PATH_TO_CLIPBOARD_APP_DATA_1 + "\Hot Key Capture Text\COMPARE DUPE.JPG"
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
'"\Hot Key Capture Text"
'PATH_TO_CLIPBOARD_APP_DATA_1 = App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName
FileInUseDelay PATH_TO_CLIPBOARD_APP_DATA_1 + "\Hot Key Capture Text\HotKey-0-Shot % Pic Data.txt"
FR1 = FreeFile
Open PATH_TO_CLIPBOARD_APP_DATA_1 + "\Hot Key Capture Text\HotKey-0-Shot % Pic Data.txt" For Append As #FR1
    Print #FR1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p")
Close #FR1

'If Wo / (Ri + Wo) * 100 >= 1 Then
'TS1 = PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY + "\HotKey-Capture- " + TimeSet2 + "-" + QQ3$ + ".jpg"

TimeSet2 = Format$(Now, "YYYY-MM-DD HH-MM-SS DDD")
TS1 = PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY + "HotKey-Capture- " + TimeSet2 + ".jpg"
SavePicture Picture3.Picture, TS1
LAST_CLIPBOARD_FILE = TS1






' ----------------------------
' PATH_TO_CLIPBOARD_APP_DATA_1
' 01 -- APPEND
' THIS IMAGE CLIPBOARD
' ----------------------------
FILE_NAME_USER = PATH_TO_CLIPBOARD_APP_DATA_1 + "\Hot Key Capture Text\HotKey-0-Shot % Pic Data Location.txt"
If DirExist_With_FileName(FILE_NAME_USER) Then
    FILE_NAME_USER_HOTKEY_SHOT_PIC_DATA_LOCATION = FILE_NAME_USER
    FileInUseDelay FILE_NAME_USER
    FR1 = FreeFile
    Open PATH_TO_CLIPBOARD_APP_DATA_1 + "\Hot Key Capture Text\HotKey-0-Shot % Pic Data Location.txt" For Append As #FR1
        Print #FR1, LAST_CLIPBOARD_FILE
    Close #FR1
End If

' ----------------------------
' PATH_TO_CLIPBOARD_APP_DATA_1 =
' 02 -- LOCATION
' 00 LAST CLIPBOARD
' ----------------------------
FILE_NAME_USER = PATH_TO_CLIPBOARD_APP_DATA_1 + "\Hot Key Capture Text\HotKey-0-Shot % Pic Data Location Current.txt"
If DirExist_With_FileName(FILE_NAME_USER) Then
    FILE_NAME_USER_HOTKEY_SHOT_PIC_DATA_LOCATION_CURRENT = FILE_NAME_USER
    FileInUseDelay FILE_NAME_USER
    FR1 = FreeFile
    Open FILE_NAME_USER For Append As #FR1
        Print #FR1, LAST_CLIPBOARD_FILE
    Close #FR1
End If

' ----------------------------
' PATH_TO_CLIPBOARD_APP_DATA_2
' 03 -- LOCATION GLOBAL
' 00 LAST CLIPBOARD
' ----------------------------

FILE_NAME_USER = PATH_TO_CLIPBOARD_APP_DATA_2 + "\Hot Key Capture Text\HotKey-0-Shot % Pic Data Location Current Network Common.txt"
If DirExist_With_FileName(FILE_NAME_USER) Then
    FILE_NAME_USER_HOTKEY_SHOT_PIC_DATA_NETWORK_COMMON = FILE_NAME_USER
    FileInUseDelay FILE_NAME_USER
    FR1 = FreeFile
    Open FILE_NAME_USER For Append As #FR1
        Print #FR1, LAST_CLIPBOARD_FILE
    Close #FR1
End If


Set F = FSO.GetFile(TS1)

Call Menu_Clipboard_size(F.Size)
Call GetFormat_And_Display

PATH_TO_CLIPBOARD_AND_FILENAME = PATH_TO_CLIPBOARD_APP_DATA_1 + "\Hot Key Capture Text\HotKey-Capture- " + TimeSet2 + "-ClipText" + ".txt"

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


Dim MyStr As String, theHwnd As Long
'Create a buffer
theHwnd = GetForegroundWindow()
MyStr = String(GetWindowTextLength(theHwnd) + 1, Chr$(0))
'Get the window's text
GetWindowText theHwnd, MyStr, Len(MyStr)

 '- Google Chrome
If InStr(MyStr, "FireShot Capture") Then

    Shell "Explorer.exe /SELECT, """ + LAST_CLIPBOARD_FILE + """", vbMaximizedFocus

End If



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
        frm_MSGBOX.Timer1.Enabled = True
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


Function GetWindowClass(ByVal hWnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hWnd, sText, 255)
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
Open App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\ChkSettings.txt" For Input As #FR1
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
If FSO.FolderExists(App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName) = False Then
    MkDir App.Path + "\# DATA"
    MkDir App.Path + "\# DATA\" + GetComputerName
    MkDir App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName
End If

Open App.Path + "\# DATA\" + GetComputerName + "_" + GetUserName + "\ChkSettings.txt" For Output As #1

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
    If Mid(LCase(Control.Name), 1, 4) = "file" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 3) = "rtb" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "auto" Then ATD = 0
    If Mid(UCase(Control.Name), 1, 4) = "MNU_" Then ATD = 2
    If Mid(UCase(Control.Name), 1, 4) = "TXT_" Then ATD = 0
    If ATD = 1 Then
        'Control.NAME
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
    If Mid(LCase(Control.Name), 1, 4) = "file" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 3) = "rtb" Then ATD = 0
    If Mid(UCase(Control.Name), 1, 4) = "COMM" Then ATD = 0
    If Mid(UCase(Control.Name), 1, 4) = "TXT_" Then ATD = 0
    If ATD = 1 Then
        'control.name
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
    If InStr(UCase(Control.Name), "FILE") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "RTB") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "AUTO") > 0 Then NOT_FLAG = 1
    If InStr(UCase(Control.Name), "MNU_") > 0 Then NOT_FLAG = -1
    If InStr(UCase(Control.Name), "TXT_") > 0 Then NOT_FLAG = -1

    If NOT_FLAG <= 0 Then
        'Control.NAME
        If NOT_FLAG = -1 Then
        A1 = 0
        Else
        A1 = Control.Value
        B1 = Control.Checked
        End If
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
If FSO.FileExists(FDS) = True Then
    If FileInUse(FDS) = False Then
        
        Filespec1 = FDS
        Set F = FSO.GetFile((Filespec1))
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
        'If FSO.FileExists(FDS2) = False Then Stop
        
            If MNU_PLAY_SOUND_ON_LOAD.Checked = True Or MODIFY_SOUND_SELECTION = True Then
                If MMControl1.Filename = "" Then
                    Call RESET_SETUP_SOUND_FILE("NOTSOUND")
                End If
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

Timer_1_MINUTE.Interval = 1000

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
'Set F = FSO.getfile((Filespec1))
'ADATE1 = F.datelastmodified
'Dim DateXVar
ADATE1 = 0
    
On Error GoTo END_EXIT
FileSpec2 = App.Path + "\" + App.EXEName + ".EXE"
Set F = FSO.GetFile(FileSpec2)

ADATE1 = F.DateLastModified
APPEXEDATE = ADATE1
    
FILESPEC3 = App.Path + "\" + App.EXEName + ".EXE"
FILESPEC3 = Replace(FILESPEC3, ":\VB6\", ":\VB6-EXE\")
If FSO.FileExists(FILESPEC3) = False Then Exit Sub
Set F = FSO.GetFile(FILESPEC3)
ADATE_MIRROR_EXE_DATE = F.DateLastModified

'If ADATE_MIRROR_EXE_DATE <= ADATE1 Then
'    Exit Sub
'End If

Call Timer_EXE_DAY_AGE_TIMER
Call Timer_EXE_DAY_AGE_MIRROR_TIMER
    
    
'WORK HERE T0 AUTOUPDATE SCRIPT


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
        Set F = FSO.GetFile((Filespec1))
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
        Set F = FSO.GetFile((FileSpec2))
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
'            Set F = FSO.getfile((Filespec2))
'            ADATE2 = F.datelastmodified
'            APPEXEDATE = ADATE2
'        End If
'        On Error GoTo 0
        
    End If
    
    
    
    'READ IT THE DATE AGAIN
    FileSpec2 = App.Path + "\" + App.EXEName + ".EXE"
    Set F = FSO.GetFile(FileSpec2)
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
        'MNU_VB.Caption = "VB ME" + Str(Pro3) + " " + DATE_CODE
        
        'WE NEED TO SWAP OUR NEW EXE IN
        'LOOK FOR OUR EXE IN MIRROR SYNC FOLDER IS HERE FIRST
        
        FILESPEC3 = App.Path + "\" + App.EXEName + ".EXE"
        'MIRROR
        FILESPEC3 = Replace(FILESPEC3, ":\VB6\", ":\VB6-EXE'S\")
        Set F = FSO.GetFile(FILESPEC3)
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
'    Set F = FSO.getfile((Filespec1))
'    ADATE1 = F.datelastmodified
'    Filespec2 = App.Path + "\01 Woarble Tone 5 Mins 02.wav"
'    Set F = FSO.getfile((Filespec2))
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
    Dim PROX_M, PROX_S
    'Timer_APP_BEGIN_TIME.Interval = 1000
    PROX_D1 = DateDiff("d", ADATE_APP_BEGIN_DATE, Now)
    PROX_D = Format(DateDiff("d", ADATE_APP_BEGIN_DATE, Now), "0") + " Day"
    If PROX_D1 = 0 Then PROX_D = Replace(PROX_D, "0 Day", "")
    PROX_H = Format(DateDiff("h", ADATE_APP_BEGIN_DATE, Now), "00") + " Hour"
    If PROX_D1 = 0 Then PROX_H = Replace(PROX_H, "00 Hour", "")
    PROX_M = Format(DateDiff("n", ADATE_APP_BEGIN_DATE, Now), "00") + " Minute"
    
    'IF IN SECOND MODE - - JOINT SUB ROUTINE
    
    PROX_S1 = DateDiff("s", ADATE_APP_BEGIN_DATE, Now)
    If PROX_S1 < 60 Then
        PROX_S = Format(DateDiff("s", ADATE_APP_BEGIN_DATE, Now), "00") + " Second"
    End If
    
    '---------------------------------------
    'USED SOMEWEHRE ELSE NOT THIS SUBROUTINE
    'FIRST CHAR OF TIME ID
    '---------------------------------------
    TIMERVAR4 = Mid(TIMERVAR3, 1, 1)
    
    TT_VAR = "CLIPBOARD LOGGER BEGIN TIME = " + Format(ADATE_APP_BEGIN_DATE, "DD-MMM-YYYY hh:mm:ss") + " --" + Str(Pro3) + " " + TIMERVAR3 + " Age -- STARUP TIME SECONDS =" + Str(ADATE_APP_BEGIN_DATE2)
    If Mnu_APP_BEGIN_TIME_TIMER.Caption <> TT_VAR Then Mnu_APP_BEGIN_TIME_TIMER.Caption = TT_VAR
    
    PROX_M = Trim(PROX_M)
    PROX_S = Trim(PROX_S)
    PROX_M = Replace(PROX_M, "Minute", "Min")
    PROX_S = Replace(PROX_S, "Second", "Sec")
    TT_VAR_RESULT_MNU_RUN_TIME = "Run " + PROX_D + " " + PROX_H + " " + PROX_M + " " + PROX_S
    For r = 1 To 5
       TT_VAR_RESULT_MNU_RUN_TIME = Replace(TT_VAR_RESULT_MNU_RUN_TIME, "  ", " ")
    Next
'   MNU_RUN_TIME.Visible = False
'   If MNU_RUN_TIME.Caption <> TT_VAR Then MNU_RUN_TIME.Caption = TT_VAR
    Call SET_LABEL3

End Sub


Sub CALC_DO_THE_ADDING()
    
    On Error Resume Next
    Err.Clear

If Clipboard_GetFormat_vbCFText_OO_1 = True Then
    CLIPPER_CALC = Clipboard.GetText
    If Err.Number > 0 Then
        CLIPBOARD_RETURN_TIMER_ERROR_5 = 1
        Exit Sub
    End If
    
    On Error GoTo 0
    NUMERIC_VAR = CLIPPER_CALC
    NUMERIC_VAR = Replace(NUMERIC_VAR, "£", "")
    NUMERIC_VAR = Replace(NUMERIC_VAR, "$", "")
    NUMERIC_VAR = Replace(NUMERIC_VAR, "+", "")
    NUMERIC_VAR = Replace(NUMERIC_VAR, "-", "")
    NUMERIC_VAR = Replace(NUMERIC_VAR, "*", "")
    NUMERIC_VAR = Replace(NUMERIC_VAR, "/", "")
    NUMERIC_VAR = Replace(NUMERIC_VAR, "\", "")
    NUMERIC_VAR = Replace(NUMERIC_VAR, ".", "")
    NUMERIC_VAR = Replace(NUMERIC_VAR, " ", "")
    NUMERIC_VAR = Replace(NUMERIC_VAR, vbCr, "")
    NUMERIC_VAR = Replace(NUMERIC_VAR, vbLf, "")
    If InStr(NUMERIC_VAR, "=") Then
        NUMERIC_VAR = Mid(NUMERIC_VAR, 1, InStr(NUMERIC_VAR, "=") - 1)
    End If
    
    'If O2_CLIPPER_CALC <> CLIPPER_CALC And IsNumeric(NUMERIC_VAR) Then
    If IsNumeric(NUMERIC_VAR) Then
        
        O2_CLIPPER_CALC = CLIPPER_CALC
        
        CALC_STRINGER_S = "CALC CLIPPER = NONE"
        CLIPPER_CALC = Replace(CLIPPER_CALC, "£", "")
        CLIPPER_CALC = Replace(CLIPPER_CALC, "$", "")
        CLIPPER_CALC = Replace(CLIPPER_CALC, vbCrLf, "+")
        CLIPPER_CALC = Replace(CLIPPER_CALC, vbCr, "+")
        CLIPPER_CALC = Replace(CLIPPER_CALC, vbLf, "+")
        CLIPPER_CALC = Replace(CLIPPER_CALC, " ", "")
        If InStr(CLIPPER_CALC, "=") Then
            CLIPPER_CALC = Mid(CLIPPER_CALC, 1, InStr(CLIPPER_CALC, "=") - 1)
        End If
        Do
            CLIPPER_CALC = Replace(CLIPPER_CALC, "++", "+")
            If O_CLIPPER_CALC = CLIPPER_CALC Then Exit Do
            O_CLIPPER_CALC = CLIPPER_CALC
        Loop Until True = False
        Do
            CLIPPER_CALC = Replace(CLIPPER_CALC, "--", "-")
            If O_CLIPPER_CALC = CLIPPER_CALC Then Exit Do
            O_CLIPPER_CALC = CLIPPER_CALC
        Loop Until True = False
        Do
            CLIPPER_CALC = Replace(CLIPPER_CALC, "**", "*")
            If O_CLIPPER_CALC = CLIPPER_CALC Then Exit Do
            O_CLIPPER_CALC = CLIPPER_CALC
        Loop Until True = False
        Do
            CLIPPER_CALC = Replace(CLIPPER_CALC, "//", "/")
            If O_CLIPPER_CALC = CLIPPER_CALC Then Exit Do
            O_CLIPPER_CALC = CLIPPER_CALC
        Loop Until True = False
        Do
            CLIPPER_CALC = Replace(CLIPPER_CALC, "\\", "\")
            If O_CLIPPER_CALC = CLIPPER_CALC Then Exit Do
            O_CLIPPER_CALC = CLIPPER_CALC
        Loop Until True = False
        OPERATORS_CALC = "+-*/\"
        Do
            For R_1 = 1 To 5
                If Mid(CLIPPER_CALC, Len(CLIPPER_CALC), 1) = Mid(OPERATORS_CALC, R_1, 1) Then
                    CLIPPER_CALC = Mid(CLIPPER_CALC, 1, Len(CLIPPER_CALC) - 1)
                End If
            Next
            If O_CLIPPER_CALC = CLIPPER_CALC Then Exit Do
            O_CLIPPER_CALC = CLIPPER_CALC
        Loop Until True = False
        
        'CALC_ARRAY_VAR = Split(CLIPPER_CALC, "+")
        ReDim CALC_ARRAY_VAR(5000)
        X_VAR_1 = 1
        OPERATORS_CALC = "+-*/\"
        CALC_ARRAY_VAR_COUNT = 0
        Do
            X_VAR_T = Len(CLIPPER_CALC) + 10
            X_VAR_V = 0
            DO_ANOTHER = False
            For R_1 = 1 To 5
                X_VAR_V = InStr(X_VAR_1, CLIPPER_CALC, Mid(OPERATORS_CALC, R_1, 1))
                If X_VAR_V > 0 Then DO_ANOTHER = True
                If X_VAR_V < X_VAR_T And X_VAR_V > 0 Then X_VAR_T = X_VAR_V
            Next
        
            If DO_ANOTHER = True Then
                END_X_VAR = X_VAR_T
            Else
                END_X_VAR = Len(CLIPPER_CALC) + 1
            End If
                        
            CALC_ARRAY_VAR_COUNT = CALC_ARRAY_VAR_COUNT + 1
            X_VAR_M = X_VAR_1 - 1
            If X_VAR_M <= 0 Then X_VAR_M = 1
            CALC_ARRAY_VAR(CALC_ARRAY_VAR_COUNT) = Trim(Mid(CLIPPER_CALC, X_VAR_M, END_X_VAR - X_VAR_M))
            CALC_ARRAY_VAR(CALC_ARRAY_VAR_COUNT) = Replace(CALC_ARRAY_VAR(CALC_ARRAY_VAR_COUNT), " ", "")
            X_VAR_1 = X_VAR_T + 1
        
        Loop Until DO_ANOTHER = False
        ReDim Preserve CALC_ARRAY_VAR(CALC_ARRAY_VAR_COUNT)
        
        CALC_STRINGER_T = 0
        If UBound(CALC_ARRAY_VAR) > 0 Then
            For r = 1 To UBound(CALC_ARRAY_VAR)
                NUMERIC_VAR = CALC_ARRAY_VAR(r)
                NUMERIC_VAR = Replace(NUMERIC_VAR, "+", "")
                NUMERIC_VAR = Replace(NUMERIC_VAR, "-", "")
                NUMERIC_VAR = Replace(NUMERIC_VAR, "*", "")
                NUMERIC_VAR = Replace(NUMERIC_VAR, "/", "")
                NUMERIC_VAR = Replace(NUMERIC_VAR, "\", "")
                If IsNumeric(NUMERIC_VAR) = True And Val(NUMERIC_VAR) <> 0 Then
                    OPERATORS_CALC = "+-*/\"
                    X_VAR_V = Mid(CALC_ARRAY_VAR(r), 1, 1)
                    If IsNumeric(X_VAR_V) Then X_VAR_V = "+"
                    If X_VAR_V = "+" Then CALC_STRINGER_T = CALC_STRINGER_T + Val(NUMERIC_VAR)
                    If X_VAR_V = "-" Then CALC_STRINGER_T = CALC_STRINGER_T - Val(NUMERIC_VAR)
                    If X_VAR_V = "*" Then CALC_STRINGER_T = CALC_STRINGER_T * Val(NUMERIC_VAR)
                    If X_VAR_V = "/" Then CALC_STRINGER_T = CALC_STRINGER_T / Val(NUMERIC_VAR)
                End If
            Next
        End If
        
        CALC_VALUE = True
        For r = 1 To UBound(CALC_ARRAY_VAR)
            If IsNumeric(CALC_ARRAY_VAR(r)) = True Then r = r + 1
            If r > UBound(CALC_ARRAY_VAR) Then r = r - 1
            If IsNumeric(Mid(CALC_ARRAY_VAR(r), 2)) = False Then
                CALC_VALUE = False
            End If
        Next
        If UBound(CALC_ARRAY_VAR) <= 1 Then CALC_VALUE = False
        
        '-100-100+100+100-100
        If CALC_VALUE = True Then
            CALC_STRINGER_S = ""
            For r = 1 To UBound(CALC_ARRAY_VAR)
                
                CALC_STRINGER_S = CALC_STRINGER_S + Trim(CALC_ARRAY_VAR(r))
            Next
            
            CALC_STRINGER_EQUAL = Trim(Str(CALC_STRINGER_T))
            If InStr(CALC_STRINGER_EQUAL, ".") > 0 Then
                LEN_COUNTING = Mid(CALC_STRINGER_EQUAL, InStr(CALC_STRINGER_EQUAL, ".") + 1)
                If Len(LEN_COUNTING) < 2 Then
                    CALC_STRINGER_EQUAL = Format(CALC_STRINGER_T, "0.00")
                End If
            End If
            CALC_STRINGER_S = CALC_STRINGER_S + " = £" + CALC_STRINGER_EQUAL
            ' A LEADING MINUS IN ORGINAL STRING WOULD NOT BE OMITED SO REQUIRING REPLACE
        
            'OPERATORS_CALC = "+-*/\"
            'For R_1 = 1 To 5
            '    CALC_STRINGER_S = Replace(CALC_STRINGER_S, Mid(OPERATORS_CALC, R_1, 1), " " + Mid(OPERATORS_CALC, R_1, 1) + " ")
            'Next
        
        End If
        
        
        
        'Debug.Print Len(CALC_STRINGER_S)
        'Debug.Print Len(O_CALC_STRINGER_S)
        
        If CALC_VALUE = True Then
            CALC_SET_GO = False
            If FindWindowPart("E:\HARDWARE") > 0 Then CALC_SET_GO = True
            If CALC_MODE = True Then CALC_SET_GO = True
            If CALC_VALUE = True Then CALC_SET_GO = True
            If CALC_VALUE = False Then
                CALC_STRINGER_S = "CALC CLIPPER = NONE"
                CALC_SET_GO = False
            End If
            If O_CALC_STRINGER_S = CALC_STRINGER_S Then
                'CALC_SET_GO = False
            End If
            If CALC_NEW_INPUT_MNU = True Then
                CALC_NEW_INPUT_MNU = False
                CALC_SET_GO = True
            End If
        End If
    End If
End If

End Sub

Sub SET_LABEL3()

If SET_LABEL3_INFO_BAR_READY_TO_GO_VAR = False Then Exit Sub

If TT_VAR_API_RESET = "" Then Exit Sub
If TT_VAR_LAST_CLIP_TIME_Second = "" Then Exit Sub
If TT_VAR_RESULT_MNU_RUN_TIME = "" Then Exit Sub

If VAL_MINUTE_API > 0 Then
    TT_VAR_API_RESET = "API " + Format(VAL_MINUTE_API, "00") '+ Test_Min_Var
    'TT_VAR_API_RESET = Replace(TT_VAR_API_RESET, "s --", "")
End If

TT_VAR_VERSION = "Ver. " + Format(Now, "YYYY") + "_" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))

var_xx = TT_VAR_LAST_CLIP_TIME_Second
var_xx = Replace(var_xx, "Start-up In Buffer", "In Start-up")
var_xx = Replace(var_xx, "Last Clipper Text -- ", "")

VAR_SIZE_VB = "Size " + Str(CLIPBOARD_SIZE_VAR)

'-----------------------------------
API_LOAD = True
If FRMCLIPTEST01.CALC_ADDER_ENTRY = True Then
    Call CALC_DO_THE_ADDING
    FRMCLIPTEST01.CALC_ADDER_ENTRY = False
'Else
 '   Exit Sub
End If

'CALC_NEW_INPUT = True
If CALC_NEW_INPUT = True Then
    CALC_NEW_INPUT = False
    If CALC_NEW_INPUT_MNU = True Then
    
        If CALC_SET_GO = True Then
            On Error Resume Next
            Do
                TIMER_Clipboard_Clear_VAR = False
                TIMER_Clipboard_Clear.Enabled = True
                Do
                    DoEvents
                    If Err.Number > 0 Then Sleep 10
                Loop Until TIMER_Clipboard_Clear_VAR = True
                Sleep 10
                API_LOAD = True
                FRMCLIPTEST01.CALC_ADDER_ENTRY = False
                
                ENTER_TEXT_IN_LOGGER_FOREGROUND_OVERRIDE = True
                Clipboard_Set_Text_VAR = CALC_STRINGER_S
                
                Call PUT_TEXT_CLIPBOARD(Clipboard_Set_Text_VAR)
                API_LOAD = True
                FRMCLIPTEST01.CALC_ADDER_ENTRY = False
                
                If Err.Number > 0 Then Sleep 100
                DoEvents
            Loop Until Clipboard.GetText = CALC_STRINGER_S
            On Error GoTo 0
        End If
    End If
End If

O_CALC_STRINGER_S = CALC_STRINGER_S
O2_CLIPPER_CALC = CALC_STRINGER_S

STRING_VAR = " " + TT_VAR_VERSION + " __" + TT_VAR_API_RESET + " __ " + TT_VAR_RESULT_MNU_RUN_TIME + " __ Last Clipper -- " + VAR_SIZE_VB + " -- " + var_xx
If CALC_STRINGER_S <> "CALC CLIPPER = NONE" Then
    STRING_VAR = STRING_VAR + vbCrLf + "  Calc_Clipper = "
    STRING_VAR = STRING_VAR + Mid(CALC_STRINGER_S, InStr(CALC_STRINGER_S, "=") + 2)
End If

If Label3.Caption <> STRING_VAR Then Label3.Caption = STRING_VAR
' Label3.Width
' Txt_Search.Left
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
            'MNU_VB.Caption = "VB ME -- NULL"
            
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

        'MNU_VB.Caption = "VB_ME FILE_EXE_CREATION" + Str(Pro3) + " " + DATE_CODE

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
        If IK2 = True Then A_TimeIdle = Now

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
    
    A_TimeIdle = Now
    
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
    
    
    If MNU_IDLE_ACTIVE.Caption <> "IDLE STATE = ACTIVE" Then
        MNU_IDLE_ACTIVE.Caption = "IDLE STATE = ACTIVE"
    End If
    
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

    'Call Timer_SCREEN_SHOT_AUTO_Timer
'------------------------------
    
    Timer_MOUSE.Enabled = False
    MouseClicks = 0
    Mousey = 0
    
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
        
        Timer_MOUSE.Enabled = False
        Timer_MOUSE.Interval = 5000
        Timer_MOUSE.Enabled = True
        KeyboardDetectPress = True
        A_TimeIdle = Now

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

FindCursor Nx, Ny

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
    Xmp5 = Nx: Ymp5 = Ny
End If


'If Quake = 0 Then
    If (Nx <> Xmp5 Or Ny <> Ymp5) Then 'And (Nx <> ScreenWidth And Ny <> ScreenHeight) Then
        'Call WinonPoint
        Mousey = Mousey + Sqr(Abs(Xmp5 - Nx) ^ 2 + Abs(Ymp5 - Ny) ^ 2)
        MouseClicks = Sqr(Abs(Xmp5 - Nx) ^ 2 + Abs(Ymp5 - Ny) ^ 2)
                    
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
            
            Timer_MOUSE.Enabled = False
            Timer_MOUSE.Interval = 5000
            Timer_MOUSE.Enabled = True
            
            'THIS IS SOURCE OF MAKE MENU BAR SCREEN FLICKER
            MouseDetectMove = True
        
        End If
        
        Xmp5 = Nx: Ymp5 = Ny
    
    End If
'End If

End Function


Sub Menu_Clipboard_size(VarSize As Double)
    
    Dim TTX, TTG

    CLIPBOARD_SIZE_VAR = VarSize

    If Clipboard_GetFormat_vbCFBitmap_OO_2 = True Then SIZE_IMAGE_NOT_MATTER = True
    
    'WORK TO DO HERE MORE THAN TEXT AN IMAGE
    
    If Clipboard_GetFormat_vbCFText_OO_1 = True Then TTX = "Text" Else TTX = "Image"
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


Sub DELETE_SOME_BACKLOG_AUTO_IMAGE_FOLDER()
    
    Dim AR_DIR(4)
    Dim AT_LEAST_ITEM_ARRAY(4)
    Dim TIME_SPAN_GIVER_ARRAY(4)
    ReDim Preserve AR_DAY_NOW(4)
    ' ------------------------------------------------------
    AR_DIR(1) = FOLDERNAME_AUTO_SWAP
    AR_DIR(2) = FOLDERNAME_AUTO_MOUSE_ACTIVE
    AR_DIR(3) = FOLDERNAME_AUTO_MINUTE
    AR_DIR(4) = FOLDERNAME_AUTO_MINUTE_LONG
    ' ------------------------------------------------------
    AT_LEAST_ITEM_ARRAY(1) = 400
    AT_LEAST_ITEM_ARRAY(2) = 400
    AT_LEAST_ITEM_ARRAY(3) = 400
    AT_LEAST_ITEM_ARRAY(4) = 400
    ' ------------------------------------------------------
    TIME_SPAN_GIVER_ARRAY(1) = 400 'DAY
    TIME_SPAN_GIVER_ARRAY(2) = 400 'DAY
    TIME_SPAN_GIVER_ARRAY(3) = 400 'DAY
    TIME_SPAN_GIVER_ARRAY(4) = TimeSerial(1, 0, 0)    ' A HOUR
    TIME_SPAN_GIVER_ARRAY(4) = 400 'DAY
    AR_DIR(1) = FOLDERNAME_AUTO_SWAP
    
    For NEXT_LOOP = 1 To 4
        If AR_DIR(NEXT_LOOP) <> "" Then
            If Day(Now) <> AR_DAY_NOW(NEXT_LOOP) Then
                AR_DAY_NOW(NEXT_LOOP) = Day(Now)
    
                ' ------------------------------------------------------
                AT_LEAST_ITEM = AT_LEAST_ITEM_ARRAY(NEXT_LOOP)
                TIME_SPAN_GIVER = Now - TIME_SPAN_GIVER_ARRAY(NEXT_LOOP)
                ' ------------------------------------------------------
                ScanPath.ListView1.ListItems.Clear
                ScanPath.TxtPath = AR_DIR(NEXT_LOOP)
                Call ScanPath.cmdScan_Click
                
                If ScanPath.ListView1.ListItems.Count > 0 Then
                    ' ---------------------------------------------------------
                    ' TRIM THE ARCHIVE
                    ' AT LEAST 10 NOT MORE THAN 40 DAY
                    ' GET THE DIR FILE INFO INTO A LISTBOX
                    ' ------------------------------------
                    For WE = ScanPath.ListView1.ListItems.Count To 1 Step -1
                        A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
                        B1 = ScanPath.ListView1.ListItems.Item(WE)
                        T_FILE = A1 + B1
                    
                        ' START WITH NEW -- TOP DOWN -- FAR NAUGHTER
                        ' GOT TO GO BACK WAY WHEN REMOVE SAME TIME
                        ' REMOVE SOME AND LEAVE 10 THERE ONLY IF DATE RIGHT 40 DAY
                        ' AT LEAST 10 NOT MORE THAN 40 DAY
                        
                        If ScanPath.ListView1.ListItems.Count <= AT_LEAST_ITEM Then Exit For
                        
                        Set F = FSO.GetFile(T_FILE)
                        ADATE1 = F.DateLastModified
                        
                        CURRENT_YEAR_START = DateSerial(Year(Now), 1, 1)
                        
                        If ADATE1 < TIME_SPAN_GIVER Then
                        If ADATE1 > CURRENT_YEAR_START Then
                            Debug.Print B1
                            If FSO.FileExists(T_FILE) = True Then
                                On Error Resume Next
                                FSO.DeleteFile T_FILE
                                On Error GoTo 0
                                ScanPath.ListView1.ListItems.Remove (WE)
                            End If
                        End If
                        End If
                    Next
                    ' ------------------
                    ' GOT 10 COUNTER
                    ' 10 OR MORE
                    ' NOT 40 DAY FOR LONG TIME -- NICE CODIN
                    ' DEBUG TEST ON ONE HOUR
                    ' 1 ONE -- 1 3 -- 1 TO 3 -- 1 2 3 -- 01 HOUR -- 001 HOUR --
                    ' ------------------
                    ' ------------------
                End If
            End If
        End If
    Next
    
End Sub


Private Sub Timer_SCREEN_SHOT_AUTO_Timer()

Timer_SCREEN_SHOT_AUTO.Interval = 1000


Call DELETE_SOME_BACKLOG_AUTO_IMAGE_FOLDER

'IDLE KEY SCREEN SHOT

Dim X22 As Long

YES_VAR = False
'If GetComputerName = "ASUS-BIGGER" Then YES_VAR = True
If GetComputerName = "1-ASUS-X5DIJ" Then YES_VAR = True
If GetComputerName = "4-ASUS-GL522VW" Then YES_VAR = True
If GetComputerName = "5-ASUS-P2520LA" Then YES_VAR = True
If GetComputerName = "7-ASUS-GL522VW" Then YES_VAR = True
If GetComputerName = "8-MSI-GP62M-7RD" Then YES_VAR = True


YES_VAR = False

'TestLowTimer = True
'If TestLowTimer = True Then YES_VAR = False
If YES_VAR = False Then Timer_SCREEN_SHOT_AUTO.Enabled = False: Exit Sub

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
    FF2$ = "AUTO_FORM____SHOT_" + Format$(Now, "YYYY-MM-DD-DDD") + "_FORM_SWAP\"
    FOLDERNAME_AUTO = PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO_FORM_ + FF2$
    FOLDERNAME_AUTO_SWAP = FOLDERNAME_AUTO
    'Debug.Print FOLDERNAME_AUTO
    If FSO.FolderExists(FOLDERNAME_AUTO) = False Then
        'If IsIDE = True Then Stop
        I = CreateFolderTree(FOLDERNAME_AUTO)
        If FSO.FolderExists(FOLDERNAME_AUTO) = True Then
            I = True
        End If
        If I = False Then
            MSGBOX2 = "ERROR PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME_AUTO
            On Error Resume Next
            frm_MSGBOX.Timer1.Enabled = True
            On Error GoTo 0
            Exit Sub
        End If
    End If
    
    FF_FORM1 = FOLDERNAME_AUTO
    FF_COUNT_FORM_CAPTURE1 = FF_COUNT_FORM_CAPTURE1 + 1
    If FSO.FolderExists(FOLDERNAME_AUTO) = False Then
        'If IsIDE = True Then Stop
        I = CreateFolderTree(FOLDERNAME_AUTO)
        If FSO.FolderExists(FOLDERNAME_AUTO) = True Then
            I = True
        End If
        If I = False Then
            MSGBOX2 = "ERROR PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME_AUTO
            On Error Resume Next  ' WHEN D DRIVE IS DOWN NOT LABELED
            frm_MSGBOX.Timer1.Enabled = True
            On Error GoTo 0
            Exit Sub
        End If
    End If
End If

'PRIORITY 2 MOUSE / KEY RESUME FROM IDLE OR ENTER ACTIVE program
If OverRideOnceExtra = True And YES_VAR = False Then
    DO_FORM = True
    OverRideOnceExtra = False
    FF2$ = "AUTO_FORM____SHOT_" + Format$(Now, "YYYY-MM-DD-DDD") + "_MOUSE_KEY_ACTIVE_IDLE\"
    FOLDERNAME_AUTO = PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO_FORM_ + FF2$
    FOLDERNAME_AUTO_MOUSE_ACTIVE = FOLDERNAME_AUTO

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
        FF2$ = "AUTO_FORM____SHOT_" + Format$(Now, "YYYY-MM-DD-DDD") + "_MINUTE\"
        FOLDERNAME_AUTO = PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO_FORM_ + FF2$
        FOLDERNAME_AUTO_MINUTE = FOLDERNAME_AUTO
        FF_FORM3 = FOLDERNAME_AUTO
        FF_COUNT_FORM_CAPTURE3 = FF_COUNT_FORM_CAPTURE3 + 1
        If FSO.FolderExists(FOLDERNAME_AUTO) = False Then
            I = CreateFolderTree(FOLDERNAME_AUTO)
            If FSO.FolderExists(FOLDERNAME_AUTO) = True Then
                I = True
            End If
            If I = False Then
                MSGBOX2 = "ERROR PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME_AUTO
                On Error Resume Next
                frm_MSGBOX.Timer1.Enabled = True
                On Error GoTo 0
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
        FF2$ = "AUTO_FORM____SHOT_" + Format$(Now, "YYYY-MM-DD-DDD") + "_MINUTE_LONG\"
        FOLDERNAME_AUTO = PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO_FORM_ + FF2$
        FOLDERNAME_AUTO_MINUTE_LONG = FOLDERNAME_AUTO
       'LEN(FOLDERNAME_AUTO)
        FF_FORM4 = FOLDERNAME_AUTO
        FF_COUNT_FORM_CAPTURE4 = FF_COUNT_FORM_CAPTURE4 + 1
        If FSO.FolderExists(FOLDERNAME_AUTO) = False Then
            I = CreateFolderTree(FOLDERNAME_AUTO)
            If FSO.FolderExists(FOLDERNAME_AUTO) = True Then
                I = True
            End If
            If I = False Then
                MSGBOX2 = "ERROR PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME_AUTO
                On Error Resume Next
                frm_MSGBOX.Timer1.Enabled = True
                On Error GoTo 0
                Exit Sub
            End If
        End If
    End If
End If
'NONE THESE HAPPEN SAME TIME

If YES_VAR = False Then Exit Sub

    If FSO.FolderExists(FOLDERNAME_AUTO) = False Then
        I = CreateFolderTree(FOLDERNAME_AUTO)
        If FSO.FolderExists(FOLDERNAME_AUTO) = True Then
            I = True
        End If
        If I = False Then
            MSGBOX2 = "ERROR PROBLEM MAKE CAPTURE FOLDER" + vbCrLf + FOLDERNAME_AUTO
            On Error Resume Next
            frm_MSGBOX.Timer1.Enabled = True
            On Error GoTo 0
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
    MSGBOX2 = "ERR.Description =" + Err.Description + vbCrLf + "ERR.Number =" + Str(Err.Number) + vbCrLf + "ERROR AT COMMAND :-" + vbCrLf + "SavePicture SCREEN_CAP.Picture, TS" + vbCrLf + "PICTURE IMAGE DID NOT SAVE" + vbCrLf + TS + vbCrLf + "Timer_SCREEN_SHOT_AUTO -- WILL BE ENABLED=FALSE -- AFTER ERROR"
    On Error Resume Next
    frm_MSGBOX.Timer1.Enabled = True
    Timer_SCREEN_SHOT_AUTO.Enabled = False
    On Error GoTo 0
    
    ERR_STAT = "ERROR -- "
Else
    ERR_STAT = ""
End If
    
'    FOLDERNAME_AUTO = TS
'    FOLDERNAME_AUTO_SWAP = FOLDERNAME_AUTO
'    FOLDERNAME_AUTO_MOUSE_ACTIVE = FOLDERNAME_AUTO
'    FOLDERNAME_AUTO_MINUTE = FOLDERNAME_AUTO
'    FOLDERNAME_AUTO_MINUTE_LONG = FOLDERNAME_AUTO
    
    MNU_DISPLAY = FOLDERNAME_AUTO_SWAP
        
    'Debug.Print TS
    
    ' MNU_TXT = Replace(TS, "D:\0 00 ART LOGGERS\", "__\")
    MNU_TXT = Replace(MNU_DISPLAY, "D:\0 00 ART LOGGERS\", "__\")
    ' REMOVE BIT HERE NOT SHOW FULL THING
    
    'MNU_TXT = Replace(MNU_TXT, GetComputerName, "...")
    'Debug.Print MNU_TXT
    
    ' Mnu_Explorer_Screen_Capture.Caption = ERR_STAT + "EXPLORE AUTO SCREEN IMAGE -- " + MNU_TXT
    ' Mnu_IVIEW_Screen_Capture.Caption = ERR_STAT + "IVIEW AUTO SCREEN IMAGE -- " + MNU_TXT
    
    Mnu_Explorer_Form_Capture.Caption = ERR_STAT + "EXPLORE AUTO SCREEN -- " + MNU_TXT
    Mnu_IVIEW_Form_Capture.Caption = ERR_STAT + "IVIEW AUTO FORM IMAGE -- " + MNU_TXT
    
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
'If FSO.FolderExists(FOLDERNAME_AUTO) = False Then
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



Sub CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY(RUN_CALL)

'1..MNU_CLIPBOARD_EXPLORER_FILE_FOLDER_Click
'2..MNU_REG_JUMP_Click
'3..MNU_URL_Browser_Click
'4..MNU_CPC_Click

    Sleep 500

'   Exit Sub
    On Error Resume Next
    Do
        Err.Clear
        TEXT_PATH = Clipboard.GetText
        If Err.Number > 0 Then Sleep 100
    URL_PATH = "https://www.ebay.co.uk/myb/PurchaseHistory#PurchaseHistoryOrdersContainer?ipp=50&Period=2&Filter=1&radioChk=1&GotoPage=1"
    URL_PATH = "https://www.ebay.co.uk/myb/PurchaseHistory#PurchaseHistoryOrdersContainer?ipp=50&Period=2&Filter=1&radioChk=1&GotoPage=1"
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    TOP_ARRAY_IAR = 12
    Dim IAR()
    ReDim IAR(TOP_ARRAY_IAR)
    IAR(1) = ":\"
    IAR(2) = "\\"
    IAR(3) = "HKEY_" 'MAIN
    IAR(4) = "HCR\"  'SHORTENED
    IAR(5) = "HCU\"
    IAR(6) = "HKCU\"
    IAR(7) = "HKU\"
    IAR(8) = "HKLM\"
    IAR(9) = "HU\"
    IAR(10) = "HCC\"
    
    IAR(11) = "HTTP"
    IAR(12) = "WWW"
    
    'USE CLIPBOARD OR IF NOT CHECK STAORED CLIPPER AD$
    'CLIPBOARD
    For r = 1 To TOP_ARRAY_IAR
        If InStr(UCase(TEXT_PATH), IAR(r)) > 0 Then RESULT_OPENER = True
    Next
    'AD$
    If RESULT_OPENER = False Then
        For r = 1 To TOP_ARRAY_IAR
            If InStr(UCase(AD$), IAR(r)) > 0 Then RESULT_OPENER = True: TEXT_PATH = AD$
        Next
    End If
    
    TEXT_PATH_ORI = TEXT_PATH
    TEXT_PATH_L = LCase(TEXT_PATH)
    TEXT_PATH_U = UCase(TEXT_PATH)
    
    If RUN_CALL = True Then I = True
    I = RUN_CALL
    XPOS = 0
    
    'FOLDER_FILE_PATH
    'EXPLORER FOLDER FILE
    If XPOS = 0 Then If InStrRev(TEXT_PATH, ":\") > 0 Then XPOS = InStrRev(TEXT_PATH, ":\") - 1
    If XPOS > 0 And OXPOS = 0 Then OXPOS = 1
    
    'NETWORK
    If XPOS = 0 Then XPOS = InStrRev(TEXT_PATH, "\\")
    If XPOS > 0 And OXPOS = 0 Then OXPOS = 2

    'REG_KEY
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HKEY_")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HCR\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HCU\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HKCU\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HKLM\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HKU\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HU\")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH), "HCC\")
    If XPOS > 0 And OXPOS = 0 Then OXPOS = 3
    
    'WEB
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH_L), "HTTP")
    If XPOS = 0 Then XPOS = InStrRev(UCase(TEXT_PATH_L), "WWW")
    If XPOS > 0 And OXPOS = 0 Then OXPOS = 4


    'CPC WEB
    If InStr(TEXT_PATH_L, "http://cpc.farnell.com/") > 0 Then XPOS = 1
    If XPOS > 0 Then OXPOS = 5

    'EXPLORER FOLDER FILE
    If OXPOS = 1 And XPOS > 0 And I Then Call MNU_CLIPBOARD_EXPLORER_FILE_FOLDER_Click
    If OXPOS = 1 And XPOS > 0 Then TEXT_MNU = "GO -- FILE FOLDER -- OPEN"
    'NETWORK FOLDER FILE
    If OXPOS = 2 And XPOS > 0 And I Then Call MNU_CLIPBOARD_EXPLORER_FILE_FOLDER_Click
    If OXPOS = 2 And XPOS > 0 Then TEXT_MNU = "GO -- NETOWRK FILE FOLDER -- OPEN"
    
    'REG_KEY
    If OXPOS = 3 And XPOS > 0 And I Then Call MNU_REG_JUMP_Click
    If OXPOS = 3 And XPOS > 0 Then TEXT_MNU = "GO -- REGISTRY KEY -- OPEN"
    
    'WEB
    If OXPOS = 4 And XPOS > 0 And I Then Call MNU_URL_Browser_Click
    If OXPOS = 4 And XPOS > 0 Then TEXT_MNU = "GO -- WEB WWW HTTP URL LINK BROWSER -- OPEN"
    
    'CPC WEB
    If OXPOS = 5 Then
        MNU_CPC_TOP_MENU.Visible = False
    Else
        MNU_CPC_TOP_MENU.Visible = True
    
    End If
    
    If OXPOS = 5 And XPOS > 0 And I Then
        Call MNU_CPC_Click
    End If
    If OXPOS = 5 And XPOS > 0 Then
        TEXT_MNU = "GO -- CPC 00 40 80 90"
'        MNU_404_CPC_PAGES.Enabled = True
    
        If InStr(TEXT_PATH_L, LCase("cpc.farnell.com/webapp/wcs/stores/servlet/ProductDisplay")) > 0 Then
            TEXT_MNU = "GO -- CPC 00 40 80 90 -- ERROR REQUIRE PUT ORDER-CODE"
        End If
    End If

    
    'USE CLIPBOARD OR IF NOT CHECK STORED CLIPPER AD$ IN RTB TEXTBOX
    'CLIPBOARD
    For r = 1 To TOP_ARRAY_IAR
        If InStr(UCase(TEXT_PATH), IAR(r)) > 0 Then RESULT_OPENER = True
    Next
    
    
    If RESULT_OPENER = False Then
        MNU_CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY.Caption = "FILE FOLDER && WEB URL && REG KEY && CPC LOADER"
        MNU_CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY.Enabled = False
        MNU_CPC_TOP_MENU.Enabled = False
        MNU_CLIPBOARD_EXPLORER_FILE_FOLDER.Enabled = False
        MNU_REG_JUMP.Enabled = False
        Mnu_URL_Browser.Enabled = False
        MNU_CPC.Enabled = False
        If VAR_MNU_404_HITT_ONCE = False Then MNU_404_CPC_PAGES.Enabled = False
    Else
        MNU_CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY.Caption = "__ " + TEXT_MNU + " __"
        MNU_CHECK_PATH_FOLDER_FILE_URL_REGISTRY_KEY.Enabled = True
    End If
    
End Sub


Function GetWindowsVersion()
Dim OSInfo As OSVERSIONINFO
Dim RetValue As Integer
    OSInfo.dwOSVersionInfoSize = 148
    OSInfo.szCSDVersion = Space$(128)
    RetValue = GetVersionEx(OSInfo)
    sngWindowsVersion = CSng((CStr(OSInfo.dwMajorVersion) & "." & CStr(OSInfo.dwMinorVersion)))
    strWindowsVersion = CStr(OSInfo.dwMajorVersion) & "." & CStr(OSInfo.dwMinorVersion) & "." & CStr(OSInfo.dwBuildNumber)
    GetWindowsVersion = OSInfo.dwMajorVersion
    GetWindowsVersion = CSng((CStr(OSInfo.dwMajorVersion) & "." & CStr(OSInfo.dwMinorVersion)))
  
    '----------------------------------------------------
    'WINDOWS XP PROBLEM _ CHANGE THE SCRIPT HERE
    'NOT TO RUN AS ADMIN OR BRING THE RUNAS DIALOG BOX UP
    'WIN XP = 5.1 _ WINDOWS 10 = 6.2
    '----------------------------------------------------

End Function


Private Sub Txt_Search_CLICK()

    If SEARCH_BOX_NEVER_BEFORE_FOCUS = True Then
        SEARCH_BOX_NEVER_BEFORE_FOCUS = False
        Txt_Search = ""
    End If

End Sub

Private Sub Txt_Search_Change()

    Dim FoundPos As Long
    Dim FoundLine As Long
    
    If Len(Txt_Search.Text) = 0 Then
        Txt_Search.BackColor = RGB(207, 240, 200)
        Exit Sub
    End If
    
    LCase_Text1 = LCase(Text1.Text)
    
    If SEARCH_F3 = "" Then
        FoundPos = InStr(Text1.SelStart, LCase_Text1, LCase(Txt_Search.Text))
        If FoundPos = 0 Then
            FoundPos = InStr(LCase_Text1, LCase(Txt_Search.Text))
        End If
    End If
    
    If SEARCH_F3 = "F3_NEXT" Then
        FoundPos = InStr(Text1.SelStart + 2, LCase_Text1, LCase(Txt_Search.Text))
        If FoundPos = 0 Then
            FoundPos = InStr(1, LCase_Text1, LCase(Txt_Search.Text))
        End If
    End If
    
    If SEARCH_F3 = "F3_REVERSE" Then
        FoundPos = InStrRev(LCase_Text1, LCase(Txt_Search.Text), Text1.SelStart - 2)
        If FoundPos = 0 Then
            FoundPos = InStrRev(LCase_Text1, LCase(Txt_Search.Text), Len(LCase_Text1))
        End If
    End If
    
    SEARCH_F3 = ""
    
    
    ' PRESS F1 MSDN HELP FOR COMMAND
    If Err.Number > 0 Then Exit Sub
    On Error GoTo 0
    
    ' Show message based on whether the text was found or not.
    
    If FoundPos <> 0 Then
       ' Returns number of line containing found text.
       ' FoundLine = Text1.GetLineFromChar(FoundPos)
       
       'MsgBox "Word found on line " & CStr(FoundLine)
       Txt_Search.BackColor = &HFFFFFF
    Else
       'MsgBox "Word not found."
       Txt_Search.BackColor = &HC0C0FF
       Exit Sub
    End If
    
    
     Text1.Enabled = False
     
     'Text1.SelStart = 1
     'Text1.SelStart = Len(Text1.Text)
     
     'Text1.SelStart = Len(Text1.Text)
     
     'Text1.SelStart = IT_TIT1 - 1
     'Text1.SelLength = IT_TIT3
     'Text1.SelStart = IT_TIT2
     
        'Text1.SelColor = RGB(0, 0, 0)
        XFOUNDPOS = FoundPos - 1000
        If XFOUNDPOS < 1 Then XFOUNDPOS = 1
        'Text1.SelStart = XFOUNDPOS
        'Text1.SelLength = 1
    
        Text1.SelStart = FoundPos - 1
        Text1.SelLength = Len(Txt_Search.Text)
        'Text1.SelColor = RGB(100, 255, 100)
    '
     Text1.Enabled = True
     Text1.SetFocus
    

End Sub


Private Sub Txt_Search_KeyDown(KeyCode As Integer, Shift As Integer)

If IsIDE = True And Date = "02/06/2020" Then
    If KeyCode = 27 Then Beep: Unload Me: Exit Sub
End If

' --------------------------------
' SHIFT AND PAGE DOWN -- INITAL
If KeyCode = 16 And Shift = 1 Then
    Exit Sub
End If
' --------------------------------
' SHIFT AND PAGE DOWN -- ANOTHER
If KeyCode = 35 And Shift = 1 Then
    Exit Sub
End If
' --------------------------------

If SEARCH_BOX_NEVER_BEFORE_FOCUS = True Then
    SEARCH_BOX_NEVER_BEFORE_FOCUS = False
    Txt_Search = ""
End If

'F3
If KeyCode = 114 And Shift = 0 Then
    SEARCH_F3 = "F3_NEXT"
    Call Txt_Search_Change
End If

'F3 AND SHIFT
If KeyCode = 114 And Shift = 1 Then
    SEARCH_F3 = "F3_REVERSE"
    Call Txt_Search_Change
End If

End Sub

Private Sub Txt_Search_LostFocus()
If SEARCH_BOX_NEVER_BEFORE_FOCUS = True Then
    SEARCH_BOX_NEVER_BEFORE_FOCUS = False
    Txt_Search = ""
End If
End Sub



Sub MENU_REMOVE_CRAPPY_HARDCORE_1()
        'SUB MENU_OLD_EBAY
   ' Begin VB.Menu MNU_EBAY
      ' Caption         =   "* EBAY GOLD TELEPHONE # * "
      ' Begin VB.Menu Mnu_Line_Spacer2
         ' Caption         =   "--------------------"
      ' End
      ' Begin VB.Menu MNU_TITLE_EBAY_MENU_TAB
         ' Caption         =   "EBAY -- FIND A GOLD TELEPHONE NUMBER -- NOT FINISHED YET"
      ' End
      ' Begin VB.Menu MNU_EBAY_NOTEPAD
         ' Caption         =   "EBAY -- BUY FILTER RESULTS -- LOAD NOTEPAD++"
      ' End
      ' Begin VB.Menu MNU_EBAY_BEGIN_FILTER_SAVE_EDITOR
         ' Caption         =   "EBAY -- FILTER RESULT OF BASIC LINE TEXT SAVE AND LOAD EDITOR"
      ' End
      ' Begin VB.Menu MNU_EBAY_FILTER_WILDCARD
         ' Caption         =   "EBAY -- FILTER LINE INPUT - * - WILDCARD"
      ' End
      ' Begin VB.Menu MNU_EBAY_FILTER_NUMERIC_SIX
         ' Caption         =   "EBAY -- FILTER LINE INPUT - NUMERIC - SIX "
      ' End
      ' Begin VB.Menu MNU_EBAY2
         ' Caption         =   "EBAY_MORE"
      ' End
   ' End
End Sub

Sub MENU_REMOVE_CRAPPY_HARDCORE_2()
         ' Begin VB.Menu MNU_CODE_MENU
            ' Caption         =   "TEST STOP ALL TIMER"
         ' End
End Sub

Sub MENU_REMOVE_CRAPPY_HARDCORE_3()
      ' Begin VB.Menu Mnu_Clips
         ' Caption         =   "Clips"
         ' Begin VB.Menu Mnu_ClipAll
            ' Caption         =   "Clip All"
         ' End
         ' Begin VB.Menu Mnu_ClearClipBoard
            ' Caption         =   "Clear ClipBoard"
         ' End
         ' Begin VB.Menu Mnu_Clear_Text
            ' Caption         =   "Clear Text"
         ' End
         ' Begin VB.Menu Mnu_Test_Clip
            ' Caption         =   "Test Clip"
         ' End
      ' End
End Sub

Sub MENU_NOTE_WORKER_HARDCORE_4()

   ' Begin VB.Menu Mnu_Loggs
      ' Caption         =   "* LOGG MENU *"
      ' Begin VB.Menu Mnu_LoggExplorer
         ' Caption         =   "Open Logg Explorer"
         ' Index           =   1
      ' End
      ' Begin VB.Menu MNU_INFO_RAPID_ALL
         ' Caption         =   "INFO RAPID ALL FOLDERS"
         ' Index           =   2
      ' End
      ' Begin VB.Menu MNU_INFO_RAPID_ALL_SMALL_FILES
         ' Caption         =   "INFO RAPID ALL FOLDERS AND SMALL FILES BEGIN ClipBoard-*.TXT"
         ' Index           =   3
      ' End
      ' Begin VB.Menu MNU_INFO_RAPID_MYUSER
         ' Caption         =   "INFO RAPID MY USER"
         ' Index           =   4
      ' End
      ' Begin VB.Menu MNU_INFO_RAPID
         ' Caption         =   "INFO RAPID TRIM LOGG"
         ' Index           =   5
      ' End
      ' Begin VB.Menu Mnu_Open_Recent_Hex
         ' Caption         =   "Hex Open Recent Trim Logg"
         ' Index           =   6
      ' End
      ' Begin VB.Menu MNU__ML_13
         ' Caption         =   "----"
         ' Index           =   7
      ' End
      ' Begin VB.Menu Mnu_Open_Recent
         ' Caption         =   "Edit Recent Trim Logg"
         ' Index           =   8
      ' End
      ' Begin VB.Menu Mnu_Open_Logg
         ' Caption         =   "Edit This Logg"
         ' Index           =   9
      ' End
      ' Begin VB.Menu Mnu_Open_Total
         ' Caption         =   "Edit Total Logg"
         ' Index           =   10
      ' End
      ' Begin VB.Menu Mnu_StripLogg
         ' Caption         =   "Edit Strip Logg"
         ' Index           =   11
      ' End
      ' Begin VB.Menu MNU__ML_14
         ' Caption         =   "----"
         ' Index           =   12
      ' End
      ' Begin VB.Menu Mnu_TEXT_Open_Recent
         ' Caption         =   "Text View Recent Trim Logg"
         ' Index           =   13
      ' End
      ' Begin VB.Menu Mnu_TEXT_Open_Logg
         ' Caption         =   "Text View This Logg"
         ' Index           =   14
      ' End
      ' Begin VB.Menu Mnu_TEXT_Open_Total
         ' Caption         =   "Text View Total Logg"
         ' Index           =   15
      ' End
      ' Begin VB.Menu Mnu_TEXT_StripLogg
         ' Caption         =   "Text View Strip Logg"
         ' Index           =   16
      ' End
      ' Begin VB.Menu MNU__ML_15
         ' Caption         =   "----"
         ' Index           =   17
      ' End
      ' Begin VB.Menu Mnu_ShellT_Todays
         ' Caption         =   "Shell T -- This Logg"
         ' Index           =   18
      ' End
      ' Begin VB.Menu Mnu_Shell_T
         ' Caption         =   "Shell T -- Total Logg"
         ' Index           =   19
      ' End
      ' Begin VB.Menu MNU__ML_17
         ' Caption         =   "----"
         ' Index           =   20
      ' End
   ' End

End Sub


Sub MENU_REMOVE_CRAPPY_HARDCORE_5()

   ' Begin VB.Menu MNU_SCROLL_DOWN
      ' Caption         =   "SCROLL DOWN"
      ' Begin VB.Menu MNU_SCROLL_DOWN_ARRAY_MNU
         ' Caption         =   "400"
         ' Index           =   1
      ' End
      ' Begin VB.Menu MNU_SCROLL_DOWN_ARRAY_MNU
         ' Caption         =   "100"
         ' Index           =   2
      ' End
      ' Begin VB.Menu MNU_SCROLL_DOWN_ARRAY_MNU
         ' Caption         =   "OFF"
         ' Index           =   3
      ' End
   ' End

End Sub
