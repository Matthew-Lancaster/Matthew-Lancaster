VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Winamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API
'******************************************************************************
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'******************************************************************************

'Constants
'******************************************************************************
Private Const m_Const_Winamp As String = "Winamp v1.x" 'This is used get the Winamp.exe instance
Private Const m_WM_COMMAND = &H111
'******************************************************************************

'Variables
'******************************************************************************
'******************************************************************************

'Enumerations
'******************************************************************************
'Check http://www.winamp.com for more Messages
Public Enum Enum_WA_Command
  WA_Previous_track = 40044
  WA_Next_track = 40048
  WA_Play = 40045
  WA_Pause_Unpause = 40046
  WA_Stop = 40047
  WA_Fadeout_and_stop = 40147
  WA_Stop_after_current = 40157
  WA_Fast_forward_5_seconds = 40148
  WA_Fast_rewind_5_seconds = 40144
  WA_Start_of_playlist = 40154
  WA_Go_to_end_of_playlist = 40158
  WA_Open_file_dialog = 40029
  WA_Open_URL_dialog = 40155
  WA_Open_file_info_box = 40188
  WA_Set_time_display_mode_to_elapsed = 40037
  WA_Set_time_display_mode_to_remaining = 40038
  WA_Toggle_preferences_screen = 40012
  WA_Open_visualization_options = 40190
  WA_Open_visualization_plug_in_options = 40191
  WA_Execute_current_visualization_plug_in = 40192
  WA_Toggle_about_box = 40041
  WA_Toggle_title_Autoscrolling = 40189
  WA_Toggle_always_on_top = 40019
  WA_Toggle_Windowshade = 40064
  WA_Toggle_Playlist_Windowshade = 40266
  WA_Toggle_doublesize_mode = 40165
  WA_Toggle_EQ = 40036
  WA_Toggle_playlist_editor = 40040
  WA_Toggle_main_window_visible = 40258
  WA_Toggle_minibrowser = 40298
  WA_Toggle_easymove = 40186
  WA_Raise_volume_by_1 = 40058
  WA_Lower_volume_by_1 = 40059
  WA_Toggle_repeat = 40022
  WA_Toggle_shuffle = 40023
  WA_Open_jump_to_time_dialog = 40193
  WA_Open_jump_to_file_dialog = 40194
  WA_Open_skin_selector = 40219
  WA_Configure_current_visualization_plug_in = 40221
  WA_Reload_the_current_skin = 40291
  WA_Close_Winamp = 40001
  WA_Moves_back_10_tracks_in_playlist = 40197
  WA_Show_the_edit_bookmarks = 40320
  WA_Adds_current_track_as_a_bookmark = 40321
  WA_Play_audio_CD = 40323
  WA_Load_a_preset_from_EQ = 40253
  WA_Save_a_preset_to_EQF = 40254
  WA_Opens_load_presets_dialog = 40172
  WA_Opens_auto_load_resets_dialog = 40173
  WA_Load_default_preset = 40174
  WA_Opens_save_preset_dialog = 40175
  WA_Opens_auto_load_save_preset = 40176
  WA_Opens_delete_preset_dialog = 40178
  WA_Opens_delete_an_auto_load_preset_dialog = 40180

End Enum
'******************************************************************************

'Properties
'******************************************************************************
'******************************************************************************

'Subs
'******************************************************************************
Public Function IsWinampRunning() As Boolean

  'FindWindow is used to get the hwnd of the Winamp.exe instance
  IsWinampRunning = IIf(FindWindow(m_Const_Winamp, vbNullString) = 0, False, True)

End Function

Public Sub Execute(ByVal l_Command As Enum_WA_Command)
  
  SendMessage FindWindow(m_Const_Winamp, vbNullString), m_WM_COMMAND, l_Command, 0

End Sub
'******************************************************************************

'Functions
'******************************************************************************
'******************************************************************************

