VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmWebBrowser 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   Caption         =   "HCL Applications"
   ClientHeight    =   6315
   ClientLeft      =   3060
   ClientTop       =   3630
   ClientWidth     =   11880
   Icon            =   "WebBrowser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet ctlInet 
      Left            =   7410
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2985
      Left            =   30
      TabIndex        =   4
      Top             =   1350
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   5265
      _Version        =   393217
      BackColor       =   16777152
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"WebBrowser.frx":030A
   End
   Begin SHDocVwCtl.WebBrowser ctlWebBrowser 
      Height          =   3060
      Left            =   60
      TabIndex        =   0
      Top             =   1350
      Width           =   5400
      ExtentX         =   9525
      ExtentY         =   5397
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7440
      Top             =   2520
   End
   Begin VB.PictureBox picAddressContainer 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   11820
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   11880
      Begin VB.ComboBox cboAddress 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   30
         Width           =   5145
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address:"
         Height          =   165
         Left            =   60
         TabIndex        =   3
         Top             =   90
         Width           =   825
      End
   End
   Begin MSComctlLib.ImageList Imagelist1 
      Left            =   7350
      Top             =   1830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":03C4
            Key             =   "tbrBack"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":0718
            Key             =   "tbrForward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":0A6C
            Key             =   "tbrStop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":0DC0
            Key             =   "tbrRefresh"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":1114
            Key             =   "tbrHome"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":12BC
            Key             =   "tbrSearch"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":1610
            Key             =   "tbrFavoritesList"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":3DC4
            Key             =   "tbrPrintDesktop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":40E0
            Key             =   "tbrCapDesktop"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":43FC
            Key             =   "tbrInbox"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":6BB0
            Key             =   "tbrOutbox"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":7004
            Key             =   "tbrContacts"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":7AD0
            Key             =   "tbrDisplayCurrHTMLDoc"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":8724
            Key             =   "tbrListLinks"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":8880
            Key             =   "tbrDisplayLinkHTMLDoc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":8B9C
            Key             =   "tbrSaveHTMLDoc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":8CF8
            Key             =   "tbrPrintHTMLDoc"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "WebBrowser.frx":9014
            Key             =   "tbrRestoreWebBrowser"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFavouritesListContainer 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11820
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   795
      Width           =   11880
      Begin VB.PictureBox picLinksListContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   0
         ScaleHeight     =   435
         ScaleWidth      =   11850
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   11880
         Begin VB.ComboBox cboLinksList 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   930
            TabIndex        =   18
            Top             =   60
            Width           =   5115
         End
         Begin VB.Label lblLinks 
            Caption         =   "Links:"
            ForeColor       =   &H00800000&
            Height          =   165
            Left            =   60
            TabIndex        =   19
            Top             =   120
            Width           =   825
         End
      End
      Begin VB.PictureBox pic1 
         Height          =   345
         Index           =   3
         Left            =   2640
         ScaleHeight     =   285
         ScaleWidth      =   225
         TabIndex        =   15
         Top             =   60
         Width           =   285
         Begin VB.CommandButton cmdExitFavoritesList 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            Picture         =   "WebBrowser.frx":9170
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Exit Favorites List (return to Links)"
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox pic1 
         Height          =   345
         Index           =   2
         Left            =   2310
         ScaleHeight     =   285
         ScaleWidth      =   225
         TabIndex        =   13
         Top             =   60
         Width           =   285
         Begin VB.CommandButton cmdDeleteFavorite 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            Picture         =   "WebBrowser.frx":92BA
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Remove from list"
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.PictureBox pic1 
         Height          =   345
         Index           =   1
         Left            =   1980
         ScaleHeight     =   285
         ScaleWidth      =   225
         TabIndex        =   11
         Top             =   60
         Width           =   285
         Begin VB.CommandButton cmdAddFavorite 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            Picture         =   "WebBrowser.frx":9404
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Add to Favorites List"
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.ComboBox cboFavoritesList 
         BackColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   60
         Width           =   7485
      End
      Begin VB.PictureBox pic1 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   0
         Left            =   1650
         ScaleHeight     =   285
         ScaleWidth      =   225
         TabIndex        =   20
         Top             =   60
         Width           =   285
         Begin VB.CommandButton cmdSelectFavorite 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            Picture         =   "WebBrowser.frx":954E
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Select favorite "
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.Label lblFavorites 
         Caption         =   "Favourites:"
         ForeColor       =   &H000000FF&
         Height          =   165
         Left            =   600
         TabIndex        =   7
         Top             =   150
         Width           =   825
      End
   End
   Begin MSComctlLib.Toolbar WebToolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "Imagelist1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrBack"
            Description     =   "Back"
            Object.ToolTipText     =   "Back"
            ImageKey        =   "tbrBack"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Forward"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrForward"
            Object.ToolTipText     =   "Forward"
            ImageKey        =   "tbrForward"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrStop"
            Object.ToolTipText     =   "Stop"
            ImageKey        =   "tbrStop"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrRefresh"
            Object.ToolTipText     =   "Refresh"
            ImageKey        =   "tbrRefresh"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrHome"
            Object.ToolTipText     =   "Home"
            ImageKey        =   "tbrHome"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrSearch"
            Object.ToolTipText     =   "Search"
            ImageKey        =   "tbrSearch"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Back"
                  Text            =   "Back"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Forward"
                  Text            =   "Forward"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Stop"
                  Text            =   "Stop"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Refresh"
                  Text            =   "Refresh"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Home"
                  Text            =   "Home"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrFavoritesList"
            Object.ToolTipText     =   "Display Favorites List"
            ImageKey        =   "tbrFavoritesList"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrPrintDesktop"
            Object.ToolTipText     =   "Print desktop"
            ImageKey        =   "tbrPrintDesktop"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrCapDesktop"
            Object.ToolTipText     =   "Capture desktop to file directly"
            ImageKey        =   "tbrCapDesktop"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrInbox"
            Object.ToolTipText     =   "Mail Inbox"
            ImageKey        =   "tbrInbox"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrOutbox"
            Object.ToolTipText     =   "Mail Outbox"
            ImageKey        =   "tbrOutbox"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrContacts"
            Object.ToolTipText     =   "Contacts"
            ImageKey        =   "tbrContacts"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrDisplayCurrHTMLDoc"
            Object.ToolTipText     =   "Display current HTML document"
            ImageKey        =   "tbrDisplayCurrHTMLDoc"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrListLinks"
            Object.ToolTipText     =   "List links"
            ImageKey        =   "tbrListLinks"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrDisplayLinkHTMLDoc"
            Object.ToolTipText     =   "Display link HTML document"
            ImageKey        =   "tbrDisplayLinkHTMLDoc"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrPrintHTMLDoc"
            Object.ToolTipText     =   "Print displayed HTML doc (highlighted or all)"
            ImageKey        =   "tbrPrintHTMLDoc"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrSaveHTMLDoc"
            Object.ToolTipText     =   "Save displayed HTML document"
            ImageKey        =   "tbrSaveHTMLDoc"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrRestoreWebBrowser"
            Object.ToolTipText     =   "Restore web browser"
            ImageKey        =   "tbrRestoreWebBrowser"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ProgressBar tbrProgress 
         Height          =   195
         Left            =   9840
         TabIndex        =   10
         Top             =   60
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.PictureBox picAnimation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   11550
         Picture         =   "WebBrowser.frx":9698
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         ToolTipText     =   "Restore web browser"
         Top             =   30
         Width           =   285
      End
   End
   Begin VB.Menu mnuGo 
      Caption         =   "&Go"
      Begin VB.Menu mnuGoBack 
         Caption         =   "&Back"
      End
      Begin VB.Menu mnuGoForward 
         Caption         =   "&Forward"
      End
      Begin VB.Menu mnuGosep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnuGoRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuGosep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoHome 
         Caption         =   "&Home"
      End
      Begin VB.Menu mnuGosep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoExit 
         Caption         =   "&Exist"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileFavoritesList 
         Caption         =   "&Favorites list"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuURLHistory 
         Caption         =   "URLHistory"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuURLHistory 
         Caption         =   "URLHistory1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuURLHistory 
         Caption         =   "URLHistory2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuURLHistory 
         Caption         =   "URLHistory3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuURLHistory 
         Caption         =   "URLHistory4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuURLHistory 
         Caption         =   "URLHistory5"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuURLHistory 
         Caption         =   "URLHistory6"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuURLHistory 
         Caption         =   "URLHistory7"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuURLHistory 
         Caption         =   "URLHistory8"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuURLHistory 
         Caption         =   "URLHistory9"
         Index           =   9
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuOptionsStatusbar 
         Caption         =   "&Statusbar"
      End
   End
   Begin VB.Menu mnuCopy 
      Caption         =   "&Copy"
      Begin VB.Menu mnuCopyPrintDesktop 
         Caption         =   "&Print desktop"
      End
      Begin VB.Menu mnuCopysep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyCapDeskTop 
         Caption         =   "&Capture desktop to file direct"
      End
      Begin VB.Menu mnuCopysep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopySaveDesktopFromClip 
         Caption         =   "Save &desktop  from clipboard (PrnScreen first)"
      End
      Begin VB.Menu mnuCopySaveTextFromClip 
         Caption         =   "Save screen &text"
      End
   End
   Begin VB.Menu mnuMail 
      Caption         =   "&Mail"
      Begin VB.Menu mnuMailInbox 
         Caption         =   "&Inbox"
      End
      Begin VB.Menu mnuMailOutbox 
         Caption         =   "&Outbox"
      End
      Begin VB.Menu mnuMailContacts 
         Caption         =   "&Contacts"
      End
   End
   Begin VB.Menu mnuLinks 
      Caption         =   "&Links"
      Begin VB.Menu mnuLinksDisplayCurrHTMLDoc 
         Caption         =   "Dispaly &current HTML doc"
      End
      Begin VB.Menu mnuLinksep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLinksListLinks 
         Caption         =   "&List links"
      End
      Begin VB.Menu mnuLinksDisplayLinkHTMLDoc 
         Caption         =   "&Dispaly Link HTML doc"
      End
      Begin VB.Menu mnuLinksep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLinksPrintHTMLdoc 
         Caption         =   "&Print HTML doc (highlighted or all)"
      End
      Begin VB.Menu mnuLinksSaveHTMLdoc 
         Caption         =   "&Save HTML doc"
      End
      Begin VB.Menu mnuLinksep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLinksRestoreWebBrowser 
         Caption         =   "&Restore web browser"
      End
   End
End
Attribute VB_Name = "frmWebBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' WebBrowser.frm
'
' By Herman Liu
'
' This code, apart from providing the usual navigational gadgets of a WebBrowser control,
' possesses the following additional attributes: (1) Has mail stuff (Inbox, Outbox and
' Contacts); (2) Maintains a Favorites List (interestingly, at the background using only
' a tiny text file as the database, rather than a wasteful big mdb file); (3) Shows URLs
' history in File menu, ready for click (URL history is kept through registry, hence no
' extra file required); (4) Captures images of desktop screens that you like and saves
' them to files while you are on Net and (5) Attempts to pinch HTML documents of a URL
' (to list links, to display a selected HTML document, and to print and/or save it), if
' you wish.
'
' ------
' Notes:
'     Include "MS Outlook 98 Object Model" in Project's Reference.
'     If you don't have Outlook, then block/remove at least the following lines (you can
'     leave the menu items and toolbar menus to remain there if you wish):
'
'     (a) Dim oloAPP As Outlook.Application
'         All  ....olo....  lines
'
'     (b) Sub Outlooklogon(mWhat As String)
'          ....
'         End Sub
'
'     (c) (In Sub WebToolBar_ButtonClick)
'         Case "TBRINBOX"
'            Outlooklogon "Inbox"
'         Case "TBROUTBOX"
'            Outlooklogon "Outbox"
'         Case "TBRCONTACTS"
'            Outlooklogon "Contacts"
'
'----------------------------------------------------------------------------
Option Explicit

Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_MENU = &H12
Private Const VK_SNAPSHOT = &H2C
Private Const KEYEVENTF_KEYUP = &H2

Dim FavoriteFileName As String
Dim mbDontNavigateNow As Boolean
Dim mFrameNum As Integer
Dim mForward As Boolean
Dim mBackward As Boolean

Dim oloAPP As Outlook.Application
Dim oloNS As Outlook.NameSpace
Dim oloFolder As Outlook.MAPIFolder
Dim oloNewMail As Outlook.MailItem
Dim oloAddrList As Outlook.AddressList
Dim oloAddrEntry As Outlook.AddressEntry
Dim oloSharedFolder As Outlook.MAPIFolder
Dim oloRecipient As Outlook.Recipient

Dim URLFailedFlag As Boolean
Dim origToolBarOn As Boolean
Dim gcdg As Object
Const gconThisapp = "HCLappDB"
Const gconURLKey = "URL History"




Private Sub Form_Load()
    On Error Resume Next
    origToolBarOn = frmFrame.FrameToolBar.Visible
       ' Get rid of existing main toolbar
    frmFrame.FrameToolBar.Visible = False
    
    Me.AutoRedraw = True
    mFrameNum = 1
    mForward = True
    mBackward = False
    
    mnuOptionsToolbar.Checked = True
    
      ' Test existence of Favorites.rtf, if not, will create one.
    FavoriteFileName = App.Path & "\Favorites.txt"
    Dim mHandle
    mHandle = FreeFile
    Open FavoriteFileName For Binary As #mHandle
    Close #mHandle
    
    Me.Show
    Form_Resize
    
    timer1.Interval = 500
    
    Text1.Visible = False
    ctlWebBrowser.Visible = True
    picLinksListContainer.Visible = True
    ToggleLinksMenus True
    GetRecentURLs
    
      ' Use whatever user's default home address
    mnuGoHome_Click
End Sub



Private Sub Form_Activate()
    FillcboFavoritesList
End Sub



Private Sub FillcboFavoritesList()
    Dim mHandle
    Dim tmp As String
    cboFavoritesList.Clear
    mHandle = FreeFile
    Open FavoriteFileName For Input As #mHandle
    Do While Not EOF(mHandle)
        Line Input #mHandle, tmp
        If Len(Trim(tmp)) > 0 Then
            cboFavoritesList.AddItem tmp
        End If
    Loop
    Close #mHandle
    If cboFavoritesList.ListCount > 0 Then
        cboFavoritesList.ListIndex = 0
    End If
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Screen.MousePointer = vbDefault
    frmFrame.FrameToolBar.Visible = origToolBarOn
    ctlInet.Cancel
End Sub




Private Sub Form_Resize()
    On Error Resume Next
    cboAddress.Width = Me.ScaleWidth - cboAddress.Left - 100
    cboLinksList.Width = Me.ScaleWidth - cboLinksList.Left - 100
    ctlWebBrowser.Width = Me.ScaleWidth - 100
    ctlWebBrowser.Height = Me.ScaleHeight - (picLinksListContainer.Top + _
        picLinksListContainer.Height) - 500
    Text1.Width = ctlWebBrowser.Width
    Text1.Height = ctlWebBrowser.Height
End Sub



Private Sub cboAddress_Click()
    If mbDontNavigateNow Then
         Exit Sub
    End If
    cboLinksList.Clear
    doNavigate cboAddress.Text
    If URLFailedFlag = False Then
         WriteRecentURLs cboAddress.Text
         GetRecentURLs
         ctlWebBrowser.SetFocus
    End If
End Sub



Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub



Private Sub doNavigate(inAddress As String)
    Screen.MousePointer = vbHourglass
    If Len(inAddress) > 0 Then
        cboLinksList.Clear
        mnuLinksRestoreWebBrowser_Click
          'try to navigate to the address
        timer1.Enabled = True
        ctlWebBrowser.Navigate inAddress
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub ctlWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = ctlWebBrowser.LocationName
End Sub



Private Sub ctlWebBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = ctlWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = ctlWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
        URLFailedFlag = False
    Else
        URLFailedFlag = True
    End If
    cboAddress.AddItem ctlWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub





' For frmWebBrowser's own toolbar
Private Sub mnuOptionsToolbar_Click()
    Me.mnuOptionsToolbar.Checked = Not Me.mnuOptionsToolbar.Checked
    If Me.mnuOptionsToolbar.Checked Then
        Me.WebToolbar.Visible = True
        ctlWebBrowser.Top = ctlWebBrowser.Top + Me.WebToolbar.Height
    Else
        Me.WebToolbar.Visible = False
        ctlWebBrowser.Top = ctlWebBrowser.Top - Me.WebToolbar.Height
    End If
End Sub



' For frmFrame's status bar
Private Sub mnuOptionsStatusbar_Click()
    OptionsStatusbarProc Me
End Sub





Private Sub mnuOptions_Click()
    mnuOptionsToolbar.Checked = Me.WebToolbar.Visible
    mnuOptionsStatusbar.Checked = frmFrame.StatusBar1.Visible
End Sub




Private Sub mnuLinksDisplayCurrHTMLDoc_Click()
     On Error Resume Next
     Screen.MousePointer = vbHourglass
     Dim HTMLcontent
     picLinksListContainer.Visible = True
     Text1.Text = ""
     HTMLcontent = ctlWebBrowser.Document.Body.innerHTML
     Text1.SelRTF = HTMLcontent
     ctlWebBrowser.Visible = False
     Text1.Visible = True
     ToggleLinksMenus False
     Text1.SetFocus
     Screen.MousePointer = vbDefault
End Sub




Private Sub mnuLinksListLinks_Click()
     If cboAddress.Text = "" Then
         MsgBox "No URL/address yet"
         Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     ctlWebBrowser.Visible = True
     Text1.Visible = False
      
     picLinksListContainer.Visible = True
      ' Fill cboLinksList with available links
     cboLinksList.Clear
     Dim i
     For i = 0 To ctlWebBrowser.Document.links.Length - 1
          cboLinksList.AddItem ctlWebBrowser.Document.links(i).href
              ' We wish to limit the number
          If i >= 100 Then
               Exit For
          End If
     Next i
     Screen.MousePointer = vbDefault
     If i = 0 Then
          MsgBox "No links listed, try a different level or a different URL"
     Else
          cboLinksList.ListIndex = 0
          ctlWebBrowser.SetFocus
     End If
 End Sub



Private Sub mnuLinksDisplayLinkHTMLDoc_Click()
     On Error Resume Next
     If cboLinksList.Text = "" Then
          MsgBox "No listed links yet"
          Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     Dim HTMLcontent

     picLinksListContainer.Visible = True
     Text1.Text = ""

      ' Open the URL/try to open.
     HTMLcontent = ctlInet.OpenURL(cboLinksList.Text)
     
     Text1.SelRTF = HTMLcontent
     
     ctlWebBrowser.Visible = False
     Text1.Visible = True
     ToggleLinksMenus False
     Text1.SetFocus
     
     Screen.MousePointer = vbDefault
End Sub



Private Sub mnuLinksPrintHTMLdoc_Click()
    On Error GoTo Errhandler
    
       ' Sometimes, when there is no apparent links are produced by
       ' the internal arrangement of the site. In that case, if you
       ' know the URL, type in the URL in the cboLinkList, you may
       ' get the HTML of that URL
    'If cboLinksList.ListCount = 0 Then
    '     MsgBox "No HTML document yet"
    '     Exit Sub
    'End If
    If Text1.Text = "" Then
        MsgBox "You have not yet fetched HTML document or HTML document empty"
          Exit Sub
    End If

    Set gcdg = frmFrame.CommonDialog1
    gcdg.DialogTitle = "Print"
    gcdg.CancelError = True
    gcdg.Flags = cdlPDReturnDC + cdlPDPageNums
    gcdg.ShowPrinter
    If Err = MSComDlg.cdlCancel Then
         Exit Sub
    End If
   
    Text1.SelPrint gcdg.hDC, True
    Set gcdg = Nothing
    Exit Sub
     
Errhandler:
     If Err <> 32755 Then
          ErrMsgProc "mnuLinkPrintHTMLDoc_Click"
     End If
End Sub



Private Sub mnuLinksSaveHTMLDoc_Click()
     On Error GoTo Errhandler
    
       ' Sometimes, when there is no apparent links are produced by
       ' the internal arrangement of the site. In that case, if you
       ' know the URL, type in the URL in the cboLinkList, you may
       ' get the HTML of that URL
     'If cboLinksList.ListCount = 0 Then
     '     MsgBox "No HTML document yet"
     '     Exit Sub
     'End If
     If Text1.Text = "" Then
          MsgBox "You have not yet fetched HTML document or HTML document empty"
          Exit Sub
     End If

     Dim mfilespec As String
     Set gcdg = frmFrame.CommonDialog1
     gcdg.FileName = mfilespec
     gcdg.Flags = cdlOFNOverwritePrompt
     gcdg.ShowSave
     If gcdg.FileName = "" Then
          Exit Sub
     End If
     mfilespec = gcdg.FileName
    
     Screen.MousePointer = vbHourglass
     If mfilespec <> "" Then
          Text1.SaveFile mfilespec, 1
     End If
     Set gcdg = Nothing
     Screen.MousePointer = vbDefault
     Exit Sub
     
Errhandler:
     Screen.MousePointer = vbDefault
     If Err <> 32755 Then
          ErrMsgProc "mnuLinkSaveHTMLDoc_Click"
     End If
End Sub



Private Sub mnuLinksDisplayHTMLDoc_Click()
     On Error Resume Next                   ' May disp empty page, e.g if not connected
     If cboLinksList.Text = "" Then
          MsgBox "No listed links yet"
          Exit Sub
     End If
     Dim HTMLcontent

     Screen.MousePointer = vbHourglass
     picLinksListContainer.Visible = True
     Text1.Text = ""
      ' Open the URL / try to open
     HTMLcontent = ctlInet.OpenURL(cboLinksList.Text)
     Text1.SelRTF = HTMLcontent
     ctlWebBrowser.Visible = False
     Text1.Visible = True
     ToggleLinksMenus False
     Text1.SetFocus
     Screen.MousePointer = vbDefault
End Sub



Private Sub mnuLinksRestoreWebBrowser_Click()
     If ctlWebBrowser.Visible = True Then
          Text1.Visible = False
          Exit Sub
     End If
     picLinksListContainer.Visible = True
     Text1.Text = ""
     Text1.Visible = False
     ctlWebBrowser.Visible = True
     ToggleLinksMenus True
End Sub




Private Sub ToggleLinksMenus(Onoff As Boolean)
     Screen.MousePointer = vbHourglass
     mnuGo.Enabled = Onoff
     mnuFile.Enabled = Onoff
     mnuCopy.Enabled = Onoff
     WebToolbar.Buttons("tbrBack").Enabled = Onoff
     WebToolbar.Buttons("tbrForward").Enabled = Onoff
     WebToolbar.Buttons("tbrStop").Enabled = Onoff
     WebToolbar.Buttons("tbrHome").Enabled = Onoff
     WebToolbar.Buttons("tbrRefresh").Enabled = Onoff
     WebToolbar.Buttons("tbrSearch").Enabled = Onoff
     WebToolbar.Buttons("tbrFavoritesList").Enabled = Onoff
     WebToolbar.Buttons("tbrPrintDesktop").Enabled = Onoff
     WebToolbar.Buttons("tbrCapDesktop").Enabled = Onoff
     mnuLinksPrintHTMLdoc.Enabled = Not Onoff
     mnuLinksSaveHTMLdoc.Enabled = Not Onoff
     WebToolbar.Buttons("tbrPrintHTMLDoc").Enabled = Not Onoff
     WebToolbar.Buttons("tbrSaveHTMLDoc").Enabled = Not Onoff
     Screen.MousePointer = vbDefault
End Sub




Private Sub Timer1_Timer()
    If ctlWebBrowser.Busy = False Then
        Me.Caption = ctlWebBrowser.LocationName
        Me.tbrProgress.Value = 0
    Else
        Me.Caption = "Processing..."
    End If
End Sub




Private Sub WebToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
    timer1.Enabled = True
    Me.tbrProgress.Value = 90
    Select Case UCase(Button.key)
        Case "TBRBACK"
            ctlWebBrowser.GoBack
        Case "TBRFORWARD"
            ctlWebBrowser.GoForward
        Case "TBRSTOP"
            Me.tbrProgress.Value = 0
            timer1.Enabled = False
            ctlWebBrowser.Stop
            Me.Caption = ctlWebBrowser.LocationName
        Case "TBRREFRESH"
            ctlWebBrowser.Refresh
        Case "TBRHOME"
            ctlWebBrowser.GoHome
        Case "TBRSEARCH"
            ctlWebBrowser.GoSearch
        Case "TBRFAVORITESLIST"
            mnuFileFavoritesList_Click
        Case "TBRPRINTDESKTOP"
            mnuCopyPrintDesktop_Click
        Case "TBRCAPDESKTOP"
            mnuCopyCapDeskTop_Click
        Case "TBRINBOX"
            Outlooklogon "Inbox"
        Case "TBROUTBOX"
            Outlooklogon "Outbox"
        Case "TBRCONTACTS"
            Outlooklogon "Contacts"
        Case "TBRDISPLAYCURRHTMLDOC"
            mnuLinksDisplayCurrHTMLDoc_Click
        Case "TBRLISTLINKS"
            mnuLinksListLinks_Click
        Case "TBRDISPLAYLINKHTMLDOC"
            mnuLinksDisplayLinkHTMLDoc_Click
        Case "TBRPRINTHTMLDOC"
            mnuLinksPrintHTMLdoc_Click
        Case "TBRSAVEHTMLDOC"
            mnuLinksSaveHTMLDoc_Click
        Case "TBRRESTOREWEBBROWSER"
            mnuLinksRestoreWebBrowser_Click
    End Select
End Sub



Private Sub picAnimation_Click()
    mnuLinksRestoreWebBrowser_Click
End Sub




Private Sub mnuGoBack_Click()
    WebToolBar_ButtonClick frmWebBrowser.WebToolbar.Buttons("tbrBack")
End Sub



Private Sub mnuGoForward_Click()
    WebToolBar_ButtonClick frmWebBrowser.WebToolbar.Buttons("tbrForward")
End Sub


Private Sub mnuGoHome_Click()
      ' Ensure ctlWebBrowser to be visible
    mnuLinksRestoreWebBrowser_Click
    WebToolBar_ButtonClick frmWebBrowser.WebToolbar.Buttons("tbrHome")
End Sub



Private Sub mnuGoRefresh_Click()
    mnuLinksRestoreWebBrowser_Click
    WebToolBar_ButtonClick frmWebBrowser.WebToolbar.Buttons("tbrRefresh")
End Sub



Private Sub mnuGoStop_Click()
    WebToolBar_ButtonClick frmWebBrowser.WebToolbar.Buttons("tbrStop")
End Sub


Private Sub mnuGoExit_Click()
    Unload Me
End Sub


Private Sub mnuFileSearch_Click()
    WebToolBar_ButtonClick frmWebBrowser.WebToolbar.Buttons("tbrSearch")
End Sub



Private Sub mnuFileFavoritesList_Click()
    picLinksListContainer.Visible = False
End Sub



Private Sub mnuURLHistory_Click(Index As Integer)
    cboAddress.Text = mnuURLHistory(Index).Caption
    cboAddress.SetFocus
End Sub




Private Sub GetRecentURLs()
    On Error Resume Next
    Dim i As Integer
    Dim varURLs As Variant      '
    Dim key As String
    If GetSetting(gconThisapp, gconURLKey, "URLHistory1") = Empty Then
        Exit Sub
    End If
    varURLs = GetAllSettings(gconThisapp, gconURLKey)
    For i = 0 To UBound(varURLs, 1)
          mnuURLHistory(0).Visible = True
          mnuURLHistory(i).Caption = varURLs(i, 1)
          mnuURLHistory(i).Visible = True
    Next i
End Sub



Sub WriteRecentURLs(inURL As String)
    If LTrim(Trim(inURL)) = "" Then
        Exit Sub
    End If
    Dim i As Integer
    Dim strURL, key As String
    Dim arrList(8) As String
    For i = 9 To 1 Step -1
        key = "URLHistory" & i
        strURL = GetSetting(gconThisapp, gconURLKey, key)
        If strURL <> "" Then
            key = "URLHistory" & (i + 1)
            SaveSetting gconThisapp, gconURLKey, key, strURL
            arrList(i - 1) = strURL
        End If
    Next i
    SaveSetting gconThisapp, gconURLKey, "URLHistory1", inURL
End Sub



Private Sub mnuMailInbox_Click()
    WebToolBar_ButtonClick frmWebBrowser.WebToolbar.Buttons("tbrInbox")
End Sub


Private Sub mnuMailOutbox_Click()
    WebToolBar_ButtonClick frmWebBrowser.WebToolbar.Buttons("tbrOutbox")
End Sub



Private Sub mnuMailContacts_Click()
    WebToolBar_ButtonClick frmWebBrowser.WebToolbar.Buttons("tbrContacts")
End Sub




Private Sub Outlooklogon(mWhat As String)
     ' Ignore errors first, when we look for a running copy
    On Error Resume Next
    Set oloAPP = GetObject(, "Outlook.Application")
    If Err.Number <> 0 Then          'If Outlook is not running then
        Set oloAPP = CreateObject("Outlook.Application")
    End If
    Err.Clear   ' In case error occurred.
    
    On Error GoTo Errhandler
       ' Get namespace for "MAPI"
    Set oloNS = oloAPP.GetNamespace("MAPI")
       ' Let the user logon to Outlook with the Outlook Profile dialog box,
       ' and then create a new session.
    oloNS.Logon "", "", True, True
    Select Case UCase(mWhat)
        Case "INBOX"
            GetInbox
        Case "OUTBOX"
            GetOutbox
        Case "CONTACTS"
            GetContacts
    End Select
    CleanUpOutlook
    Exit Sub
    
Errhandler:
   ErrMsgProc "OutlookLogon"
End Sub



Private Sub GetInbox()
    Set oloFolder = oloNS.GetDefaultFolder(olFolderInbox)
      'Display the Inbox in a new Explorer window
    oloFolder.Display
End Sub



Private Sub GetOutbox()
    Set oloFolder = oloNS.GetDefaultFolder(olFolderOutbox)
    oloFolder.Display
End Sub



Private Sub GetContacts()
    Set oloFolder = oloNS.GetDefaultFolder(olFolderContacts)
    oloFolder.Display
End Sub



Private Sub CleanUpOutlook()
    oloNS.Logoff
    Set oloNS = Nothing
    Set oloAPP = Nothing
End Sub



Private Sub GetSharedFolder(strRecipient As String)
      'Create a new recipient object and resolve it.
    Set oloRecipient = oloNS.CreateRecipient(strRecipient)
    oloRecipient.Resolve
    'If this user exists on the Exchange server..
    If oloRecipient.Resolved Then
          'Get the shared calendar folder
        Set oloSharedFolder = oloNS.GetSharedDefaultFolder _
              (oloRecipient, olFolderCalendar)
          'Display it in a new Outlook Explorer window.
        oloSharedFolder.Display
    Else
        MsgBox "Unable to locate " & strRecipient & _
           "Try another name.", vbInformation
    End If
End Sub




Private Sub NewMailMessage()
      'Create a new mail message item.
    Set oloNewMail = oloAPP.CreateItem(olMailItem)
    With oloNewMail
           'Add the subject of the mail message.
        .Subject = ""
           'Create some body text.
        .Body = ""
           'Add a recipient and test to make sure that the
           'address is valid using the Resolve method.
        With .Recipients.Add("xxxx@yyyyyyy.com")
            .Type = olTo
        End With
        With .Attachments.Add _
            ("", olByReference)
            .DisplayName = ""
        End With
        'Send the mail message.
        .Send
    End With
End Sub



Private Sub mnuCopyPrintDesktop_Click()
    On Error GoTo Errhandler
    Set gcdg = frmFrame.CommonDialog1
    gcdg.DialogTitle = "Print"
    gcdg.CancelError = True
    gcdg.Flags = cdlPDReturnDC + cdlPDPageNums
    gcdg.ShowPrinter
    If Err = MSComDlg.cdlCancel Then
         Set gcdg = Nothing
         Exit Sub
    End If
    Me.PrintForm
    Set gcdg = Nothing
    Exit Sub
Errhandler:
    If Err <> 32755 Then
         ErrMsgProc "mnuCopyPrintDesktop_Click"
    End If
End Sub


' Save current screen to file directly
Private Sub mnuCopyCapDeskTop_Click()
    On Error GoTo Errhandler
    Dim mfilespec As String
    Set gcdg = frmFrame.CommonDialog1
    gcdg.FileName = mfilespec
    gcdg.Flags = cdlOFNOverwritePrompt
    gcdg.ShowSave
    If gcdg.FileName = "" Then
        Set gcdg = Nothing
        Exit Sub
    End If
    mfilespec = gcdg.FileName
    Clipboard.Clear
    keybd_event VK_MENU, 0, 0, 0          ' Plant "Alt" key
    DoEvents
    keybd_event VK_SNAPSHOT, 1, 0, 0      ' Plant "Print Screen" key
    DoEvents
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0     ' Release "Alt" key
    DoEvents
      ' (Image is now in clipboard) print from clipboard to file
    SavePicture Clipboard.GetData(0), mfilespec
    Clipboard.Clear
    Set gcdg = Nothing
    Exit Sub
    
Errhandler:
    If Err <> 32755 Then
         ErrMsgProc "mnuCopyCapDeskTop_Click"
    End If
End Sub



' Save desktop in clipboard to file
' Sometimes you cannot invoke Save menu when a current screen
' is active, in this case you need to Alt+PrintScreen to capture
' that screen to clipboard first.  When menu is visible again,
' use this function to save stored image to file.
Private Sub mnuCopySaveDesktopFromClip_Click()
    On Error GoTo Errhandler
    If Not (Clipboard.GetFormat(vbCFBitmap) Or Clipboard.GetFormat(vbCFMetafile) Or _
       Clipboard.GetFormat(vbCFDIB) Or Clipboard.GetFormat(vbCFPalette)) Then
         MsgBox "No picture in clipboard yet"
         Exit Sub
    End If
    Dim mfilespec As String
    Set gcdg = frmFrame.CommonDialog1
    gcdg.FileName = mfilespec
    gcdg.Flags = cdlOFNOverwritePrompt
    gcdg.ShowSave
    If gcdg.FileName = "" Then
         Set gcdg = Nothing
         Exit Sub
    End If
    mfilespec = gcdg.FileName
    Clipboard.Clear
    keybd_event VK_MENU, 0, 0, 0          ' Plant "Alt" key
    DoEvents
    keybd_event VK_SNAPSHOT, 1, 0, 0      ' Plant "Print Screen" key
    DoEvents
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0     ' Release "Alt" key
    DoEvents
      ' (Image is now in clipboard) print from clipboard to file
    SavePicture Clipboard.GetData(0), mfilespec
    Clipboard.Clear
    Set gcdg = Nothing
    Exit Sub
Errhandler:
    If Err <> 32755 Then
         ErrMsgProc "mnuCopySaveDesktopFromClip_Click"
    End If
End Sub



Private Sub mnuCopySaveTextFromClip_Click()
    On Error Resume Next
    ctlWebBrowser.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DONTPROMPTUSER, _
        ctlWebBrowser.LocationName & ".txt"
End Sub



Private Sub cmdSelectFavorite_Click()
    If cboFavoritesList.ListCount = 0 Then
        MsgBox "No entry in Favorites List yet"
        Exit Sub
    End If
    cboAddress.Text = cboFavoritesList.Text
    cboAddress.SetFocus
End Sub




Private Sub cmdAddFavorite_Click()
    Screen.MousePointer = vbHourglass
    Dim mHandle As Variant
    Dim arrLines() As String
    Text1.Text = ""
    mHandle = FreeFile
    Open FavoriteFileName For Input As #mHandle
    Text1 = StrConv(InputB(LOF(mHandle), 1), vbUnicode)
    If Text1.Text <> "" Then
        If cboFavoritesList.ListCount > 0 Then
             If LineThere(cboAddress.Text) = True Then
                  Close #mHandle
                  Screen.MousePointer = vbDefault
                  MsgBox "Favorites List already has this entry"
                  Exit Sub
             End If
        End If
        Text1.Text = Text1.Text & vbCrLf & LTrim(Trim(cboAddress.Text))
    Else
        Text1.Text = LTrim(Trim(cboAddress.Text))
    End If
    Close #mHandle
    Me.Text1.SaveFile FavoriteFileName, 1
    Me.Text1.Text = ""
    SortFileLines
    FillcboFavoritesList
    ctlWebBrowser.SetFocus
    Screen.MousePointer = vbDefault
End Sub



Private Sub cmdDeleteFavorite_Click()
    If cboFavoritesList.ListCount = 0 Then
         MsgBox "No entry in Favorites List yet"
         Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    Dim mHandle
    Dim h, i
    Dim tmp
    i = cboFavoritesList.ListIndex
    cboFavoritesList.RemoveItem i
    DoEvents
      ' Copy combo items to file.
    mHandle = FreeFile
    Open FavoriteFileName For Output As #mHandle
    For i = 0 To cboFavoritesList.ListCount - 1
         tmp = cboFavoritesList.List(i)
         Print #mHandle, tmp
         If i < (cboFavoritesList.ListCount - 1) Then
              Print #mHandle, Chr(13)
         End If
    Next i
    Close #mHandle
    If cboFavoritesList.ListCount > 0 Then
         cboFavoritesList.ListIndex = 0
    End If
    ctlWebBrowser.SetFocus
    Screen.MousePointer = vbDefault
End Sub



Private Sub cmdExitFavoritesList_Click()
    picLinksListContainer.Visible = True
    ctlWebBrowser.SetFocus
End Sub



Private Function LineThere(inText) As Boolean
    Dim strLine As String
    Dim h, i
    h = cboFavoritesList.ListIndex
    LineThere = False
    For i = 0 To cboFavoritesList.ListCount - 1
        strLine = LTrim(Trim(cboFavoritesList.List(i)))
        If strLine = inText Then
             LineThere = True
             Exit For
        End If
    Next i
    cboFavoritesList.ListIndex = h
End Function




Private Sub SortFileLines()
    Dim strLine As String
    Dim intLineNums As Integer
    Dim arrLines() As String
    Dim mHandle As Variant
    mHandle = FreeFile
      ' We know file is in existence
    Open FavoriteFileName For Input As #mHandle
    Do While Not EOF(mHandle)
        Line Input #mHandle, strLine
        strLine = LTrim(Trim(strLine))
        intLineNums = intLineNums + 1
        ReDim Preserve arrLines(1 To intLineNums)
        arrLines(intLineNums) = strLine
    Loop
    Close #mHandle
    SelectionSort arrLines, 1, intLineNums
       ' Copy back to original file
       ' We use Open for "Output" to truncate the file first
       ' before writing new lines.  File is small in this case.
    mHandle = FreeFile
    Open FavoriteFileName For Output As #mHandle
    Dim i
    For i = 1 To intLineNums
         If Len(Trim(arrLines(i))) > 0 Then
             Print #mHandle, arrLines(i)
         End If
    Next i
    Close #mHandle
End Sub




Private Sub SelectionSort(inSortList() As String, ByVal inStart As Integer, ByVal inEnd As Integer)
    Dim i, j, intSelect
    Dim strSelect As String, strTemp As String
    For i = inStart To (inEnd - 1)
        intSelect = i
        strSelect = inSortList(i)
        For j = i + 1 To inEnd
            If StrComp(inSortList(j), strSelect, vbTextCompare) < 0 Then
                strSelect = inSortList(j)
                intSelect = j
            End If
        Next j
        inSortList(intSelect) = inSortList(i)
        inSortList(i) = strSelect
    Next i
End Sub



Sub OptionsStatusbarProc(CurrentForm As Form)
    CurrentForm.mnuOptionsStatusbar.Checked = Not CurrentForm.mnuOptionsStatusbar.Checked
    If CurrentForm.mnuOptionsStatusbar.Checked Then
        frmFrame.StatusBar1.Visible = True
    Else
        frmFrame.StatusBar1.Visible = False
    End If
End Sub



Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub

