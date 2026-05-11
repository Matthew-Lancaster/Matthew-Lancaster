VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Form3_MISC_FORM 
   Caption         =   "Form3"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3000
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   645
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
End
Attribute VB_Name = "Form3_MISC_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
