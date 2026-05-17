VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "GeoFacts"
   ClientHeight    =   4440
   ClientLeft      =   1410
   ClientTop       =   3165
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4440
   ScaleWidth      =   10995
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Form_Load()
    Setup   ' Sets the object variables to the Excel objects.
    FillContinentsList ' Fills the continents.
End Sub

Sub Form_Unload(Cancel As Integer)
    CleanUp ' Sets object variables to Nothing.
End Sub

Private Sub cmbContinents_Change()
    FillFeaturesList    ' Fill the Features ComboBox.
End Sub

Private Sub cmbFeatures_Change()
    FillTopRankingList ' Fill the TopRanking ListBox.
End Sub

Sub cmbFeatures_Click()
    FillTopRankingList ' Fill the TopRanking ListBox.
End Sub

Sub cmbContinents_Click()
    FillFeaturesList    ' Fill the Features ComboBox.
End Sub

