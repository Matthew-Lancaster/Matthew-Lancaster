Attribute VB_Name = "ReSize2"

Public Sub PicShow(ByVal PixPath As String, fForm As Form)
    On Error GoTo noshow
    Dim dHeight, dIHeight
    Dim dWidth, dIWidth
    Dim dPercent
    
        'If fForm.ViewImage.Picture = 0 Then fForm.ViewImage.Picture = LoadPicture(PixPath)
        fForm.ViewImage.Picture = LoadPicture(PixPath)


        fForm.Height = fForm.ViewImage.Height + 700
        fForm.Width = fForm.ViewImage.Width + 150
        fForm.ViewImage.Height = fForm.Height - 700
        fForm.ViewImage.Width = fForm.Width - 100
        fForm.ViewImage.Top = 0
        fForm.ViewImage.Left = 0
        '.PicBack
        fForm.PicBack.Height = fForm.ViewImage.Height
        fForm.PicBack.Width = fForm.ViewImage.Width
        fForm.PicBack.Top = 0
        fForm.PicBack.Left = 0
   
    
    With fForm
        .ViewImage.Visible = False
        .ViewImage.Stretch = False
        .Caption = App.Title & " - " & UCase(PixPath)
 
        'Bad Code
        '.ViewImage.Picture = PowerResize.PowerResize(.ViewImage.Picture, .PicBack.Height, .PicBack.Width)

'GoTo jump:

        If .ViewImage.Height < .PicBack.Height And .ViewImage.Width < .PicBack.Width Then
            .ViewImage.Visible = True
            Exit Sub
        End If
        dHeight = .ViewImage.Height
        dWidth = .ViewImage.Width
        dIHeight = .PicBack.Height - 1
        dIWidth = .PicBack.Width - 1
        .ViewImage.Stretch = True
        .ViewImage.Height = .PicBack.Height - 2
        dPercent = (.PicBack.Height - 2) / dHeight * 100
        .ViewImage.Width = dWidth / 100 * dPercent


        If .ViewImage.Width > (.PicBack.Width - 2) Then
            .ViewImage.Stretch = False
            dHeight = .ViewImage.Height
            dWidth = .ViewImage.Width
            dIHeight = .PicBack.Height - 1
            dIWidth = .PicBack.Width - 1
            .ViewImage.Stretch = True
            .ViewImage.Width = .PicBack.Width - 1
            dPercent = (.PicBack.Width - 1) / dWidth * 100
            .ViewImage.Height = dHeight / 100 * dPercent
        End If
        
        
        .ViewImage.Visible = True
        'MidPic frmMain2000
        MidPic PicShowReSize
    End With
    Exit Sub
noshow:
    Resume noshow1
noshow1:
End Sub

Public Sub MidPic(ByVal fForm As Form)
    fForm.ViewImage.Move (fForm.PicBack.Width - fForm.ViewImage.Width) / 2, (fForm.ViewImage.Height - fForm.ViewImage.Height) / 2
End Sub


