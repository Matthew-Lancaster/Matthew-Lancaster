Attribute VB_Name = "aModule1"
'Private Sub Form_Load()
'Dim lPic As Picture
'Me.Picture1.AutoRedraw = True
'Set lPic = LoadPicture("C:\YourPicture.jpg") 'Use the correct path and filename here
'ResizePicture Me.Picture1, lPic
'End Sub

Public FS

Public Sub ResizePicture(pBox As PictureBox, pPic As Picture)
Dim lWidth As Single, lHeight As Single
Dim lnewWidth As Single, lnewHeight As Single
'Clear the Picture in the PictureBox
pBox.Picture = Nothing
'Clear the Image in the Picturebox
pBox.Cls
'Get the size of the Image, but in the same Scale than the scale used by the PictureBox
lWidth = pBox.ScaleX(pPic.Width, vbHiMetric, pBox.ScaleMode)
lHeight = pBox.ScaleY(pPic.Height, vbHiMetric, pBox.ScaleMode)
'Old Rems
'If image Width > pictureBox Width, resize Width
If lWidth > pBox.ScaleWidth Then
    lnewWidth = pBox.ScaleWidth 'new Width = PB width
    lHeight = lHeight * (lnewWidth / lWidth) 'Risize Height keeping proportions
Else
    'if you want true size on small Pics then This
        'lnewWidth = lWidth 'If not, keep the original Width value

    'OR if you want expand small and shrink big then this
    lnewWidth = pBox.ScaleWidth 'new Width = PB width
    lHeight = lHeight / (lWidth / lnewWidth) 'Risize Height keeping proportions

End If
    'Old Rems
    'If the image Height > The pictureBox Height, resize Height

If lHeight > pBox.ScaleHeight Then
    lnewHeight = pBox.ScaleHeight 'new Height = PB Height
    lnewWidth = lnewWidth * (lnewHeight / lHeight) 'Risize Width keeping proportions
Else
    'if you want true size on small then This only this line
        lnewHeight = lHeight 'If not, use the same value

    'OR if you want expand small and shrink big then this
    lnewWidth = lnewWidth / (lHeight / lnewHeight) 'Risize Width keeping proportions
End If

'add resized and centered to Picturebox
pBox.PaintPicture pPic, (pBox.ScaleWidth - lnewWidth) / 2, _
(pBox.ScaleHeight - lnewHeight) / 2, _
lnewWidth, lnewHeight

'Update the Picture with the new image if you need it
'Set pBox.Picture = pBox.Image
End Sub

