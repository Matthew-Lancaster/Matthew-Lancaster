VERSION 5.00
Begin VB.Form WebPhoto 
   Caption         =   "Web photo's"
   ClientHeight    =   3465
   ClientLeft      =   8655
   ClientTop       =   1275
   ClientWidth     =   9630
   Icon            =   "Web_Photos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   9630
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1365
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   90
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   690
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Web_Photos.frx":0442
      Top             =   90
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3390
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "Web_Photos.frx":044A
      Top             =   90
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   630
      Left            =   8745
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   8
      Top             =   300
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2745
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   90
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox Statctext2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Web_Photos.frx":0452
      Top             =   90
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2025
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Web_Photos.frx":045F
      Top             =   90
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   630
      Top             =   1395
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Hidden          =   -1  'True
      Left            =   0
      Pattern         =   "*.jpg"
      TabIndex        =   2
      Top             =   3510
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9615
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   3555
      TabIndex        =   0
      Top             =   3510
      Visible         =   0   'False
      Width           =   5505
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   3015
      Width           =   1095
   End
End
Attribute VB_Name = "WebPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ti2 As Date
Dim store5$()
Public storec As Integer


Private Sub Dir1_Change()
File1.path = Dir1.path

End Sub

Private Sub Form_Activate()
WebPhotComplete = False
ti2 = Now + TimeSerial(0, 10, 0)
Call subr1
ti2 = Now + TimeSerial(0, 0, 6)
WebPhotComplete = True


Webphot = 1
Main.Command1.Enabled = False
Main.Command2.Enabled = True
Main.Command3.Enabled = True
Main.Check1.Enabled = True
Main.Command4.Enabled = True

Main.Timer1.Enabled = False

End Sub

Private Sub Form_Load()

Text3.Text = "<a href=""photos.html""><font face=""Comic Sans MS""><b>Back to Photo / Image Menu"
Text3.Text = Text3.Text + "Page</b></font></a>" + vbCr + "<br>"

Text4.Text = "<body bgcolor=""#000066"" link=""#00FF00"" vlink=""#FFFFFF"" alink=""#FFFF00"" text=""#3399FF"" style=""font-family: Arial"">"
Text6.Text = "<a href=""http://matthewlan.com""><font face=""Comic Sans MS""><b>Back to Home Page</b></font></a>" + vbCr + "<br>"


'<META NAME="ROBOTS" CONTENT="INDEX,FOLLOW">
qc1$ = "<html>" + vbCrLf
qc1$ = qc1$ + "<head>" + vbCrLf
qc1$ = qc1$ + "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">" + vbCrLf
qc1$ = qc1$ + "<meta name=""GENERATOR"" content=""Microsoft FrontPage 4.0"">" + vbCrLf
qc1$ = qc1$ + "<meta name=""ProgId"" content=""FrontPage.Editor.Document"">" + vbCrLf
qc1$ = qc1$ + "<META name=""Description"" content=""My Photo/Image Album - Sub Folder - MatthewLan.com"">" + vbCrLf
qc1$ = qc1$ + "<META name=""keywords"" content=""matthewlan.com, Photo,Photos, Image, Images, Album, Graphics"">" + vbCrLf
qc1$ = qc1$ + "<META NAME=""ROBOTS"" CONTENT=""INDEX,NOFOLLOW"">" + vbCrLf
qc1$ = qc1$ + "<style fprolloverstyle>A:hover {color: #00FF00; font-weight: bold}" + vbCrLf + vbCrLf
qc1$ = qc1$ + "</style>" + vbCrLf

Text1.Text = qc1$

qc1$ = Statc$


Statctext2.Text = qc1$


'Text2.Text = "<html>" + vbCr
'Text2.Text = Text2.Text + "<head>" + vbCr
'Text2.Text = Text2.Text + "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">" + vbCr
'Text2.Text = Text2.Text + "<meta name=""GENERATOR"" content=""Microsoft FrontPage 4.0"">" + vbCr
'Text2.Text = Text2.Text + "<meta name=""ProgId"" content=""FrontPage.Editor.Document"">" + vbCr
'Text2.Text = Text2.Text + "<style fprolloverstyle>A:hover {color: #00FF00; font-weight: bold}" + vbCr
'Text2.Text = Text2.Text + "</style>" + vbCr





WebPhoto.Show

'D:\My Webs\Matt Lan Web3\Photos
Dir1.path = Mdu$ + "My Webs\MatthewLan.Com Web\Photos\Photos_Jpgs"
'Dir1.Path = "D:\Documents and Settings\All Users\Documents\DSSPlayer\Message\FolderA\2003-4 004"


ReDim store5$(560)

freef5 = FreeFile
Open App.path + "\WebPhotos\photo_template.html" For Input As #freef5
For r = 1 To 160
    Line Input #freef5, store5$(r)
    If EOF(freef5) Then Exit For
Next
Close #freef5
storec = r
End Sub



Sub subr1()
Dim store$(15)
dit$ = """"

fontcol2$ = "<font color=""#FFFFFF"">"

'For rtg = Dir1.ListIndex + 1 To Dir1.ListCount
'a$ = Dir1.List(rtg)
'If InStr(a$, "A1_Menu") Then Exit For
'Next

For r = 0 To Dir1.ListCount - 1

a$ = Dir1.List(r)


File1.path = Dir1.List(r)



i5 = 0
For rt = Len(a$) To 1 Step -1
If Mid$(a$, rt, 1) = "\" Then i5 = i5 + 1
If i5 = 2 Then Exit For
Next
b3$ = Mid$(a$, rt + 1)
b4$ = Mid$(a$, 1, rt)
'b5$ = b3$
b5$ = Mid$(b3$, 8)
b6$ = "Photos\A1Menu\Photos_HTML\" + b5$
For rt = Len(b3$) To 1 Step -1
If Mid$(b3$, rt, 1) = "\" Then Mid$(b3$, rt, 1) = "_"
If Mid$(b5$, rt, 1) = "\" Then Mid$(b5$, rt, 1) = "/"
Next
'b3$ = LCase$(b3$)

'Open b4$ + b3$ + ".html" For Output As #1
Open App.path + "\WebPhotos\temp.html" For Output As #1
Print #1, Text1.Text

d6$ = b3$
For rt = 1 To Len(d6$)
If Mid$(d6$, rt, 1) = "_" Then Mid$(d6$, rt, 1) = " "
Next



Print #1, "<title>Photo Image's of :- " + Mid$(d6$, 8) + "</title>"
Print #1, "<META name=""description"" content=""Matthew Lan's Images"">"
Print #1, "<META name=""keywords"" content=""fractal, mathmatical images, funny, paradise, my album, cloud photos, equinox"">"
'Print #1, "<META NAME=""ROBOTS"" CONTENT=""NOINDEX,NOFOLLOW"">"
'<META name=""description"" content=""Matthew Lancaster' Images"">
'<META name=""keywords"" content=""fractal, mathmatical images, funny, paradise, my album, cloud photos, equinox"">




Print #1, "</head>"
Print #1,
asd$ = Text4.Text

Print #1, asd$




'<p><b><font face="Arial" size="5">-&lt;&lt;&lt;Photo's Of Clouds&gt;&gt;&gt;-</font></b></p>
Print #1, "<b><font face=" + dit$ + "Arial" + dit$ + " size=" + dit$ + "5" + dit$ + ">--- Photo's / Image's Of --- <br>" + "--- " + Mid$(d6$, 8) + " ---" + fontcol2$ + Str$((File1.ListCount)) + "</font> Image's ---</font></b>"



Print #1, "<br>"
Print #1, Text3.Text

Print #1,
Print #1, "<div align=" + dit$ + "left" + dit$ + ">"
Print #1, "<table border=" + dit$ + "8" + dit$ + " width=" + dit$ + "1" + dit$ + " height=" + dit$ + "1" + dit$ + ">"

Print #1, "<tr>"



xc = 0
crumb = 0
crumb2 = 8
filecount2 = 0
For t = File1.ListIndex + 1 To File1.ListCount - 1
'SSA = InStr(UCase$(File1.List(t)), ".JPG")
'b$ = a$ + "\" + Mid$(File1.List(t), 1, SSA - 1) + ".WAV"
d1$ = File1.List(t)
d2$ = d1$

'For rt = Len(d1$) To 1 Step -1
'If Mid$(d1$, rt, 1) = " " Then Mid$(d1$, rt, 1) = "_"
'Next
'If d1$ <> d2$ Then
'Name a$ + "\" + d2$ As a$ + "\" + d1$
'End If

b$ = a$ + "\" + d1$


i5 = 0
For rt = Len(b$) To 1 Step -1
If Mid$(b$, rt, 1) = "\" Then i5 = i5 + 1
If i5 = 3 Then Exit For
Next
b$ = Mid$(b$, rt)

'Name b4$ + b$ As LCase$(b4$ + b$)

'b$ = LCase$(b$)

tic1$ = Mid$(b$, InStr(2, b$, "\"))
tic2$ = "Photos" + Mid$(tic1$, 1, Len(tic1$) - 3) + "htm"
tic3$ = Mid$(tic1$, 2, InStr(2, tic1$, "\") - 2)

If tic3$ <> tic4$ Then
On Local Error Resume Next
MkDir b4$ + "A1_Menu\Photos_HTML\" + tic3$
'Kill b4$ + "Photos\" + tic3$ + "\_vti_cnf\*.html"
'RmDir b4$ + "Photos\" + tic3$ + "\_vti_cnf"
On Local Error GoTo 0
End If
tic4$ = tic3$

If InStr(b$, "tn_") = False Then
filecount2 = filecount2 + 1
d6$ = d1$
For rt = 1 To Len(d6$)

'If Mid$(d1$, rt, 1) = "_" Then Mid$(d6$, rt, 1) = " "

Next

d6$ = Mid$(d6$, 1, Len(d6$) - 4)




freef5 = FreeFile
'Open b4$ + tic2$ For Output As #freef5


Open App.path + "\WebPhotos\temp2.html" For Output As #freef5

      
      Set Picture1.Picture = LoadPicture(b4$ + b$)
      'If InStr(b$, "TSHIRT04.jpg") Then Stop
              
      ax = Picture1.ScaleWidth
      ay = Picture1.ScaleHeight
      
      ax2$ = Trim$(Round(ax, 0))
      ay2$ = Trim$(Round(ay, 0))
      
      ratio = 1 + (1 / 2000)
      compsize = 1
      Do
      Tag = 0
      If ax > 640 Or ay > 480 Then
      ax = ax / ratio: ay = ay / ratio: Tag = 1
      compsize = compsize / ratio
      End If
      
      If ax < 639.6 And ay < 479.6 Then
      ax = ax * ratio: ay = ay * ratio: Tag = 1
      compsize = compsize * ratio
      End If
      
      DoEvents
      Loop Until Tag = 0
      
      ax1$ = Trim$(Round(ax, 0))
      ay1$ = Trim$(Round(ay, 0))
      
      comps$ = Format$(compsize, "##0.0000")
    WebPhoto.Refresh

DoEvents

List1.AddItem ax1$ + "x" + ay1$ + "    " + b$

Label1.Caption = Str$(List1.ListCount)

List1.ListIndex = List1.ListCount - 1

xe1 = 0
font2 = 0


For ri8 = 1 To storec
slab$ = store5$(ri8)

If store5$(ri8) = "<title>" Then
slab$ = "<title>Photo Image of :- " + d1$ + "</title>"
ri8 = ri8 + 1
End If


If Mid$(store5$(ri8), 1, 8) = "<body bg" Then
asd$ = store5$(ri8)
asf$ = Mid$(asd$, 21, 2)
ash$ = "&h" + asf$
ascol = Val(ash$)
If crumb >= 80 And crumb2 = 8 Then crumb2 = -8
If crumb <= -80 And crumb2 = -8 Then crumb2 = 8
crumb = crumb + crumb2
ascol2 = ascol + crumb
Mid$(asd$, 21, 2) = Hex$(ascol2)
slab$ = asd$
End If



If Mid$(store5$(ri8), 1, 10) = "<!--title1" Then
tic31$ = tic3$
For ri5 = 1 To Len(tic3$)
If Mid$(tic3$, ri5, 1) = "_" Then Mid$(tic31$, ri5, 1) = " "
Next
d11$ = d1$
For ri5 = 1 To Len(d1$)
If Mid$(d1$, ri5, 1) = "_" Then Mid$(d11$, ri5, 1) = " "
Next
slab$ = store5$(ri8) + fontcol2$ + tic31$ + "</font> -- " + fontcol2$ + Mid$(d11$, 1, Len(d11$) - 4) + "</font>.jpg" + store5$(ri8 + 2)
ri8 = ri8 + 2
End If

If Mid$(store5$(ri8), 1, 10) = "<!--title2" Then

comrat$ = "Compression Ratio "
If Val(comps$) > 1 Then comrat$ = "DeCompression Ratio "
If Val(comps$) = 1 Then
comrat$ = "No Compression - Ratio "
End If
slab$ = store5$(ri8) + "Image Size " + fontcol2$ + ax1$ + "x" + ay1$ + "</font> Originaly " + fontcol2$ + ax2$ + "x" + ay2$ + "</font> " + comrat$ + fontcol2$ + comps$ + "</font>" + store5$(ri8 + 2)
ri8 = ri8 + 2
End If

If Mid$(store5$(ri8), 1, 10) = "<!--title3" Then
slab$ = store5$(ri8) + "Image - " + fontcol2$ + Str$(filecount2) + "</font> Of " + fontcol2$ + Str$((File1.ListCount)) + "</font>." + store5$(ri8 + 2)
ri8 = ri8 + 2
End If

'----Source for Main Picture Image source in the html
If Mid$(store5$(ri8), 1, 8) = "<img bor" Then
slab$ = store5$(ri8) + "../../../Photos_Jpgs/" + tic3$ + "/" + d1$ + store5$(ri8 + 2) + " width=""" + ax1$ + """ height=""" + ay1$ + """>"
ri8 = ri8 + 3
End If

If Mid$(store5$(ri8), 1, 8) = "<a href=" And xe1 = 0 Then
xe1 = 1
'back to thumb menu
slab$ = store5$(ri8) + "../../Photos_" + tic3$ + ".html" + store5$(ri8 + 2)
backtolastpage$ = slab$
ri8 = ri8 + 2
End If



'Next and previous photo image
If Mid$(store5$(ri8), 1, 8) = "<a href=" And xe1 = 1 Then
'Next Photo Image-------
xe1 = 2
d1a$ = ""
For tri = 1 To 3000
If t + tri <= File1.ListCount - 1 Then
If Mid$(File1.List(t + tri), 1, 3) <> "tn_" Then d1a$ = File1.List(t + tri)
'If InStr(d1a$, "TSHIRT04") Then Stop
End If
If d1a$ <> "" Or t + tri > File1.ListCount - 1 Then Exit For
Next

If d1a$ <> "" Then
'    b7$ = "../" + tic3$ + "/" + Mid$(d1a$, 1, Len(d1a$) - 4) + ".html"
    b7$ = Mid$(d1a$, 1, Len(d1a$) - 4) + ".html"
End If
slab$ = store5$(ri8) + b7$ + store5$(ri8 + 2)
If d1a$ = "" Then
slab$ = backtolastpage$
End If
ri8 = ri8 + 2
End If



'previous---------
If Mid$(store5$(ri8), 1, 8) = "<a href=" And xe1 = 2 Then
xe1 = 3
d1a$ = ""
For tri = 1 To 3000
If t - tri <= File1.ListCount - 1 Then
If Mid$(File1.List(t - tri), 1, 3) <> "tn_" Then d1a$ = File1.List(t - tri)
End If
If d1a$ <> "" Or t - tri < 0 Then Exit For
Next

If d1a$ <> "" Then
    b7$ = "../" + tic3$ + "/" + Mid$(d1a$, 1, Len(d1a$) - 4) + ".html"
End If

slab$ = store5$(ri8) + b7$ + store5$(ri8 + 2)
If d1a$ = "" Then
slab$ = ""
End If
ri8 = ri8 + 2
End If


'---Link for main Image
If Mid$(store5$(ri8), 1, 8) = "<a href=" And xe1 = 3 Then
xe1 = 4
slab$ = store5$(ri8) + "../../../Photos_Jpgs/" + tic3$ + "/" + d1$ + store5$(ri8 + 2)
ri8 = ri8 + 2
End If



Print #freef5, slab$

Next

Close #freef5


Compare = 1
If Dir$(b4$ + tic2$) <> "" Then
xf1 = FreeFile
Open b4$ + tic2$ For Input As #xf1
xf2 = FreeFile
Open App.path + "\WebPhotos\temp2.html" For Input As #xf2
Compare = 0
Do
Line Input #xf1, comp1$
Line Input #xf2, comp2$
If comp1$ <> comp2$ Then Compare = 1: Exit Do
Loop Until EOF(xf1)
Close #xf1, #xf2
End If

If Compare = Compare Then '=1
xf1 = FreeFile
tic2$ = Mid$(tic2$, 8)
tic2$ = "A1_Menu\Photos_HTML\" + tic2$

On Local Error Resume Next
Open b4$ + tic2$ For Output As #xf1
If Err.Description = "Path not found" Then
MkDir b4$ + "A1_Menu\Photos_HTML"
MkDir b4$ + "A1_Menu\Photos_HTML\" + tic3$
End If
Open b4$ + tic2$ For Output As #xf1

On Local Error GoTo 0



xf2 = FreeFile
Open App.path + "\WebPhotos\temp2.html" For Input As #xf2
Do
Line Input #xf2, comp1$
Print #xf1, comp1$
Loop Until EOF(xf2)
Close #xf1, #xf2
End If





'b7$ = "Photos/" + tic3$ + "/" + Mid$(d1$, 1, Len(d1$) - 4) + ".html"
'b7$ = "../" + tic3$ + "/" + Mid$(d1$, 1, Len(d1$) - 4) + ".html"


'from the thumb menu to the picture
b7$ = "Photos_HTML/" + tic3$ + "/" + Mid$(d1$, 1, Len(d1$) - 4) + ".html"

stor5$ = "<A HREF=""" + b7$ + """" + "><IMG SRC=""" + "../Photos_Jpgs/" + tic3$ + "/001-Thumbs/tn_" + d1$ + """" + " border 0 ></A>"

Print #1, "<td align=""" + "center" + """" + ">"
Print #1, stor5$
Print #1, "</td>"
store$(xc) = d6$
xc = xc + 1
If xc = 3 Then
xc = 0
Print #1, "</tr>"
Print #1, "<tr>"
For rt4 = 0 To 2
Print #1, "<td align=""center"" valign+""top"">" + store$(rt4) + "</td>"
Next
If t <> File1.ListCount - 1 Then
Print #1, "</tr>"
Print #1, "<tr>"
End If

store$(0) = ""
store$(1) = ""
store$(2) = ""
store$(3) = ""
End If

End If
Next

If xc > 0 Then
xc = 0
Print #1, "</tr>"
Print #1, "<tr>"
For rt4 = 0 To 2
Print #1, "<td align=" + dit$ + "center" + dit$ + " valign+" + dit$ + "top" + dit$ + ">" + store$(rt4) + "</td>"
Next
Print #1, "</tr>"
Print #1, "<tr>"

End If

Print #1, "</tr>"
Print #1, "</table>"
Print #1, "</div>"

Print #1, "<br>"
Print #1, Text3.Text
Print #1, Text6.Text

Print #1, Statctext2.Text






Print #1, "</body>"
Print #1, "</html>"

Close #1

Compare = 1

cool2$ = b4$ + "A1_Menu\Photos_" + tic3$ + ".html"
If Dir$(cool2$) <> "" Then
Open cool2$ For Input As #1
Open App.path + "\WebPhotos\temp.html" For Input As #2
Compare = 0
Do
Line Input #1, comp1$
Line Input #2, comp2$
If comp1$ <> comp2$ Then Compare = 1: Exit Do
Loop Until EOF(1)
Close #1, #2
End If

If Compare = Compare Then
Open cool2$ For Output As #1
Open App.path + "\WebPhotos\temp.html" For Input As #2
Do
Line Input #2, comp1$
Print #1, comp1$
Loop Until EOF(2)
Close #1, #2
End If


Next

Label1.Caption = Str$(List1.ListCount)
Dim quty$(2000)

Open b4$ + "\A1_Menu\Photos.html" For Input As #1
'Open App.Path + "\temp3.html" For Input As #1
we = 0
Do
we = we + 1
Line Input #1, quty$(we)
Loop Until EOF(1)
Close #1
Open App.path + "\WebPhotos\temp3.html" For Output As #1
For r = 1 To we
slab$ = quty$(r)
If InStr(quty$(r), "Total Images =") Then
slab$ = "Total Images =" + fontcol2$ + Label1.Caption + "</font>"
End If
Print #1, slab$
Next
Close #1



xf1 = FreeFile
Open b4$ + "A1_Menu\" + "Photos.html" For Input As #xf1
xf2 = FreeFile
Open App.path + "\WebPhotos\temp3.html" For Input As #xf2
Compare = 0
Do
Line Input #xf1, comp1$
Line Input #xf2, comp2$
If comp1$ <> comp2$ Then Compare = 1: Exit Do
Loop Until EOF(xf1)
Close #xf1, #xf2
If Compare = Compare Then '=1
xf1 = FreeFile
Open b4$ + "A1_Menu\Photos.html" For Output As #xf1
xf2 = FreeFile
Open App.path + "\WebPhotos\temp3.html" For Input As #xf2
Do
Line Input #xf2, comp1$
'If InStr(comp1$, "TSHIRT04") Then Stop
Print #xf1, comp1$
Loop Until EOF(xf2)
Close #xf1, #xf2
End If



End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
'End
End Sub



Private Sub Timer1_Timer()
If ti2 < Now Then
Unload WebPhoto
End If
End Sub
