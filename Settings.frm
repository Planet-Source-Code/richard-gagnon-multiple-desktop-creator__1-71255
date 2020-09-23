VERSION 5.00
Begin VB.Form Settings 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   2640
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5805
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   387
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Bkgnd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   1740
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   152
      TabIndex        =   10
      Top             =   340
      Width           =   2285
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SAMPLE"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   735
         TabIndex        =   11
         Top             =   735
         Width           =   960
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tiled"
      Height          =   225
      Index           =   2
      Left            =   210
      TabIndex        =   9
      Top             =   1260
      Width           =   1170
   End
   Begin VB.PictureBox PicSize 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   315
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   8
      Top             =   2940
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CommandButton ACKbuttons 
      Caption         =   "Apply"
      Height          =   435
      Index           =   3
      Left            =   4410
      TabIndex        =   7
      Top             =   2100
      Width           =   1275
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Centered"
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   6
      Top             =   945
      Width           =   1170
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stretched"
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   5
      Top             =   630
      Value           =   -1  'True
      Width           =   1170
   End
   Begin VB.CommandButton Fcolor 
      Caption         =   "Font Color"
      Height          =   435
      Left            =   105
      TabIndex        =   4
      ToolTipText     =   "Font Color"
      Top             =   2100
      Width           =   1275
   End
   Begin VB.CommandButton BGcolor 
      Caption         =   "Back Color"
      Height          =   435
      Left            =   105
      TabIndex        =   3
      ToolTipText     =   "Background Color"
      Top             =   1575
      Width           =   1275
   End
   Begin VB.CommandButton BGImage 
      Caption         =   "Back Image"
      Height          =   435
      Left            =   105
      TabIndex        =   2
      ToolTipText     =   "Background Image"
      Top             =   105
      Width           =   1275
   End
   Begin VB.CommandButton ACKbuttons 
      Caption         =   "Cancel"
      Height          =   435
      Index           =   1
      Left            =   4410
      TabIndex        =   1
      Top             =   1050
      Width           =   1275
   End
   Begin VB.CommandButton ACKbuttons 
      Caption         =   "OK"
      Height          =   435
      Index           =   0
      Left            =   4410
      TabIndex        =   0
      Top             =   105
      Width           =   1275
   End
   Begin VB.Image Monitor 
      Height          =   2400
      Left            =   1575
      Picture         =   "Settings.frx":0ECA
      Stretch         =   -1  'True
      Top             =   105
      Width           =   2655
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private INI As New INIclass
Dim SelectedFileName As String
Dim SelectedBackgroundColor As Long
Dim SelectedFontColor As Long
Dim SelectedImageMode As Byte
'----------------------------
Dim DefaultFileName As String
Dim DefaultBackgroundColor As Long
Dim DefaultFontColor As Long
Dim DefaultImageMode As Byte
'----------------------------
Dim multX As Single, multY As Single
Private Sub ACKbuttons_Click(Index As Integer)
Dim I As Integer
VDT.Vform.Cls
Select Case Index
    Case 0  'OK
        INI.Path = UserPath & MySystem
        INI.Section = VDT.DesktopList.Text
        INI.Key = "BGimage"
        INI.Value = SelectedFileName
        INI.Key = "BGcolor"
        INI.Value = SelectedBackgroundColor
        INI.Key = "Fcolor"
        INI.Value = SelectedFontColor
        INI.Key = "Smode"
        INI.Value = SelectedImageMode
        VDT.FormSettings.Tag = SelectedImageMode
        VDT.Vform.Tag = SelectedFileName
        '--------------------------------------------
        VDT.Vform.BackColor = SelectedBackgroundColor
        If SelectedFileName <> "" Then ProcessFormImage SelectedImageMode, SelectedFileName
        For I = 1 To VDT.IconCaption.UBound
            VDT.IconCaption(I).ForeColor = SelectedFontColor
            VDT.Outline(I).BorderColor = SelectedFontColor
        Next I
        VDT.Iselect.BorderColor = SelectedFontColor
        VDT.Title.ForeColor = SelectedFontColor
        Unload Me
    Case 1  'Cancel
        VDT.Vform.BackColor = DefaultBackgroundColor
        If DefaultFileName <> "" Then ProcessFormImage DefaultImageMode, DefaultFileName
        For I = 1 To VDT.IconCaption.UBound
            VDT.IconCaption(I).ForeColor = DefaultFontColor
            VDT.Outline(I).BorderColor = DefaultFontColor
        Next I
        VDT.Iselect.BorderColor = DefaultFontColor
        Unload Me
    Case 3  'Apply
        VDT.Vform.BackColor = SelectedBackgroundColor
        If SelectedFileName <> "" Then ProcessFormImage SelectedImageMode, SelectedFileName
        For I = 1 To VDT.IconCaption.UBound
            VDT.IconCaption(I).ForeColor = SelectedFontColor
            VDT.Outline(I).BorderColor = SelectedFontColor
        Next I
        VDT.Iselect.BorderColor = SelectedFontColor
End Select

End Sub

Private Sub BGImage_Click()
Dim sOpen As SelectedFile
On Error GoTo OPNerr

FileDialog.sFilter = "All Image Types( *.bmp; *.jpg; *.gif )" & Chr$(0) & "*.BMP;*.JPG;*.GIF"
FileDialog.flags = OFN_FILEMUSTEXIST
FileDialog.sDlgTitle = "Wallpaper"
FileDialog.sInitDir = "C:\Windows"
sOpen = ShowOpen(Me.hwnd)
If sOpen.bCanceled Then GoTo OPNerr
If Err.Number <> 32755 Then
    SelectedFileName = sOpen.sLastDirectory & sOpen.sFiles(1)
    If SelectedFileName <> "" Then ProcessLocalImage SelectedFileName
End If
OPNerr:
End Sub

Private Sub BGcolor_Click()
Dim sColor As SelectedColor
On Error GoTo OPNerr

sColor = ShowColor(Me.hwnd)
If sColor.bCanceled Then GoTo OPNerr
SelectedBackgroundColor = sColor.oSelectedColor
Bkgnd.BackColor = sColor.oSelectedColor
If SelectedFileName <> "" Then ProcessLocalImage SelectedFileName
OPNerr:
End Sub

Private Sub Fcolor_Click()
Dim sColor As SelectedColor
On Error GoTo OPNerr

sColor = ShowColor(Me.hwnd)
If sColor.bCanceled Then GoTo OPNerr
SelectedFontColor = sColor.oSelectedColor
Label2.ForeColor = sColor.oSelectedColor
OPNerr:
End Sub

Private Sub Form_Load()
DefaultBackgroundColor = VDT.Vform.BackColor
DefaultFontColor = VDT.IconCaption(1).ForeColor
DefaultFileName = VDT.Vform.Tag
'--------------------------------
SelectedBackgroundColor = VDT.Vform.BackColor
SelectedFontColor = VDT.IconCaption(1).ForeColor
SelectedFileName = VDT.Vform.Tag
'--------------------------------
Bkgnd.BackColor = DefaultBackgroundColor
Label2.ForeColor = DefaultFontColor
'--------------------------------
SelectedImageMode = Val(VDT.FormSettings.Tag)
DefaultImageMode = Val(VDT.FormSettings.Tag)
Option1(DefaultImageMode).Value = True
'--------------------------------
If SelectedFileName <> "" Then ProcessLocalImage SelectedFileName
End Sub
Private Sub ProcessLocalImage(mFile)
On Error GoTo BGerr
Dim IMGW As Integer, IMGH As Single
Dim multX As Single, multY As Single
Dim X As Integer, Y As Integer
Dim I As Integer, J As Integer

Bkgnd.Cls: PicSize.Cls
PicSize.Picture = LoadPicture(mFile)
IMGW = PicSize.Width * (Bkgnd.Width / VDT.Vform.Width)
IMGH = PicSize.Height * (Bkgnd.Height / VDT.Vform.Height)
Select Case SelectedImageMode
    Case 0    'Stretched
        multX = Bkgnd.Width / IMGW
        multY = Bkgnd.Height / IMGH
        Bkgnd.PaintPicture PicSize.Image, 0, 0, IMGW * multX, IMGH * multY, 0, 0, PicSize.Width, PicSize.Height
    Case 1    'Centered
        multX = Bkgnd.Width / IMGW
        multY = Bkgnd.Height / IMGH
        If multX > 1 And multY > 1 Then
            multX = 1: multY = 1
        Else
            If multX > multY Then multX = multY Else multY = multX
        End If
        
        X = Bkgnd.Width / 2 - IMGW / 2 * multX
        Y = Bkgnd.Height / 2 - IMGH / 2 * multY
        Bkgnd.PaintPicture PicSize.Image, X, Y, IMGW * multX, IMGH * multY, 0, 0, PicSize.Width, PicSize.Height
    Case 2    'Tiled
        X = Bkgnd.Width / IMGW
        Y = Bkgnd.Height / IMGH
        For J = 0 To Y * IMGH Step IMGH
            For I = 0 To X * IMGW Step IMGW
               Bkgnd.PaintPicture PicSize.Image, I, J, IMGW, IMGH, 0, 0, PicSize.Width, PicSize.Height
            Next I
        Next J
End Select
Exit Sub
BGerr:
MsgBox "Unable to Locate Background Image file:" & vbCrLf & vbCrLf _
& SelectedFileName, vbExclamation, "File Not Found"
End Sub

Private Sub Bkgnd_Click()
RemoveBackgroundImage
End Sub

Private Sub Bkgnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DetectBackgroundImage Bkgnd
End Sub

Private Sub Label2_Click()
RemoveBackgroundImage
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DetectBackgroundImage Label2
End Sub

Private Sub Option1_Click(Index As Integer)
SelectedImageMode = Index
If SelectedFileName <> "" Then ProcessLocalImage SelectedFileName
End Sub
Private Sub ProcessFormImage(Smode As Byte, mFile As String)
On Error GoTo BGerr
Dim multX As Single, multY As Single
Dim X As Integer, Y As Integer
Dim I As Integer, J As Integer

PicSize.Picture = LoadPicture(mFile)
Select Case Smode
    Case 0    'Stretched
        multX = VDT.Vform.Width / PicSize.Width
        multY = VDT.Vform.Height / PicSize.Height
        VDT.Vform.PaintPicture PicSize.Image, 0, 0, PicSize.Width * multX, PicSize.Height * multY, 0, 0, PicSize.Width, PicSize.Height
    Case 1    'Centered
        multX = VDT.Vform.Width / PicSize.Width
        multY = VDT.Vform.Height / PicSize.Height
        If multX > 1 And multY > 1 Then
            multX = 1: multY = 1
        Else
            If multX > multY Then multX = multY Else multY = multX
        End If
        
        X = VDT.Vform.Width / 2 - PicSize.Width / 2 * multX
        Y = VDT.Vform.Height / 2 - PicSize.Height / 2 * multY
        VDT.Vform.PaintPicture PicSize.Image, X, Y, PicSize.Width * multX, PicSize.Height * multY, 0, 0, PicSize.Width, PicSize.Height
    Case 2    'Tiled
        X = VDT.Vform.Width / PicSize.Width
        Y = VDT.Vform.Height / PicSize.Height
        For J = 0 To Y * PicSize.Height Step PicSize.Height
            For I = 0 To X * PicSize.Width Step PicSize.Width
               VDT.Vform.PaintPicture PicSize.Image, I, J, PicSize.Width, PicSize.Height, 0, 0, PicSize.Width, PicSize.Height
            Next I
        Next J
End Select
Exit Sub
BGerr:
MsgBox "Unable to Locate Background Image file:" & vbCrLf & vbCrLf _
& SelectedFileName, vbExclamation, "File Not Found"
End Sub
Private Sub RemoveBackgroundImage()
Bkgnd.Cls: PicSize.Cls
SelectedFileName = ""
End Sub
Private Sub DetectBackgroundImage(ctrl As Control)
If SelectedFileName = "" Then
    ctrl.ToolTipText = ""
Else
    ctrl.ToolTipText = "Click Here to Remove Background Image"
End If
End Sub

