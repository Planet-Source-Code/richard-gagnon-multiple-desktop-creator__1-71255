VERSION 5.00
Begin VB.Form VDT 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Virtual Desktop"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10845
   Icon            =   "VDesktop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   723
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Menu3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   3150
      ScaleHeight     =   720
      ScaleWidth      =   1245
      TabIndex        =   32
      Top             =   630
      Visible         =   0   'False
      Width           =   1275
      Begin VB.VScrollBar VScroll1 
         Height          =   750
         Left            =   945
         TabIndex        =   35
         Top             =   0
         Width           =   250
      End
      Begin VB.PictureBox Menu2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   0
         ScaleHeight     =   330
         ScaleWidth      =   750
         TabIndex        =   33
         Top             =   0
         Width           =   750
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   100
            TabIndex        =   34
            Top             =   0
            Width           =   465
         End
      End
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   31
      Top             =   2625
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox Menu1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   1575
      ScaleHeight     =   1350
      ScaleWidth      =   1455
      TabIndex        =   24
      Top             =   630
      Visible         =   0   'False
      Width           =   1485
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   37
         Top             =   40
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Move To -->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   0
         TabIndex        =   30
         Top             =   1050
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copy To -->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   0
         TabIndex        =   29
         Top             =   840
         Width           =   1380
      End
      Begin VB.Line Line2 
         X1              =   105
         X2              =   1365
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   26
         Top             =   250
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rename"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   0
         TabIndex        =   25
         Top             =   460
         Width           =   1380
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   1680
   End
   Begin VB.PictureBox Broken 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      Picture         =   "VDesktop.frx":0ECA
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   4
      Top             =   1155
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox TBF 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   723
      TabIndex        =   0
      Top             =   0
      Width           =   10845
      Begin VB.CommandButton SCEdit 
         BackColor       =   &H00D8E9EC&
         Height          =   435
         Index           =   0
         Left            =   105
         Picture         =   "VDesktop.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Add Shortcut"
         Top             =   60
         Width           =   435
      End
      Begin VB.CheckBox ShowT 
         Caption         =   "Show Title"
         Height          =   225
         Left            =   4620
         TabIndex        =   28
         Top             =   315
         Value           =   1  'Checked
         Width           =   1170
      End
      Begin VB.CommandButton CmdOK 
         BackColor       =   &H00D8E9EC&
         Caption         =   "OK"
         Height          =   435
         Left            =   105
         TabIndex        =   20
         Top             =   2940
         Width           =   1065
      End
      Begin VB.CheckBox CHK1 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Do not show again...EVER"
         Height          =   330
         Left            =   1365
         TabIndex        =   19
         Top             =   3000
         Width           =   2325
      End
      Begin VB.PictureBox Arrow 
         BackColor       =   &H00D8E9EC&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   105
         Picture         =   "VDesktop.frx":193E
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   15
         Top             =   630
         Width           =   435
      End
      Begin VB.PictureBox Arrow1 
         BackColor       =   &H00D8E9EC&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   6720
         Picture         =   "VDesktop.frx":20A8
         ScaleHeight     =   435
         ScaleWidth      =   345
         TabIndex        =   14
         Top             =   630
         Width           =   345
      End
      Begin VB.CheckBox AllEdit 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Edit"
         Height          =   225
         Left            =   9030
         TabIndex        =   13
         ToolTipText     =   "Edit on/off"
         Top             =   150
         Value           =   1  'Checked
         Width           =   750
      End
      Begin VB.CommandButton SCEdit 
         BackColor       =   &H00D8E9EC&
         Height          =   435
         Index           =   2
         Left            =   1155
         Picture         =   "VDesktop.frx":2812
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Rename Shortcut"
         Top             =   60
         Width           =   435
      End
      Begin VB.CommandButton FormSettings 
         BackColor       =   &H00D8E9EC&
         Height          =   435
         Left            =   3780
         Picture         =   "VDesktop.frx":2F7C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Settings"
         Top             =   60
         Width           =   435
      End
      Begin VB.CheckBox Ahide 
         BackColor       =   &H00D8E9EC&
         Caption         =   "AutoHide"
         Height          =   225
         Left            =   4620
         TabIndex        =   10
         ToolTipText     =   "AutoHide"
         Top             =   75
         Width           =   1065
      End
      Begin VB.PictureBox MinimizeBut 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   10080
         Picture         =   "VDesktop.frx":36E6
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   9
         ToolTipText     =   "Minimize"
         Top             =   210
         Width           =   345
      End
      Begin VB.PictureBox CloseBut 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   10395
         Picture         =   "VDesktop.frx":3DA0
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   8
         ToolTipText     =   "Close"
         Top             =   210
         Width           =   345
      End
      Begin VB.ComboBox DesktopList 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5880
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select Desktop"
         Top             =   60
         Width           =   2955
      End
      Begin VB.CommandButton DesktopEdit 
         BackColor       =   &H00D8E9EC&
         Height          =   435
         Index           =   0
         Left            =   1995
         Picture         =   "VDesktop.frx":445A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Add Desktop"
         Top             =   60
         Width           =   435
      End
      Begin VB.CommandButton DesktopEdit 
         BackColor       =   &H00D8E9EC&
         Height          =   435
         Index           =   1
         Left            =   2520
         Picture         =   "VDesktop.frx":4BC4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Remove Desktop"
         Top             =   60
         Width           =   435
      End
      Begin VB.CommandButton DesktopEdit 
         BackColor       =   &H00D8E9EC&
         Height          =   435
         Index           =   2
         Left            =   3045
         Picture         =   "VDesktop.frx":532E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Rename Desktop"
         Top             =   60
         Width           =   435
      End
      Begin VB.CommandButton SCEdit 
         BackColor       =   &H00D8E9EC&
         Height          =   435
         Index           =   1
         Left            =   630
         Picture         =   "VDesktop.frx":5A98
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Remove Shortcut"
         Top             =   60
         Width           =   435
      End
      Begin VB.Label LBL3 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   105
         TabIndex        =   18
         Top             =   2205
         Width           =   8415
      End
      Begin VB.Label LBL1 
         BackColor       =   &H00D8E9EC&
         Height          =   1170
         Left            =   105
         TabIndex        =   17
         Top             =   1050
         Width           =   2220
         WordWrap        =   -1  'True
      End
      Begin VB.Label LBL2 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         Caption         =   "LBL2"
         Height          =   195
         Left            =   6720
         TabIndex        =   16
         Top             =   1050
         Width           =   1065
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Index           =   2
         X1              =   294
         X2              =   294
         Y1              =   7
         Y2              =   31
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Index           =   1
         X1              =   242
         X2              =   242
         Y1              =   7
         Y2              =   31
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Index           =   0
         X1              =   119
         X2              =   119
         Y1              =   7
         Y2              =   31
      End
   End
   Begin VB.PictureBox Holder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   105
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   3
      Top             =   735
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox Vform 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   735
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   21
      Top             =   630
      Width           =   750
      Begin VB.Label Title 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   210
         TabIndex        =   27
         Top             =   1365
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Grid 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   0
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label IconCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Tag             =   "0"
         Top             =   525
         Visible         =   0   'False
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Image IconImage 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   0
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         Picture         =   "VDesktop.frx":6202
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Outline 
         BorderWidth     =   3
         Height          =   480
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Iselect 
         BorderStyle     =   3  'Dot
         Height          =   330
         Left            =   105
         Top             =   945
         Visible         =   0   'False
         Width           =   360
      End
   End
End
Attribute VB_Name = "VDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------\
'Author: Richard E. Gagnon.                                |
'URL:    http://members.cox.net/reg501/                    |
'Email:  reg501@cox.net                                    |
'Copyright © 2007 Richard E. Gagnon. All Rights Reserved.  |
'----------------------------------------------------------/

Option Explicit
'----------------
Private Const Mtest = 0 'For test/debug mode set to "1"
'----------------
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'----------------
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SHF_desktop = "\My Desktop"
Private Const SHF_shortcut = "\My Shortcuts"
Private Const SHF_wallpaper = "\My Wallpaper"

'----------------
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'APIs needed to keep this form locked as a Desktop
'May not be compatible with VISTA
'****See also Form_Load****
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'--------------------------------------------------
Private MyFolder As String
Private INI As New INIclass
Private Const ScList = "\IconList.ini"
Private Const LastView = "LV"
Private Const SEP = ">"
Private oX As Single, oY As Single
Private CurScreenX As Single, CurScreenY As Single
Private ScreenX As Integer, ScreenY As Integer
Private DefaultBackColor As Long, DefaultFontColor As Long
Private IconLeft As Integer
Private PrevIndex As Integer
Private Smove As Boolean, Mmove As Boolean
Private ResChange As Boolean
Private NoMove As Boolean

Private Sub CmdOK_Click()
Timer1.Enabled = False
TBF.Height = 37
INI.Path = UserPath & MySystem
INI.Section = "Instructions"
INI.Key = "Show"
INI.Value = CHK1.Value
AllEdit.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.WindowState = 1 And X = LeftDown Then
    Me.WindowState = 0
    Me.Show
    TrayDelete
Else
    If Y < TBF.Height Then If Not TBF.Visible Then TBF.Visible = True
End If
End Sub

Private Sub Form_Resize()
Dim X As Single, Y As Single
X = GetSystemMetrics(SM_CXSCREEN) 'Screen width in pixels
Y = GetSystemMetrics(SM_CYSCREEN) 'Screen height in pixels
If Me.WindowState = 0 And (X <> CurScreenX Or Y <> CurScreenY) Then
    SetUpForm
    CurScreenX = X: CurScreenY = Y
    If Not ResChange Then
        ResChange = True
        MsgBox "Warning! Screen resolution may have changed. " & vbCrLf _
                & "Close program and restart to resize Desktop.", vbCritical, "Resolution Change"
    End If
End If
End Sub

Private Sub FormSettings_Click()
If DesktopList.ListCount > 0 Then Settings.Show vbModal
End Sub

Private Sub AllEdit_Click()
If AllEdit.Value = 0 Then
    SCEdit(0).Visible = False: SCEdit(1).Visible = False
    SCEdit(2).Visible = False
    DesktopEdit(0).Visible = False: DesktopEdit(1).Visible = False
    DesktopEdit(2).Visible = False: FormSettings.Visible = False
    ShowT.Visible = False: Ahide.Visible = False
    Line1(0).Visible = False: Line1(2).Visible = False
    MinimizeBut.Visible = False: CloseBut.Visible = False
    DesktopList.Left = 7: DesktopList.Width = 582
    
Else
    DesktopList.Left = 392: DesktopList.Width = 197
    SCEdit(0).Visible = True: SCEdit(1).Visible = True
    SCEdit(2).Visible = True
    DesktopEdit(0).Visible = True: DesktopEdit(1).Visible = True
    DesktopEdit(2).Visible = True: FormSettings.Visible = True
    ShowT.Visible = True: Ahide.Visible = True
    MinimizeBut.Visible = True: CloseBut.Visible = True
    Line1(0).Visible = True: Line1(2).Visible = True
End If
End Sub

Private Sub Grid_Click(Index As Integer)
'ResChange = False
'If IconImage(Index).Visible Then
'    Dim sFile As String
'    Dim sCommand As String
'    Dim sWorkDir As String
'    sFile = IconCaption(Index).Tag  'The file to execute
'    If InStr(1, LCase(sFile), ".url") Then
'        sCommand = vbNullString         'Command line parameters
'        sWorkDir = "C:\"                'The working directory
'        ExecuteShellCmd Me.hwnd, "open", sFile, sCommand, sWorkDir, 1
'    End If
'End If
End Sub

Private Sub Label1_Click(Index As Integer)
If Index < 3 Then
    Menu1.Visible = False
    SCEdit_Click (Index)
End If
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ML As label
For Each ML In Label1
    If ML.Index = Index Then
        Menu1.Tag = 1
        Label1(Index).BackColor = vbBlue
        Label1(Index).ForeColor = vbWhite
    Else
        Label1(ML.Index).BackColor = vbWhite
        Label1(ML.Index).ForeColor = vbBlack
    End If
Next
If Menu3.Visible And Menu3.Tag <> "" Then
    Menu3.Tag = ""
    For Each ML In Label2
        Label2(ML.Index).BackColor = vbWhite
        Label2(ML.Index).ForeColor = vbBlack
    Next
End If
If Index > 2 Then
    If Menu1.Left + Menu1.Width + Menu3.Width < CurScreenX Then
        Menu3.Left = Menu1.Left + Menu1.Width
    Else
        Menu3.Left = Menu1.Left - Menu3.Width
    End If
    VScroll1.Value = 0
    Menu3.Visible = True
Else
    Menu3.Visible = False
End If
End Sub

Private Sub Label2_Click(Index As Integer)
If Menu1.Visible Then
    DoMoveCopy (Label2(Index).Caption)
    Menu1.Visible = False
    Menu3.Visible = False
    If Label2(Index).Caption = DesktopList.Text Then GetIcons (DesktopList.Text)
Else
    'Setting DesktopList.Text attribute will trigger DesktopList_Click event
    DesktopList.Text = Label2(Index).Caption
End If
End Sub

Private Sub DoMoveCopy(DT As String)
Dim I As Integer
Dim Gpos As Integer
Dim sKeys() As String
Dim iKeycount As Long
Dim V As String

If Label1(4).BackColor = vbBlue And DT = DesktopList.Text Then Exit Sub
'Get occupied grids from file and sort in a list
List1.Clear
INI.Path = UserPath & ScList
INI.Section = DT
INI.EnumerateCurrentSection sKeys(), iKeycount
List1.AddItem "00000"
For I = 1 To iKeycount
    INI.Key = sKeys(I)
    V = INI.Key
    Do Until Len(V) = 5
        V = "0" & V
    Loop
    List1.AddItem V
Next

For Gpos = 1 To Grid.UBound
    If IconCaption(Gpos).Visible And Outline(Gpos).Visible Then
        'Look for next available grid for desktop
        INI.Key = 0
        If List1.ListCount > 0 Then
            For I = 1 To List1.ListCount - 1
                If Val(List1.List(I)) <> I Then
                    INI.Key = I
                    Exit For
                End If
            Next I
        End If
        If INI.Key = 0 And I < Grid.UBound + 1 Then INI.Key = I
        V = INI.Key
        If INI.Key > 0 Then
            INI.Value = IconCaption(Gpos).Tag & SEP & IconCaption(Gpos).Caption
            'Delete from current Desktop if a Move operation
            If Label1(4).BackColor = vbBlue Then
                INI.Section = DesktopList.Text
                Outline(Gpos).Visible = False
                IconCaption(Gpos).Tag = 0
                IconCaption(Gpos) = ""
                Set IconImage(Gpos) = Nothing
                IconCaption(Gpos).Visible = False
                IconImage(Gpos).Visible = False
                INI.Key = Gpos
                INI.DeleteKey
                INI.Section = DT
            End If
        End If
        Do Until Len(V) = 5
            V = "0" & V
        Loop
        List1.AddItem V
    End If
Next Gpos

'Free memory, empty list
List1.Clear

End Sub
Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ML As label
For Each ML In Label2
    If ML.Index = Index Then
        Menu3.Tag = 1
        Label2(Index).BackColor = vbBlue
        Label2(Index).ForeColor = vbWhite
    Else
        Label2(ML.Index).BackColor = vbWhite
        Label2(ML.Index).ForeColor = vbBlack
    End If
Next
End Sub

Private Sub SCEdit_Click(Index As Integer)
Dim I As Integer, Q As Integer, cnt As Integer
For I = 1 To Grid.UBound
    If IconCaption(I).Visible And Outline(I).Visible Then cnt = cnt + 1
Next I
Select Case Index
    Case 0
        ExecuteShellCmd Me.hwnd, "open", "C:\WINDOWS\explorer.scf", vbNullString, "C:\", 1
    Case 1  'Delete
        If cnt > 0 Then
            Q = MsgBox("Are you sure you want to Remove the Selected Shortcut(s)?", vbYesNo + vbQuestion, "Shortcut Delete")
            If Q = vbYes Then
                INI.Path = UserPath & ScList
                INI.Section = DesktopList.Text
                For I = 1 To Grid.UBound
                    If IconCaption(I).Visible And Outline(I).Visible Then
                        If InStr(1, IconCaption(I).Tag, MyFolder) And Dir(IconCaption(I).Tag) <> "" Then
                            Q = MsgBox("Remove '" & IconCaption(I).Caption & "' from Hardrive also?", vbYesNo + vbQuestion, "Hardrive Delete")
                            If Q = vbYes Then Kill IconCaption(I).Tag
                        End If
                        Outline(I).Visible = False
                        IconCaption(I).Tag = 0
                        IconCaption(I) = ""
                        Set IconImage(I) = Nothing
                        IconCaption(I).Visible = False
                        IconImage(I).Visible = False
                        INI.Key = I
                        INI.DeleteKey
                    End If
                Next I
                PrevIndex = 0
            End If
        End If
    Case 2  'Rename
        If cnt > 0 Then
            If cnt = 1 Then
                Dim NewName As String, OldName As String
                If Outline(PrevIndex).Visible Then
                    OldName = IconCaption(PrevIndex).Caption
                    NewName = InputBox("Enter A New Name For " & "'" & OldName & "'", "New Name", OldName)
                    If LTrim(RTrim(NewName)) = "" Then NewName = OldName
                    If OldName <> NewName Then
                        INI.Path = UserPath & ScList
                        INI.Section = DesktopList.Text
                        INI.Key = PrevIndex
                        INI.DeleteKey
                        INI.Key = PrevIndex
                        INI.Value = IconCaption(PrevIndex).Tag & SEP & NewName
                        IconCaption(PrevIndex).Caption = NewName
                    End If
                End If
            Else
                MsgBox "Can not rename multiple selection. " & vbCrLf _
                & "Please select one Shortcut to rename.", vbInformation, "Shortcut Rename Error"
            End If
        End If
End Select
End Sub

Private Sub ShowT_Click()
If ShowT.Value = 0 Then Title.Visible = False Else DoTitle
End Sub

Private Sub Timer1_Timer()
Static cnt As Byte
Dim ArrowPos As Integer
Dim ArrowCap As String
Timer1.Interval = 5000
Select Case cnt
    Case 0
        ArrowPos = SCEdit(0).Left
        ArrowCap = "Click this button to open Explorer to add Shortcut(s) to the Desktop or Right click on a shortcut."
    Case 1
        ArrowPos = SCEdit(1).Left
        ArrowCap = "Click this button to remove Shortcut(s) from the Desktop or Right click on a shortcut."
    Case 2
        ArrowPos = SCEdit(2).Left
        ArrowCap = "Click this button to rename a Shortcut or Right click on a shortcut."
    Case 3
        ArrowPos = DesktopEdit(0).Left
        ArrowCap = "Click this button to add a new Desktop"
    Case 4
        ArrowPos = DesktopEdit(1).Left
        ArrowCap = "Click this button to remove the Desktop."
    Case 5
        ArrowPos = DesktopEdit(2).Left
        ArrowCap = "Click this button to rename the Desktop."
    Case 6
        ArrowPos = FormSettings.Left
        ArrowCap = "Click this button to open a utility to set wallpaper, background and text color."
    Case 7
        ArrowPos = ShowT.Left
        ArrowCap = "Check the 'Show Title' box to show the Desktop title."
    Case 8
        ArrowPos = DesktopList.Left + DesktopList.Width / 2
        ArrowCap = "Select a Desktop using this dropdown box or Right click on an empty area."
    Case 9
        ArrowPos = AllEdit.Left
        ArrowCap = "Check/Uncheck this to show/hide all Desktop edit functions"
End Select
Arrow.Left = ArrowPos
LBL1.Left = ArrowPos
LBL1.Caption = ArrowCap
cnt = cnt + 1: If cnt = 10 Then cnt = 0
End Sub

Private Sub Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not TBF.Visible Then TBF.Visible = True
End Sub

Private Sub vForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.WindowState = 1 And X = LeftDown * 15 Then
    Me.WindowState = 0
    Me.Show
    TrayDelete
Else
    If Y / 15 < TBF.Height Then If Not TBF.Visible Then TBF.Visible = True
End If
End Sub

Private Sub CloseBut_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Mtest = 0 And App.PrevInstance = True Then
    MsgBox "VDesktop is already loaded", vbInformation
    End
End If
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'The following Lines of code Keep this Form always on bottom.
'(May not be compatible with VISTA)
Dim ProgMan&, shellDllDefView&, sysListView&
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
CurScreenX = GetSystemMetrics(SM_CXSCREEN)
CurScreenY = GetSystemMetrics(SM_CYSCREEN)
DefaultBackColor = Me.BackColor
DefaultFontColor = IconCaption(0).ForeColor
SetUpForm
'------------
If Mtest = 1 Then
    UserPath = App.Path & MyPath
    MyFolder = App.Path & SHF_desktop
Else
    UserPath = GetShellFolderPath(Me.hwnd, CSIDL_LOCAL_APPDATA) & MyPath
    MyFolder = GetShellFolderPath(Me.hwnd, CSIDL_MYDOCUMENTS) & SHF_desktop
End If
'------------
If Dir(UserPath, vbDirectory) = vbNullString Then MkDir UserPath
If Dir(MyFolder, vbDirectory) = vbNullString Then CreateDefaultDesktop MyFolder
'------------
LoadFormSettings
'------------
ShowInstructions
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'The following Lines of code Keep this Form always on bottom.
'(May not be compatible with VISTA)
ProgMan = FindWindow("progman", vbNullString)
shellDllDefView = FindWindowEx(ProgMan&, 0&, "shelldll_defview", vbNullString)
sysListView = FindWindowEx(shellDllDefView&, 0&, "syslistview32", vbNullString)
SetParent Me.hwnd, sysListView
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
App.TaskVisible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveFormSettings
End Sub

Private Sub Grid_DblClick(Index As Integer)
If IconImage(Index).Visible Then
    Dim sFile As String
    Dim sCommand As String
    Dim sWorkDir As String
    sFile = IconCaption(Index).Tag  'The file to execute
    sCommand = vbNullString         'Command line parameters
    sWorkDir = "C:\"                'The working directory
    ExecuteShellCmd Me.hwnd, "open", sFile, sCommand, sWorkDir, 1
End If
End Sub

Private Sub Grid_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If Index <> PrevIndex And Not IconImage(Index).Visible Then
    Outline(PrevIndex).Visible = False
    Outline(Index).Visible = True
    IconImage(Index).Picture = IconImage(PrevIndex).Picture
    IconCaption(Index).Caption = IconCaption(PrevIndex).Caption
    IconCaption(Index).Tag = IconCaption(PrevIndex).Tag
    IconCaption(PrevIndex).Tag = 0
    IconCaption(PrevIndex) = ""
    Set IconImage(PrevIndex) = Nothing
    IconImage(PrevIndex).Visible = False
    IconCaption(PrevIndex).Visible = False
    IconImage(Index).Visible = True
    IconCaption(Index).Visible = True
    INI.Path = UserPath & ScList
    INI.Section = DesktopList.Text
    INI.Key = PrevIndex
    INI.DeleteKey
    INI.Key = Index
    INI.Value = IconCaption(Index).Tag & SEP & IconCaption(Index).Caption
    PrevIndex = Index
End If
NoMove = False
End Sub

Private Sub Grid_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Menu1.Visible = False: Menu3.Visible = False
If Button = vbLeftButton Then ClearSCHighlite
If IconImage(Index).Visible And (X / 15 >= IconLeft And X / 15 <= IconLeft + IconImage(Index).Width) And (Y / 15 >= 0 And Y / 15 <= IconImage(Index).Height) Then
    PrevIndex = Index
    Outline(Index).Visible = True
End If
Iselect.Width = 1: Iselect.Height = 1
oX = Grid(Index).Left + X / 15: oY = Grid(Index).Top + Y / 15
'-----------------------------
Menu1.Top = oY + 2: Menu1.Left = oX + 2
If Menu1.Width + oX > CurScreenX Then Menu1.Left = oX - 2 - Menu1.Width
If Menu1.Height + oY + GetTaskBarHeight > CurScreenY Then Menu1.Top = oY - 2 - Menu1.Height
'-----------------------------
Menu3.Top = oY + 2: Menu3.Left = oX + 2
If Menu3.Width + oX > CurScreenX Then Menu3.Left = oX - 2 - Menu3.Width
If Menu3.Height + oY + GetTaskBarHeight > CurScreenY Then Menu3.Top = oY - 2 - Menu3.Height
'-----------------------------
If Button = vbLeftButton Then
    Smove = True
    If Outline(Index).Visible Then Mmove = True: Smove = False
Else
    If Outline(Index).Visible Then
        Menu1.Visible = True
    Else
        If DesktopList.ListCount > 0 Then VScroll1.Value = 0: Menu3.Visible = True
    End If

End If
End Sub

Private Sub Grid_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ML As label
If Not Menu3.Visible And (Menu1.Visible And Menu1.Tag <> "") Then
    Menu1.Tag = ""
    For Each ML In Label1
        Label1(ML.Index).BackColor = vbWhite
        Label1(ML.Index).ForeColor = vbBlack
    Next
End If
If Menu3.Visible And Menu3.Tag <> "" Then
    Menu3.Tag = ""
    For Each ML In Label2
        Label2(ML.Index).BackColor = vbWhite
        Label2(ML.Index).ForeColor = vbBlack
    Next
End If

If Mmove Then
    If Not NoMove Then
        NoMove = True
    Else
        IconImage(Index).Drag
        IconImage(Index).DragIcon = IconImage(Index).Picture
        Mmove = False
    End If
Else
    If Not Smove Then
        If IconImage(Index).Visible Then
            If (X / 15 >= IconLeft And X / 15 <= IconLeft + IconImage(Index).Width) And (Y / 15 >= 0 And Y / 15 <= IconImage(Index).Height) Then
                If Outline(Index).Visible Then
                    Grid(Index).ToolTipText = IconCaption(Index).Caption
                Else
                    Grid(Index).ToolTipText = IconCaption(Index).Tag
                End If
            Else
                Grid(Index).ToolTipText = ""
            End If
        Else
            Grid(Index).ToolTipText = ""
        End If
        If Ahide.Value Then If TBF.Visible Then TBF.Visible = False
        If Me.WindowState = 0 And (GetSystemMetrics(SM_CXSCREEN) <> CurScreenX Or GetSystemMetrics(SM_CYSCREEN) <> CurScreenY) Then Form_Resize
    End If
End If
If Smove Then
    Iselect.Visible = False
    Dim Mx As Single, My As Single
    Mx = Grid(Index).Left + X / 15
    My = Grid(Index).Top + Y / 15
    If Mx < oX Then Iselect.Left = Mx Else Iselect.Left = oX
    If My < oY Then Iselect.Top = My Else Iselect.Top = oY
    Iselect.Width = Abs(Mx - oX)
    Iselect.Height = Abs(My - oY)
    Dim I As Integer
    Dim Gridx As Single, Gridy As Single
    Dim Gridw As Single, Gridh As Single
    Gridw = Iselect.Width + Iselect.Left
    Gridh = Iselect.Height + Iselect.Top
    Iselect.Visible = True
    For I = 1 To Grid.UBound
        If IconImage(I).Visible = True Then
            'Get the screen location at the center of each grid
            Gridx = Grid(I).Left + Grid(I).Width / 2
            Gridy = Grid(I).Top + Grid(I).Height / 2
            If Iselect.Left < Gridx And Gridw > Gridx And Iselect.Top < Gridy And Gridh > Gridy Then
                If Not Outline(I).Visible Then Outline(I).Visible = True
            Else
                If Outline(I).Visible Then Outline(I).Visible = False
            End If
        End If
    Next
End If
End Sub

Private Sub Grid_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Mmove = False: NoMove = False: Smove = False: Iselect.Visible = False
End Sub

Private Sub Grid_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errout
Dim FullPath As String
Dim Fname As Variant
Dim I As Integer
INI.Path = UserPath & ScList
INI.Section = DesktopList.Text
ClearSCHighlite
If Data.GetFormat(vbCFFiles) Then
    If Data.Files.Count = 1 And Not IconImage(Index).Visible Then
        ShowIcon Index, Data.Files.Item(1)
        INI.Key = Index
        INI.Value = IconCaption(Index).Tag & SEP & IconCaption(Index).Caption
        Outline(Index).Visible = True
        PrevIndex = Index
    Else
        For Each Fname In Data.Files
            For I = 1 To IconImage.UBound
                If Not IconImage(I).Visible Then Exit For
            Next I
            If I < IconImage.UBound + 1 Then
                FullPath = Fname
                ShowIcon I, FullPath
                INI.Key = I
                INI.Value = IconCaption(I).Tag & SEP & IconCaption(I).Caption
            End If
        Next Fname
    End If
Else
    Dim Fnum As Integer
    Dim Flink As String, URLname As String
    URLname = InputBox("Enter A Name For This Link", "URL Name")
    Flink = MyFolder & SHF_shortcut & "\" & URLname & ".url"
    If Dir(Flink) = "" Then
        Fnum = FreeFile
        Open Flink For Output As Fnum
        Print #Fnum, "[InternetShortcut]"
        Print #Fnum, "URL=" & Data.GetData(vbCFText)
        Close Fnum
        If Not IconImage(Index).Visible Then
            ShowIcon Index, Flink
            INI.Key = Index
            INI.Value = IconCaption(Index).Tag & SEP & IconCaption(Index).Caption
            Outline(Index).Visible = True
            PrevIndex = Index
        End If
    Else
        MsgBox "'" & URLname & "' Already Exists", vbExclamation, "URL Exists"
    End If
End If
errout:
End Sub

Private Sub DesktopEdit_Click(Index As Integer)
Dim NewName As String, OldName As String
Dim I As Integer
Dim Chg As Boolean, Dup As Boolean
ClearSCHighlite
Select Case Index
    Case 0  'Add Desktop
        NewName = InputBox("Enter a name for this Desktop ", "New Desktop", "Desktop Name")
        If LTrim(RTrim(NewName)) <> "" Then
            If Not DuplicateName(NewName) Then
                DesktopList.AddItem NewName
                RemoveAllShortcuts
                INI.Path = UserPath & ScList
                INI.Section = NewName
                INI.Key = 1
                INI.Value = "Dummy"
                INI.DeleteKey
                'Setting DesktopList.Text attribute will trigger DesktopList_Click event
                DesktopList.Text = NewName
                Chg = True
            Else
                Dup = True
            End If
        End If
    Case 1  'Remove Desktop
        If DesktopList.ListCount > 0 Then
            I = MsgBox("Are you sure you want to remove the Desktop " _
                & vbCrLf & vbCrLf & Space(20) & "'" & DesktopList.Text _
                & "'", vbYesNo + vbQuestion, "Desktop Delete!")
            If I = vbYes Then
                Chg = True
                RemoveAllShortcuts
                INI.Path = UserPath & ScList
                INI.Section = DesktopList.Text
                INI.DeleteSection
                '----------------
                For I = 0 To DesktopList.ListCount
                    If DesktopList.List(I) = DesktopList.Text Then Exit For
                Next
                INI.Path = UserPath & MySystem
                INI.Section = DesktopList.Text
                INI.DeleteSection
                DesktopList.RemoveItem I
                If DesktopList.ListCount > 0 Then
                    'Setting DesktopList.Text attribute will trigger DesktopList_Click event
                    DesktopList.Text = DesktopList.List(0)
                    INI.Path = UserPath & MySystem
                    INI.Section = LastView
                    INI.Key = LastView
                    INI.Value = DesktopList.Text
                Else
                    Label2(0).Caption = ""
                    Vform.Cls
                    Vform.BackColor = DefaultBackColor
                End If
            End If
        End If
    Case 2  'Rename Desktop
        If DesktopList.ListCount > 0 Then
            Dim Fname As String
            Dim CurrBackColor As String
            Dim CurrFontColor As String
            Dim CurrBackImg As String
            OldName = DesktopList.Text
            NewName = InputBox("Enter a new name for Desktop " & "'" & DesktopList.Text & "'", "New Name", OldName)
            If LTrim(RTrim(NewName)) = "" Then NewName = OldName
            If OldName <> NewName Then
                If Not DuplicateName(NewName) Then
                    For I = 0 To DesktopList.ListCount
                        If DesktopList.List(I) = DesktopList.Text Then Exit For
                    Next
                    DesktopList.RemoveItem I
                    DesktopList.AddItem NewName
                    INI.Path = UserPath & MySystem
                    INI.Section = LastView
                    INI.Key = LastView
                    INI.Value = NewName
                    '===================
                    INI.Section = OldName
                    INI.Key = "BGimage"
                    If INI.Value <> "" Then CurrBackImg = INI.Value
                    INI.Key = "BGcolor"
                    If INI.Value <> "" Then CurrBackColor = INI.Value
                    INI.Key = "Fcolor"
                    If INI.Value <> "" Then CurrFontColor = INI.Value
                    INI.Key = "Smode"
                    FormSettings.Tag = INI.Value
                    INI.DeleteSection
                    '-------------------
                    INI.Section = NewName
                    If CurrBackImg <> "" Then
                        INI.Key = "BGimage"
                        INI.Value = CurrBackImg
                    End If
                    If CurrBackColor <> "" Then
                        INI.Key = "BGcolor"
                        INI.Value = CurrBackColor
                    End If
                    If CurrFontColor <> "" Then
                        INI.Key = "Fcolor"
                        INI.Value = CurrFontColor
                    End If
                    INI.Key = "Smode"
                    INI.Value = FormSettings.Tag
                    '===================
                    INI.Path = UserPath & ScList
                    INI.Section = OldName
                    INI.DeleteSection
                    SaveDesktopIcons NewName
                    'Setting DesktopList.Text attribute will trigger DesktopList_Click event
                    DesktopList.Text = NewName
                    DoTitle
                    Chg = True
                Else
                    Dup = True
                End If
            Else
                Dup = True
            End If
        End If
End Select
If Chg Then LoadDesktopList
If Dup Then MsgBox "Desktop name '" & NewName & "' already exists", vbInformation, "Duplicate Desktop Name"
End Sub
Private Function DuplicateName(Dname As String) As Boolean
Dim DT As label
Dim Dup As Boolean
For Each DT In Label2
    If UCase(Label2(DT.Index).Caption) = UCase(Dname) Then Dup = True
Next
DuplicateName = Dup
End Function

Private Sub DesktopList_Click()
FreezeWindow Me.hwnd
ClearSCHighlite
ProcessFormSettings
GetIcons DesktopList.Text
If TBF.Visible Then TBF.SetFocus
FreezeWindow 0
End Sub

Private Sub SaveDesktopIcons(Gname As String)
On Error Resume Next
Dim I As Integer
INI.Path = UserPath & ScList
INI.Section = Gname
INI.DeleteSection
INI.Section = Gname
For I = 1 To IconImage.UBound
    If IconCaption(I).Visible Then
        INI.Key = I
        INI.Value = IconCaption(I).Tag & SEP & IconCaption(I).Caption
    End If
Next I
End Sub
Private Sub GetIcons(GPname As String)
Dim I As Integer
Dim sKeys() As String
Dim iKeycount As Long
Dim Enough As Boolean
FreezeWindow Me.hwnd
RemoveAllShortcuts
INI.Path = UserPath & ScList
INI.Section = GPname
INI.EnumerateCurrentSection sKeys(), iKeycount
If iKeycount > Grid.UBound Then
    Enough = True
    I = MsgBox("There are more Shortcuts than can be displayed. " & vbCrLf _
        & "Consider increasing screen resolution" & vbCrLf _
        & "Continue Anyway?", vbYesNo + vbQuestion, "Shortcut Load Error")
End If
If I = vbYes Or iKeycount <= Grid.UBound Then
    For I = 1 To iKeycount
        INI.Key = sKeys(I)
        If Val(INI.Key) <= Grid.UBound Then
            ShowIcon INI.Key, INI.Value
            DoEvents
        Else
            If Not Enough Then
                MsgBox "There are more Shortcuts than can be displayed. " & vbCrLf _
                & "Consider increasing screen resolution", vbInformation, "Shortcut Load Error"
            End If
            Enough = True
        End If
    Next
End If
FreezeWindow 0
End Sub
Private Sub RemoveAllShortcuts()
Dim I As Integer
For I = 1 To Grid.UBound
    IconCaption(I).Tag = 0
    IconImage(I).Visible = False
    IconCaption(I).Visible = False
    IconCaption(I).Caption = ""
    Set IconImage(I) = Nothing
Next I
End Sub
Private Sub ShowIcon(Index As Integer, Path As String)
Dim Bname As String, Pname As String
Dim Pos As Integer
On Error Resume Next
Pos = InStr(1, Path, SEP)
If Pos > 0 Then
    Pname = Mid(Path, 1, Pos - 1)
    Bname = Mid(Path, Pos + 1, Len(Path))
Else
    Pname = Path
    Bname = GetBase(Path)
End If
IconImage(0).Picture = GetIcon(Pname, 0)
'IconImage(0).Picture = GetBigIcon(Pname)
If IconImage(0).Picture <> 0 Then
    IconImage(Index).Picture = IconImage(0).Picture
Else
    IconImage(Index).Picture = Broken.Picture
End If
IconCaption(Index).Caption = Bname
IconCaption(Index).Tag = Pname
IconCaption(Index).Visible = True
IconImage(Index).Visible = True
End Sub
Private Sub LoadFormSettings()
Dim I As Integer
Dim sKeys() As String
Dim iKeycount As Long
Dim sSections() As String
Dim iSectionCount As Long
Dim Dprev As String

IconImage(0).Height = 32: IconImage(0).Width = 32
IconCaption(0).Height = 39: IconCaption(0).Width = 51
CreateGrids
INI.Path = UserPath & MySystem
INI.Section = "Edit"
INI.Key = "Edit"
If INI.Value <> "" Then AllEdit.Value = Val(INI.Value)
'----------------------
INI.Section = "Title"
INI.Key = "Title"
If INI.Value <> "" Then ShowT.Value = Val(INI.Value)
'----------------------
INI.Section = "AutoHide"
INI.Key = "AutoHide": Ahide.Value = Val(INI.Value)
'---------------------
INI.Section = LastView
INI.Key = LastView: Dprev = INI.Value
INI.Path = UserPath & ScList
If Dprev <> "" Then
    INI.EnumerateAllSections sSections(), iSectionCount
    If iSectionCount > 0 Then
        For I = 1 To iSectionCount
           DesktopList.AddItem sSections(I)
        Next I
    Else
        DesktopList.AddItem Dprev
    End If
    'Setting DesktopList.Text attribute will trigger DesktopList_Click event
    DesktopList.Text = Dprev
End If
'----------------------------------
LoadDesktopList
'----------------------------------
End Sub
Private Sub LoadDesktopList()
Dim I As Integer, cnt As Integer
Dim LT As Integer
Dim label As label

For Each label In Label2
    If label.Index > 0 Then Unload Label2(label.Index)
Next

cnt = DesktopList.ListCount
LT = cnt - 1
If cnt > 0 Then
    Label2(0).Left = 100
    Label2(0).Caption = DesktopList.List(0)
    Menu2.Width = (Label2(0).Width + 200)
    If cnt > 1 Then
        cnt = cnt - 1
        For I = 1 To cnt
            Load Label2(I)
            Label2(I).Visible = True
            Label2(I).Caption = DesktopList.List(I)
            Label2(I).Top = Label2(I - 1).Top + Label2(I - 1).Height + 10
            Label2(I).Left = 100
            If (Label2(I).Width + 200) > Menu2.Width Then
                Menu2.Width = (Label2(I).Width + 200)
            End If
        Next I
    End If
    Menu2.Height = (Label2(LT).Top + Label2(LT).Height + 100)
End If
'======================
If Label2.Count > 8 Then
    Menu3.Height = (Label2(7).Top + Label2(7).Height + 50) / 15
    Menu3.Width = (Menu2.Width + 200) / 15
    VScroll1.Max = Label2.Count - 8
    VScroll1.Left = Menu3.Width * 15 - VScroll1.Width
    VScroll1.Height = Menu3.Height * 15
    VScroll1.Visible = True
Else
    Menu3.Width = (Menu2.Width + 100) / 15
    Menu3.Height = Menu2.Height / 15
    VScroll1.Visible = False
End If
End Sub

Private Sub SaveFormSettings()
INI.Path = UserPath & MySystem
INI.Section = LastView
INI.Key = LastView
INI.Value = DesktopList.Text
'------------------------
INI.Section = "AutoHide"
INI.Key = "AutoHide"
INI.Value = Ahide.Value
'------------------------
INI.Section = "Edit"
INI.Key = "Edit"
INI.Value = AllEdit.Value
'------------------------
INI.Section = "Title"
INI.Key = "Title"
INI.Value = ShowT.Value
End Sub
Private Sub CreateGrids()
Dim X As Integer, Y As Single
Dim I As Integer, J As Integer
Dim X1 As Single
Dim SizeX As Single, SizeY As Single
Dim J1 As Integer
Dim x2 As Integer, y2 As Integer
Dim A As Integer

Grid(0).Width = 1024 / 13
Grid(0).Height = 768 / 9.5

A = GetTaskBarHeight
SizeX = Grid(0).Width
x2 = ScreenX / SizeX
SizeX = ScreenX / x2
Grid(0).Width = SizeX

SizeY = Grid(0).Height
y2 = (ScreenY - TBF.Height - A) / SizeY
SizeY = (ScreenY - TBF.Height - A) / y2
Grid(0).Height = SizeY


X = 0
Y = TBF.Height + 2
IconLeft = (SizeX - IconImage(J1).Width) / 2
For I = 0 To y2 - 1
    X1 = X
    For J = I * x2 To I * x2 + (x2 - 1)
        J1 = J + 1
        Load Grid(J1)
        Grid(J1).BorderStyle = 0
        Grid(J1).Visible = True
        Grid(J1).Top = Y
        Grid(J1).Left = X1
        Load IconImage(J1)
        Load IconCaption(J1)
        IconCaption(J1).BorderStyle = 0
        Load Outline(J1)
        IconImage(J1).Top = Grid(J1).Top
        IconImage(J1).Left = Grid(J1).Left + IconLeft
        Outline(J1).Top = Grid(J1).Top
        Outline(J1).Left = Grid(J1).Left + IconLeft
        IconCaption(J1).Width = Grid(J1).Width - 4
        IconCaption(J1).Top = IconImage(J1).Top + IconImage(J1).Height
        IconCaption(J1).Left = Grid(J1).Left + 2
        IconCaption(J1).Height = Grid(J1).Height - IconImage(J1).Height - 6
        IconImage(J1).ZOrder vbBringToFront
        IconCaption(J1).ZOrder vbBringToFront
        Outline(J1).ZOrder vbBringToFront
        Grid(J1).ZOrder vbBringToFront
        X1 = X1 + SizeX
    Next J
    Y = Y + SizeY
Next I
Mmove = False: NoMove = False
TBF.ZOrder vbBringToFront
End Sub
Private Sub SetUpForm()

ScreenX = GetSystemMetrics(SM_CXSCREEN) 'Screen width in pixels
ScreenY = GetSystemMetrics(SM_CYSCREEN) 'Screen height in pixels

Me.Left = 0: Me.Top = 0
Me.Width = ScreenX * 15
Me.Height = (ScreenY - GetTaskBarHeight) * 15

Vform.Left = 0: Vform.Top = 0
Vform.Width = ScreenX
Vform.Height = ScreenY - GetTaskBarHeight
Vform.ZOrder vbSendToBack

LBL3.Width = TBF.Width * 15
CloseBut.Left = TBF.Width - CloseBut.Width
MinimizeBut.Left = TBF.Width - CloseBut.Width - MinimizeBut.Width
Arrow1.Left = MinimizeBut.Left
LBL2.Left = MinimizeBut.Left - LBL2.Width
End Sub

Private Sub MinimizeBut_Click()
Me.WindowState = 1
TrayAdd Me.hwnd, Me.Icon, "Virtual Desktop", MouseMove
Me.Hide
End Sub
Private Sub ProcessFormSettings()
Dim I As Integer
Dim FC As Long
On Error Resume Next
If DesktopList.Text <> "" Then
    INI.Path = UserPath & MySystem
    INI.Section = DesktopList.Text
    '---Get Background Image
    INI.Key = "BGimage"
    Vform.Tag = INI.Value
    '---Get Background Color
    INI.Key = "BGcolor"
    Vform.Cls
    If INI.Value <> "" Then
        Vform.BackColor = Val(INI.Value)
    Else
        Vform.BackColor = DefaultBackColor
    End If
    '---Get Font Color
    INI.Key = "Fcolor"
    If INI.Value = "" Then FC = DefaultFontColor Else FC = Val(INI.Value)
    Iselect.BorderColor = FC
    For I = 1 To IconCaption.UBound
        IconCaption(I).ForeColor = FC
        Outline(I).BorderColor = FC
    Next I
    '---Get Image Mode
    INI.Key = "Smode"
    FormSettings.Tag = INI.Value
    If FormSettings.Tag = "False" Then ' For backward compatibility
        FormSettings.Tag = "0"
        INI.Value = 0
    End If
    If FormSettings.Tag = "True" Then ' For backward compatibility
        FormSettings.Tag = "1"
        INI.Value = 1
    End If
    '---Process Image
    If Vform.Tag <> "" Then
        Holder.Picture = LoadPicture(Vform.Tag)
        If Err.Number > 0 Then
            MsgBox "Unable to Locate Background Image file:" & vbCrLf & vbCrLf _
            & Vform.Tag, vbExclamation, "File Not Found"
        Else
            ProcessBackground
        End If
    Else
        Holder.Picture = Nothing
        Vform.Picture = Nothing
    End If
    '---Align and Display Title
    If ShowT.Value Then DoTitle Else Title.Visible = False
End If
End Sub
Private Sub DoTitle()
Dim FC As Long
INI.Path = UserPath & MySystem
INI.Section = DesktopList.Text
INI.Key = "Fcolor"
If INI.Value = "" Then FC = DefaultFontColor Else FC = Val(INI.Value)
Title.Caption = Trim(DesktopList.Text)
Title.ForeColor = FC
Title.Top = 0
Title.Left = (CurScreenX / 2) - (Title.Width / 2)
Title.Visible = True
End Sub
Private Sub ShowInstructions()
INI.Path = UserPath & MySystem
INI.Section = "Instructions"
INI.Key = "Show"
If Val(INI.Value) <> 1 Then
    LBL2.Caption = "Minimize to System Tray or Close Virtual Desktop"
    LBL3.Caption = " To add Shortcuts to your Virtual Desktop, Open My Computer or " _
             & "explorer, select and drag files or folders to the Virtual Desktop"
    Timer1.Enabled = True
    TBF.Height = 227
    AllEdit.Value = 1
    AllEdit.Enabled = False
End If
End Sub

Private Sub ClearSCHighlite()
Dim I As Integer
For I = 1 To Grid.UBound
    Outline(I).Visible = False
Next I
Menu1.Visible = False
Menu3.Visible = False
End Sub
Private Sub CreateDefaultDesktop(mpath As String)
Dim FileList() As String
Dim I As Integer
Dim Src As String, Dst As String

'Create User folders and copy existing desktop shortcuts to new desktop and folder
MkDir mpath
MkDir mpath & SHF_wallpaper
MkDir mpath & SHF_shortcut
CopyFiles Me.hwnd, GetShellFolderPath(Me.hwnd, CSIDL_DESKTOP) & "\*.*", mpath & SHF_shortcut
FileList = Split(GetFiles(mpath & SHF_shortcut & "\*"), Chr(0))
'---------------------------
INI.Path = UserPath & MySystem
INI.Section = LastView
INI.Key = LastView
INI.Value = "My Desktop"
'===========================
INI.Path = UserPath & ScList
INI.Section = "My Desktop"
For I = 0 To UBound(FileList) - 1
    INI.Key = I + 1
    INI.Value = mpath & SHF_shortcut & "\" & FileList(I) & SEP & FileList(I)
Next I
End Sub
Private Sub ProcessBackground()
Dim multX As Single, multY As Single
Dim X As Integer, Y As Integer
Dim I As Integer, J As Integer

Select Case Val(FormSettings.Tag)
    Case 0    'Stretched
        multX = Vform.Width / Holder.Width
        multY = Vform.Height / Holder.Height
        Vform.PaintPicture Holder.Image, 0, 0, Holder.Width * multX, Holder.Height * multY, 0, 0, Holder.Width, Holder.Height
    Case 1    'Centered
        multX = Vform.Width / Holder.Width
        multY = Vform.Height / Holder.Height
        If multX > 1 And multY > 1 Then
            multX = 1: multY = 1
        Else
            If multX > multY Then multX = multY Else multY = multX
        End If
        
        X = Vform.Width / 2 - Holder.Width / 2 * multX
        Y = Vform.Height / 2 - Holder.Height / 2 * multY
        Vform.PaintPicture Holder.Image, X, Y, Holder.Width * multX, Holder.Height * multY, 0, 0, Holder.Width, Holder.Height
    Case 2    'Tiled
        X = Vform.Width / Holder.Width
        Y = Vform.Height / Holder.Height
        For J = 0 To Y * Holder.Height Step Holder.Height
            For I = 0 To X * Holder.Width Step Holder.Width
               Vform.PaintPicture Holder.Image, I, J, Holder.Width, Holder.Height, 0, 0, Holder.Width, Holder.Height
            Next I
        Next J
End Select
End Sub

Private Sub VScroll1_Change()
Menu2.Top = Label2(VScroll1.Value).Top * -1
End Sub
