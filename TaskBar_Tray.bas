Attribute VB_Name = "TaskBar_Tray"
Option Explicit

'================Taskbar==============
Private Const ABM_GETTASKBARPOS = &H5
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type APPBARDATA
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

'================Tray=================
Const NIF_MESSAGE    As Long = &H1     'Message
Const NIF_ICON       As Long = &H2     'Icon
Const NIF_TIP        As Long = &H4     'TooTipText
Const NIM_ADD        As Long = &H0     'Add to tray
Const NIM_MODIFY     As Long = &H1     'Modify
Const NIM_DELETE     As Long = &H2     'Delete From Tray

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Enum TrayRetunEventEnum
    MouseMove = &H200       'On Mousemove
    LeftUp = &H202          'Left Button Mouse Up
    LeftDown = &H201        'Left Button MouseDown
    LeftDbClick = &H203     'Left Button Double Click
    RightUp = &H205         'Right Button Up
    RightDown = &H204       'Right Button Down
    RightDbClick = &H206    'Right Button Double Click
    MiddleUp = &H208        'Middle Button Up
    MiddleDown = &H207      'Middle Button Down
    MiddleDbClick = &H209   'Middle Button Double Click
End Enum

Public Enum ModifyItemEnum
    ToolTip = 1             'Modify ToolTip
    Icon = 2                'Modify Icon
End Enum

Private TrayIcon As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'[Task Bar Height]
Public Function GetTaskBarHeight() As Long
    Dim ABD As APPBARDATA

    SHAppBarMessage ABM_GETTASKBARPOS, ABD
    GetTaskBarHeight = ABD.rc.Bottom - ABD.rc.Top
End Function
'[Task Bar Width]
Public Function GetTaskBarWidth() As Long
    Dim ABD As APPBARDATA

    SHAppBarMessage ABM_GETTASKBARPOS, ABD
    GetTaskBarWidth = ABD.rc.Right - ABD.rc.Left
End Function
'[Add to Tray]
Public Sub TrayAdd(hWnd As Long, Icon As Picture, _
                    ToolTip As String, ReturnCallEvent As TrayRetunEventEnum)
    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hWnd = hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = ReturnCallEvent
        .hIcon = Icon
        .szTip = ToolTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, TrayIcon
End Sub
'[Remove From tray]
Public Sub TrayDelete()
    Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub
'[Modify the tray]
Public Sub TrayModify(Item As ModifyItemEnum, vNewValue As Variant)
    Select Case Item
        Case ToolTip
            TrayIcon.szTip = vNewValue & vbNullChar
        Case Icon
            TrayIcon.hIcon = vNewValue
    End Select
    Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub

