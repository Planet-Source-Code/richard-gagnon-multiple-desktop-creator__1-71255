Attribute VB_Name = "Functions"
Option Explicit

'=============Globals========================
Public UserPath As String
Public Const MyPath = "\VirtualDesktop"
Public Const MySystem = "\Layout.ini"
'============================================
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE As Long = -1
'============================================
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
'============================================
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'dwFlags
Private Const CSIDL_FLAG_PER_USER_INIT = &H800
Private Const CSIDL_FLAG_NO_ALIAS = &H1000
Private Const CSIDL_FLAG_DONT_VERIFY = &H4000
Private Const CSIDL_FLAG_CREATE = &H8000
Private Const CSIDL_FLAG_MASK = &HFF00
Private Const SHGFP_TYPE_CURRENT = &H0 'current value for user, verify it exists
Private Const SHGFP_TYPE_DEFAULT = &H1
'============================================
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * 14
End Type
'=======Launch Associated File API===========
Declare Function ShellExecute _
    Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

'============ShellFolders API================
Declare Function SHGetFolderPath _
   Lib "shfolder.dll" Alias "SHGetFolderPathA" ( _
   ByVal hwndOwner As Long, _
   ByVal nFolder As Long, _
   ByVal hToken As Long, _
   ByVal dwReserved As Long, _
   ByVal lpszPath As String) As Long
'-------ShellFolders API Constants ----------
Public Enum SF
    CSIDL_MYDOCUMENTS = &H5
    CSIDL_MYMUSIC = &HD
    CSIDL_MYVIDEO = &HE
    CSIDL_MYPICTURES = &H27
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_DOCANDSET = &H28
    
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_COMMON_PICTURES = &H36
    '-------------------------------------------
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    
    CSIDL_COMMON_OEM_LINKS = &H3A
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_ADMINTOOLS = &H30
    CSIDL_CONNECTIONS = &H31
    CSIDL_RESOURCES = &H38
    CSIDL_RESOURCES_LOCALIZED = &H39
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMPUTERSNEARME = &H3D
End Enum
'===============ShellFileOperations================
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (ByRef lpFileOp As SHFILEOPSTRUCT) As Long
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Long
    hNameMaps As Long
    sProgress As String
End Type
'================GetIconFromFile================
Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type

Private Type CLSID
    id(16) As Byte
End Type

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private SIconInfo As SHFILEINFO
Private Const SHGFI_ICON = &H100

Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Public Function ExecuteShellCmd(Handle As Long, mOp As String, mFile As String, mParm As String, mDir As String, mCmd As Long) As Long
ExecuteShellCmd = ShellExecute(Handle, mOp, mFile, mParm, mDir, mCmd)
End Function

Public Function GetShellFolderPath(hWind As Long, SpcFolder As Long) As String
'Part of ShellFolders API
Dim ShellBuff As String
Dim dwFlags As Long
Dim uType As Long
ShellBuff = Space(260)
dwFlags = CSIDL_FLAG_PER_USER_INIT
uType = SHGFP_TYPE_CURRENT

If SHGetFolderPath(hWind, SpcFolder Or dwFlags, -1, uType, ShellBuff) = 0 Then
    GetShellFolderPath = Left$(ShellBuff, lstrlenW(StrPtr(ShellBuff)))
Else
    GetShellFolderPath = "ERR"
End If

End Function

Public Function FreezeWindow(Handle As Long) As Long
LockWindowUpdate Handle
End Function
Public Function GetBase(Path As String) As String
Dim S0 As Long, S1 As Long
S0 = InStrRev(Path, ".", Len(Path))
If S0 = 0 Then S0 = Len(Path) Else S0 = S0 - 1
S1 = InStrRev(Path, "\", Len(Path))
GetBase = Mid(Path, S1 + 1, S0 - S1)
End Function
Private Function TrimNull(sFileName As String) As String
    Dim I As Long
    I = InStr(1, sFileName, vbNullChar)
    If I = 0 Then
        TrimNull = sFileName
    Else
        TrimNull = Left$(sFileName, I - 1)
    End If
End Function
Public Function GetFiles(Mfiles As String) As String
Dim sBuff As String
Dim iSearchHandle As Long
Dim pBuf As WIN32_FIND_DATA
Dim Tfile As String

sBuff = vbNullString
iSearchHandle = FindFirstFile(Mfiles, pBuf)
If iSearchHandle <> INVALID_HANDLE_VALUE Then
    Do While FindNextFile(iSearchHandle, pBuf)
        Tfile = TrimNull(pBuf.cFileName)
        If Tfile <> "." And Tfile <> ".." Then
            sBuff = sBuff & Tfile & Chr(0)
        End If
    Loop
    Call FindClose(iSearchHandle)
    GetFiles = sBuff
End If
End Function
Public Sub CopyFiles(hWind As Long, sSource As String, sTarget As String)

Dim SHFO As SHFILEOPSTRUCT
Dim Ret As Long
Const FOF_SILENT As Long = &H4
Const FO_COPY As Long = &H2

SHFO.hWnd = hWind
SHFO.fFlags = FOF_SILENT
SHFO.pFrom = sSource
SHFO.pTo = sTarget
SHFO.wFunc = FO_COPY
Ret = SHFileOperation(SHFO)
End Sub
'GetIconFromFile Private Function
'Convert an icon handle into an IPictureDisp.
Private Function IconToPicture(hIcon As Long) As IPictureDisp
Dim cls_id As CLSID
Dim hRes As Long
Dim new_icon As TypeIcon
Dim lpUnk As IUnknown

With new_icon
    .cbSize = Len(new_icon)
    .picType = vbPicTypeIcon
    .hIcon = hIcon
End With
With cls_id
    .id(8) = &HC0
    .id(15) = &H46
End With
hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
If hRes = 0 Then Set IconToPicture = lpUnk
End Function
'GetIconFromFile Public Function
Public Function GetIcon(FileName As String, icon_size As Long) As IPictureDisp
'Large Icon=0,Small Icon=1
Dim Index As Integer
Dim hIcon As Long
Dim item_num As Long
Dim icon_pic As IPictureDisp
Dim sh_info As SHFILEINFO

SHGetFileInfo FileName, 0, sh_info, Len(sh_info), SHGFI_ICON + icon_size
hIcon = sh_info.hIcon
Set icon_pic = IconToPicture(hIcon)
Set GetIcon = icon_pic
End Function
