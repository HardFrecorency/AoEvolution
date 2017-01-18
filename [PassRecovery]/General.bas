Attribute VB_Name = "General"

Option Explicit

Const CurrentModule As String = "modMain"
Public endc

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Public Const PROMPT_NONE As Long = 1
Public Const MIME_TYPE As Long = 2

Public Type typeUser
    ID As Long
    Name As String
    Email As String
    Password As String
    SMTP As String
    POP3 As String
End Type

Public Type typeMail
    ID As Long
    'Folder As Long
    'From As String
    To As String
    Subject As String
    'Date As Date
    Read As Boolean
    'Header As String
    Body As String
    'Boundary As String
    'Attachments As Long
End Type

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Type typeRule
    PartID As Long
    FindPhrase As String
    FolderID As Long
End Type

Public Rules() As typeRule

Public User As typeUser





Public Declare Function GetEncodedFile Lib "DECENC32.DLL" (ByVal strOutFile As String, ByVal nIndex As Long) As Long
Public Declare Function DecodeFile Lib "DECENC32.DLL" (ByVal strInFile As String, ByVal strOutFile As String, ByVal nPrompts As Long) As Long
Public Declare Sub SetEncodingApplication Lib "DECENC32.DLL" (ByVal strInFile As String)
Public Declare Function EncodeFile Lib "DECENC32.DLL" (ByVal SourceFile As String, ByVal EncodedFile As String, ByVal strBoundary As String, ByVal CodeOption As Long, ByVal xAppend As Long) As Long
Public Declare Sub FinishAttachments Lib "DECENC32.DLL" (ByVal strFileOut As String)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Sub Main()
    On Error GoTo Err_Init
    endc = Chr(1)
    frmMain.Socket1.AddressFamily = AF_INET
    frmMain.Socket1.Protocol = IPPROTO_IP
    frmMain.Socket1.SocketType = SOCK_STREAM
    frmMain.Socket1.Binary = False
    frmMain.Socket1.Blocking = False
    frmMain.Socket1.BufferSize = 1024
    
    frmMain.Socket2(0).AddressFamily = AF_INET
    frmMain.Socket2(0).Protocol = IPPROTO_IP
    frmMain.Socket2(0).SocketType = SOCK_STREAM
    frmMain.Socket2(0).Binary = False
    frmMain.Socket2(0).Blocking = False
    frmMain.Socket2(0).BufferSize = 2048
    
    frmMain.Socket1.LocalPort = 7668
    frmMain.Socket1.Listen
    
    frmMain.Show

    Exit Sub

Err_Init:

End Sub

Public Function MD5String(p As String) As String
' compute MD5 digest on a given string, returning the result
    Dim r As String * 32, t As Long
    r = Space(32)
    t = Len(p)
    MDStringFix p, t, r
    MD5String = r
End Function

Public Function MD5File(f As String) As String
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function

