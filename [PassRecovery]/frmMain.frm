VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Password Recovery Server"
   ClientHeight    =   2205
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   2595
   StartUpPosition =   3  'Windows Default
   Begin SocketWrenchCtrl.Socket Socket2 
      Index           =   0
      Left            =   675
      Top             =   105
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   225
      Top             =   120
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Busca los chr en el directorio \charfile\
'usa el spool del Win2k para envíar los passwords por correo electronico

'Otto se ocupó de que funcionará bien, thanks Otto ;)

Dim sRes(0 To 500) As String

Private Sub EnviarPasswd(ByVal mail As String, ByVal Passwd As String)
Dim SMTPFileIndex As Long
SMTPFileIndex = 1
    Do While FileExist("C:\inetpub\mailroot\pickup\Argentum." & SMTPFileIndex & ".eml", vbNormal)
        SMTPFileIndex = SMTPFileIndex + 1
        If SMTPFileIndex > 200 Then
            Exit Sub
        End If
    Loop
    Open "C:\inetpub\mailroot\pickup\Argentum." & SMTPFileIndex & ".eml" For Append As #1
    Print #1, ("X-Sender: noreply@noreply.com.ar")
    Print #1, ("X-Receiver: " & mail)
    Print #1, ("From: ""Gulfas Morgolock"" <noreply@noreply.com.ar>")
    Print #1, ("To: " & mail)
    Print #1, ("Subject: -*- Password -*-")
    Print #1, ("Date: " & WeekdayName(Weekday(Date), True) & ", " & Day(Date) & " " & MonthName(Month(Date), True) & " " & Year(Date) & " " & Time & " -0300")
    Print #1, ("MIME-Version: 1.0")
    Print #1, ("Content-Type: text/plain;")
    Print #1, (Chr$(9) & "charset=""iso-8859-1""")
    Print #1, ("Content-Transfer-Encoding: 7bit")
    Print #1, ("")
    Print #1, "Su password es:" & Passwd & vbCrLf & "Le recomendamos que cambien el password de su personaje periodicamente y no use passwords muy cortos para aumentar la seguridad." & _
    vbCrLf & "Atentamente," & vbCrLf & "El staff de Noland Studios."
    Close #1
End Sub

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function

Private Function GenerarPass() As String

Dim temp As String
Dim i As Integer

For i = 1 To 8
    temp = temp & Chr(RandomNumber(Asc("a"), Asc("z")))
Next i

GenerarPass = temp

End Function


Private Sub Form_Load()

'With mapSess
'    .SignOn
'    mapMess.SessionID = .SessionID
'    .DownLoadMail = True
'End With
    
    'Lets put an icon in the system tray
    With sysIcon
        .cbSize = LenB(sysIcon)
        .hWnd = Me.hWnd
        .uFlags = NIF_DOALL
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .sTip = "PasswordRecovery Server" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, sysIcon
    Me.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ShellMsg As Long
    
    ShellMsg = X / Screen.TwipsPerPixelX
    Select Case ShellMsg
    Case WM_LBUTTONDBLCLK
        Me.Visible = True
    End Select
End Sub

Private Sub mnuCerrar_Click()
End
End Sub

Private Sub mnuSystray_Click()
Me.Visible = False
End Sub

Private Sub Socket1_Accept(SocketId As Integer)
    Dim Index As Integer
    Index = Socket2.UBound + 1
    Load Socket2(Index)
    Socket2(Index).AddressFamily = AF_INET
    Socket2(Index).Protocol = IPPROTO_IP
    Socket2(Index).SocketType = SOCK_STREAM
    Socket2(Index).Binary = False
    Socket2(Index).BufferSize = 2048
    Socket2(Index).Blocking = False
    Socket2(Index).Accept = SocketId
    Socket2(Index).Linger = 1
End Sub
Function GetVar(file As String, Main As String, Var As String) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space(5000) ' This tells the computer how long the longest string can be
  
  
getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file
  
GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)
  
End Function

Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Devuelve el string del campo
'*****************************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = Mid(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i

FieldNum = FieldNum + 1
If FieldNum = Pos Then
    ReadField = Mid(Text, LastPos + 1)
End If

End Function

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Private Sub HandleData(ByVal Data As String, ByVal Index As Integer)
On Error Resume Next

If Left$(Data, 8) = "PASSRECO" Then
    Dim Pass$, mail$, nombre$, file$
        
    Data = Right$(Data, Len(Data) - 8)
    nombre$ = ReadField(1, Data, 126)
    mail$ = ReadField(2, Data, 126)
    If Not AsciiValidos(nombre$) Then Call CloseSocket(Index)
    file$ = App.Path & "\charfile\" & nombre$ & ".chr"
    
    If FileExist(file$, vbNormal) Then
        '¿El mail es valido?
        If mail$ = GetVar(file$, "CONTACTO", "Email") Then
                   Pass$ = GenerarPass
                   Call WriteVar(file$, "INIT", "Password", MD5String(Pass$))
                   Call EnviarPasswd(mail$, Pass$)
                   'Call SendData("!!El nuevo password es " & Pass$, Index)
                   Call SendData("RECPASSOK", Index)
        Else
                   Call SendData("RECPASSER", Index)
        End If
        
    End If
    
End If


End Sub

Private Sub CloseSocket(ByVal Index As Integer)
Socket2(Index).Cleanup
Unload Socket2(Index)
End Sub

Function FileExist(file As String, FileType As VbFileAttribute) As Boolean

On Error Resume Next

If Dir(file, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If

End Function

Private Sub Socket2_Disconnect(Index As Integer)
Call CloseSocket(Index)
End Sub

Private Sub Socket2_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
On Error Resume Next
Dim Data As String
Socket2(Index).Read Data, DataLength
Call HandleData(Data, Index)
End Sub

Private Sub SendData(sndData As String, Index As Integer)
On Error Resume Next
sndData = sndData & endc
Socket2(Index).Write sndData, Len(sndData)
End Sub

Sub WriteVar(file As String, Main As String, Var As String, Value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, Value, file
    
End Sub



