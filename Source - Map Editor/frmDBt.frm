VERSION 5.00
Begin VB.Form frmDBt 
   Caption         =   "DB MIXER"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   345
      Left            =   495
      TabIndex        =   5
      Top             =   1230
      Width           =   3330
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   2130
      TabIndex        =   2
      Top             =   360
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   525
      TabIndex        =   0
      Top             =   375
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   1695
      TabIndex        =   4
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta:"
      Height          =   300
      Left            =   2145
      TabIndex        =   3
      Top             =   105
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Desde:"
      Height          =   300
      Left            =   540
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "frmDBt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conexion As Connection
Dim Conexion2 As Connection


Private Sub AbrirBD()
Err.Clear
On Error GoTo fin
Set Conexion = New Connection
Conexion.Provider = "Microsoft.Jet.OLEDB.3.51"
Conexion.ConnectionString = "Data Source=" & App.Path & "\grhindex\grhindex.mdb"
Conexion.Open

Set Conexion2 = New Connection
Conexion2.Provider = "Microsoft.Jet.OLEDB.3.51"
Conexion2.ConnectionString = "Data Source=" & App.Path & "\grhindex\grhindex2.mdb"
Conexion2.Open
Exit Sub

fin:
If Err Then
    MsgBox "No se puede abrir la base de datos." & Err.Description
    End
End If

End Sub


Private Sub Posicionar(n As Integer)
Dim i As Integer

Dim rs As New Recordset
Dim rs2 As New Recordset
i = 0
rs.Open "GrhIndex", Conexion2, , adLockReadOnly, adCmdTable

Do While i < n
    i = i + 1
    Label3.Caption = rs!Nombre
    rs.MoveNext
Loop
rs2.Open "GrhIndex", Conexion, , adLockOptimistic, adCmdTable
Do While Not rs.EOF
    rs2.AddNew
    rs2!Nombre = rs!Nombre
    rs2!GrhIndice = rs!GrhIndice
    rs2!Ancho = rs!Ancho
    rs2!alto = rs!alto
    rs2.Update
    rs2.MoveNext
    rs.MoveNext
Loop
    
 

rs.Close
rs2.Close
Set rs = Nothing
Set rs2 = Nothing
End Sub








Private Sub Command1_Click()
Call Posicionar(Val(Text1.Text))
End Sub

Private Sub Form_Load()
AbrirBD
End Sub
