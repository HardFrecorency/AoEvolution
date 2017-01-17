VERSION 5.00
Begin VB.Form frmAddBigBMP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Automated tool"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1170
      TabIndex        =   15
      Text            =   "0"
      Top             =   2655
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   2565
      TabIndex        =   14
      Text            =   "0"
      Top             =   2655
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1620
      TabIndex        =   13
      Top             =   585
      Width           =   1590
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Do it!"
      Height          =   285
      Left            =   1170
      TabIndex        =   11
      Top             =   3105
      Width           =   2445
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar a la lista de los tiles"
      Height          =   285
      Left            =   270
      TabIndex        =   10
      Top             =   5310
      Width           =   4425
   End
   Begin VB.ListBox AutoList 
      Height          =   1620
      Left            =   270
      TabIndex        =   9
      Top             =   3690
      Width           =   4425
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   8
      Top             =   1980
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   6
      Top             =   1260
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1170
      TabIndex        =   4
      Top             =   1980
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1170
      TabIndex        =   2
      Top             =   1305
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   285
      Left            =   4320
      TabIndex        =   0
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Offset X"
      Height          =   195
      Index           =   5
      Left            =   1215
      TabIndex        =   17
      Top             =   2385
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Offset Y"
      Height          =   195
      Index           =   4
      Left            =   2565
      TabIndex        =   16
      Top             =   2385
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion"
      Height          =   195
      Left            =   1935
      TabIndex        =   12
      Top             =   315
      Width           =   840
   End
   Begin VB.Shape Shape1 
      Height          =   3300
      Left            =   990
      Top             =   180
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BMP Height"
      Height          =   195
      Index           =   3
      Left            =   2610
      TabIndex        =   7
      Top             =   1710
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BMP Width"
      Height          =   195
      Index           =   2
      Left            =   2610
      TabIndex        =   5
      Top             =   990
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Grh Logico Inicio"
      Height          =   195
      Index           =   1
      Left            =   1125
      TabIndex        =   3
      Top             =   1710
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Grh Fisico"
      Height          =   195
      Index           =   0
      Left            =   1305
      TabIndex        =   1
      Top             =   1035
      Width           =   705
   End
End
Attribute VB_Name = "frmAddBigBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer
Dim Y As Integer
Private Sub Command1_Click()
Me.Visible = False
End Sub

Function DataOk() As Boolean
DataOk = Text1(0).Text <> "" And Text1(1).Text <> "" _
And Text1(2).Text <> "" And Text1(3).Text <> ""
End Function




Private Sub Command2_Click()

If AutoList.ListCount > 0 Then
    Dim j As Integer
    For j = 0 To AutoList.ListCount - 1
        frmAdmin.TXTList.AddItem AutoList.List(j)
    Next j
    Call BuscarDescrip(UCase(Text1(4).Text & 1))
    Call AddBD(ReadField(1, AutoList.List(0), 45), ReadField(2, AutoList.List(0), 45), X, Y)
    frmAdmin.MiLastGrh.Caption = "LastGrh:" & BiggestVal() - 1
    'Call AddBD(ReadField(1, cad, 45), Str(Val(log)), CantX, CantY)
End If
End Sub

Private Sub Command3_Click()
If DataOk Then
    Call Automate
Else
    MsgBox "Datos invalidos."
End If

End Sub

Private Sub Form_Deactivate()
If Me.Visible Then Me.SetFocus
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If Index <> 4 And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    Else
        Changed = True
    End If
End If
End Sub

Private Sub Automate()
'Cool feature coded!
Dim FileNum As String
Dim NumFrames As String
Dim sX As Integer
Dim sY As Integer
Dim pW As String
Dim pH As String
Dim LoopX As Integer
Dim LoopY As Integer
Dim log As Integer
Dim CantX As Integer
Dim CantY As Integer
Dim Cant As Integer
Dim cad As String
Dim OffSetX As Integer
Dim OffSetY As Integer

OffSetX = Val(Text1(6).Text)
OffSetY = Val(Text1(5).Text)


AutoList.Clear
FileNum = Text1(0).Text
NumFrames = 1
sX = 0 + OffSetX
sY = 0 + OffSetY
Cant = 1
log = Val(Text1(1).Text)

CantY = Val(Text1(3).Text) / 32
If CantY <> 0 Then
    If Val(Text1(3).Text) Mod 32 <> 0 Then MsgBox "El tamaño en Y del bmp no es multiplo de 32, posiblemente esto le ocasionara problemas"
End If

CantX = Val(Text1(2).Text) / 32
If CantX <> 0 Then
    If Val(Text1(2).Text) Mod 32 <> 0 Then MsgBox "El tamaño en X del bmp no es multiplo de 32, posiblemente esto le ocasionara problemas"
End If
X = CantX
Y = CantY
For LoopY = 1 To CantY
    For LoopX = 1 To CantX
            cad = Text1(4).Text & Cant & "-" & "Grh" & log & "-" _
            & NumFrames _
            & "-" & FileNum _
            & "-" & sX _
            & "-" & sY _
            & "-" & 32 & "-" & 32
            sX = (32 * LoopX) + OffSetX
            log = log + 1
            AutoList.AddItem cad
            
            Cant = Cant + 1
    Next
    sY = (32 * LoopY) + OffSetY
    sX = 0 + OffSetX
Next

End Sub
