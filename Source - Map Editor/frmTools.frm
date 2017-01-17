VERSION 5.00
Begin VB.Form frmTools 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tools"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   330
      Left            =   3870
      TabIndex        =   1
      Top             =   675
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar un BMP grande"
      Height          =   330
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   3840
   End
   Begin VB.Shape Shape1 
      Height          =   510
      Left            =   135
      Top             =   90
      Width           =   4335
   End
End
Attribute VB_Name = "frmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If frmOptions.Check3 = vbChecked Then
    If Val(frmAdmin.Text1(1).Text) > 1 Then
        MsgBox "No puedo tomar parametros desde una animacion."
        Exit Sub
    End If
    frmAddBigBMP.Text1(0).Text = frmAdmin.Text1(0).Text
    frmAddBigBMP.Text1(1).Text = frmAdmin.Text1(7).Text
    frmAddBigBMP.Text1(2).Text = Right(frmAdmin.XDIM.Caption, Len(frmAdmin.XDIM.Caption) - 6)
    frmAddBigBMP.Text1(3).Text = Right(frmAdmin.YDIM.Caption, Len(frmAdmin.YDIM.Caption) - 7)
End If


Me.Visible = False
frmAddBigBMP.AutoList.Clear
frmAddBigBMP.Visible = True
End Sub

Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Form_Deactivate()

If Me.Visible Then Me.SetFocus

End Sub

