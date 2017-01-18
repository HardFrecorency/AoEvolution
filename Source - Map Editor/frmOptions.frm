VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox MiLastGrh 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   2100
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   5325
      Begin VB.CheckBox Check5 
         Caption         =   "Auto Update LASTGRH"
         Height          =   330
         Left            =   180
         TabIndex        =   9
         Top             =   1650
         Width           =   5100
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ir a la Descripcion despues del click en la lista."
         Height          =   420
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   4920
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Arcualiza el ""Grh Fisico"" con clicks en la lista de bmps."
         Height          =   420
         Left            =   180
         TabIndex        =   4
         Top             =   675
         Width           =   4785
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Automate sub toma los parametros del grafico seleccionado."
         Height          =   375
         Left            =   180
         TabIndex        =   3
         Top             =   990
         Width           =   5100
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Actualizar automaticamente la base de datas del editor de mapas"
         Height          =   330
         Left            =   180
         TabIndex        =   2
         Top             =   1320
         Value           =   1  'Checked
         Width           =   5100
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   825
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   120
      Top             =   2490
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Grhs a indexar"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2730
      Width           =   2055
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Me.Visible = False
End Sub

Private Sub Command2_Click()
LastGrh = Val(MiLastGrh.Text)
End Sub

Private Sub Form_Deactivate()

If Me.Visible Then Me.SetFocus

End Sub

