VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmCargando"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar Progreso 
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Top             =   945
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Contador 
      Caption         =   "Cargando base de datos, por favor espere..."
      Height          =   285
      Left            =   630
      TabIndex        =   1
      Top             =   225
      Width           =   3345
   End
End
Attribute VB_Name = "frmProgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_LostFocus()
If frmProgreso.Visible Then frmProgreso.SetFocus
End Sub
