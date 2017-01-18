VERSION 5.00
Begin VB.Form frmCargar 
   BackColor       =   &H8000000D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cargar mapa"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Hide Me!"
      Height          =   285
      Left            =   2565
      TabIndex        =   2
      Top             =   3195
      Width           =   825
   End
   Begin VB.ListBox MapLst 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2370
      Left            =   540
      TabIndex        =   1
      Top             =   675
      Width           =   2445
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mapas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1350
      TabIndex        =   0
      Top             =   225
      Width           =   780
   End
End
Attribute VB_Name = "frmCargar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Visible = False
End Sub

Private Sub MapLst_DblClick()

If MapInfo.Changed = 1 Then
    If MsgBox("Este mapa há sido modificado. Vas a perder todos los cambios si no lo grabas. Lo queres grabar ahora?", vbYesNo) = vbYes Then
        Call SaveMapData(CurMap)
    End If
End If


If MapLst.ListIndex <> -1 Then
    UserPos.X = (WindowTileWidth \ 2) + 1
    frmMain.HScroll1.value = UserPos.X
    UserPos.Y = (WindowTileHeight \ 2) + 1
    frmMain.VScroll1.value = UserPos.Y
    Call SwitchMap(MapLst.ListIndex + 1)
    EngineRun = True
    
Else
    MsgBox ("Elija un mapa.")
End If

End Sub
