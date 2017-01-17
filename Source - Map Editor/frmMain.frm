VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor"
   ClientHeight    =   8280
   ClientLeft      =   390
   ClientTop       =   960
   ClientWidth     =   11895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   552
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1320
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   11670
      Begin VB.CheckBox Completar 
         BackColor       =   &H8000000D&
         Caption         =   "AutoCompletar"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1800
         TabIndex        =   37
         Top             =   675
         Width           =   1590
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
         Height          =   240
         Index           =   3
         Left            =   4995
         TabIndex        =   34
         Top             =   900
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-"
         Height          =   240
         Index           =   2
         Left            =   3960
         TabIndex        =   33
         Top             =   900
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-"
         Height          =   240
         Index           =   1
         Left            =   3960
         TabIndex        =   32
         Top             =   495
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
         Height          =   240
         Index           =   0
         Left            =   4995
         TabIndex        =   31
         Top             =   495
         Width           =   240
      End
      Begin VB.TextBox DMLargo 
         Height          =   330
         Left            =   4185
         TabIndex        =   30
         Text            =   "0"
         Top             =   855
         Width           =   780
      End
      Begin VB.TextBox DMAncho 
         Height          =   330
         Left            =   4185
         TabIndex        =   29
         Text            =   "0"
         Top             =   450
         Width           =   780
      End
      Begin VB.CheckBox DespMosaic 
         BackColor       =   &H8000000D&
         Caption         =   "DespMosaico"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   3915
         TabIndex        =   28
         Top             =   90
         Width           =   1320
      End
      Begin VB.CommandButton Command5 
         Caption         =   "D&esbloquear mapa"
         Height          =   375
         Left            =   7560
         TabIndex        =   27
         Top             =   495
         Width           =   1500
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Bl&oquear mapa"
         Height          =   375
         Left            =   7560
         TabIndex        =   26
         Top             =   90
         Width           =   1500
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   5715
         TabIndex        =   25
         Top             =   585
         Width           =   1635
      End
      Begin VB.TextBox mLargo 
         Height          =   330
         Left            =   360
         TabIndex        =   23
         Text            =   "1"
         Top             =   855
         Width           =   1140
      End
      Begin VB.TextBox mAncho 
         Height          =   330
         Left            =   360
         TabIndex        =   22
         Text            =   "1"
         Top             =   270
         Width           =   1140
      End
      Begin VB.CheckBox MOSAICO 
         BackColor       =   &H8000000D&
         Caption         =   "Mosaico"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1800
         TabIndex        =   21
         Top             =   270
         Width           =   1230
      End
      Begin VB.TextBox StatTxt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1035
         Left            =   9720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "frmMain.frx":030A
         Top             =   225
         Width           =   1920
      End
      Begin VB.CommandButton Command2 
         Caption         =   "H&erramientas"
         Height          =   375
         Left            =   7560
         TabIndex        =   1
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Alto"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   405
         TabIndex        =   39
         Top             =   675
         Width           =   270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   38
         Top             =   90
         Width           =   465
      End
      Begin VB.Shape Shape4 
         BorderUidth     =   2
         Height          =   1230
         Left            =   5625
         Top             =   45
         Width           =   1830
      End
      Begin VB.Shape Shape3 
         BorderUidth     =   2
         Height          =   1230
         Left            =   3735
         Top             =   45
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del mapa"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5895
         TabIndex        =   24
         Top             =   315
         Width           =   1245
      End
      Begin VB.Shape Shape6 
         BorderUidth     =   2
         Height          =   1230
         Left            =   180
         Top             =   45
         Width           =   3405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Info:"
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
         Left            =   9720
         TabIndex        =   3
         Top             =   0
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   2670
      Left            =   90
      TabIndex        =   11
      Top             =   5490
      Width           =   2085
      Begin VB.CommandButton PlaceGrhCmd 
         Caption         =   "Poner Grh"
         Enabled         =   0   'False
         Height          =   255
         Left            =   195
         TabIndex        =   18
         Top             =   2295
         Width           =   1515
      End
      Begin VB.TextBox Grhtxt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   195
         TabIndex        =   17
         Text            =   "1"
         Top             =   390
         Width           =   795
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Up"
         Height          =   255
         Left            =   1035
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Down"
         Height          =   255
         Left            =   1035
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   630
         Width           =   615
      End
      Begin VB.TextBox Layertxt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   195
         TabIndex        =   14
         Text            =   "1"
         Top             =   1170
         Width           =   555
      End
      Begin VB.CheckBox EraseAllchk 
         BackColor       =   &H8000000D&
         Caption         =   "Borrar todo"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   225
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2025
         Width           =   1335
      End
      Begin VB.CheckBox Erasechk 
         BackColor       =   &H8000000D&
         Caption         =   "Borrar Layer"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   225
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1755
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BorderUidth     =   2
         Height          =   2535
         Left            =   135
         Top             =   90
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Layer"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   195
         TabIndex        =   20
         Top             =   960
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Grh"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   195
         TabIndex        =   19
         Top             =   150
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4020
      Left            =   90
      TabIndex        =   4
      Top             =   1440
      Width           =   2085
      Begin VB.ListBox lCelda 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1620
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   1620
      End
      Begin VB.CheckBox WalkModeChk 
         BackColor       =   &H8000000D&
         Caption         =   "Modo Interactivo"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   150
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000D&
         Caption         =   "Mostrar Celdas Bloqueadas"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   150
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2205
         Width           =   1815
      End
      Begin VB.CheckBox DrawGridChk 
         BackColor       =   &H8000000D&
         Caption         =   "Grilla"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   150
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1890
         Width           =   1155
      End
      Begin VB.CommandButton PlaceBlockCmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cambiar bloqueado"
         Height          =   255
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3375
         Width           =   1515
      End
      Begin VB.CheckBox Blockedchk 
         BackColor       =   &H8000000D&
         Caption         =   "Bloqueado"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   315
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BorderUidth     =   2
         Height          =   930
         Left            =   135
         Top             =   3000
         Width           =   1725
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6225
      Left            =   10485
      Max             =   100
      Min             =   1
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1530
      Value           =   1
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   2295
      Max             =   100
      Min             =   1
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   7785
      Value           =   1
      Width           =   8160
   End
   Begin VB.Shape MainViewShp 
      Height          =   6240
      Left            =   2295
      Top             =   1485
      Width           =   8160
   End
   Begin VB.Menu FileMnu 
      Caption         =   "Archivo"
      Begin VB.Menu mnuCargar 
         Caption         =   "Cargar Mapa"
      End
      Begin VB.Menu SaveMnu 
         Caption         =   "Grabar"
      End
      Begin VB.Menu SaveNewMnu 
         Caption         =   "Grabar como mapa nuevo"
      End
      Begin VB.Menu nAbout 
         Caption         =   "Acerca de"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu OptionMnu 
      Caption         =   "Opciones"
      Begin VB.Menu ClsRoomMnu 
         Caption         =   "Borrar Mapa"
      End
      Begin VB.Menu ClsBordMnu 
         Caption         =   "Borrar Borde"
      End
      Begin VB.Menu mnuMusica 
         Caption         =   "Musica"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function DameGrhIndex(Nombre As String) As Integer
Dim rs As New Recordset
rs.Open "GrhIndex", Conexion, , adLockReadOnly, adCmdTable

Do While Not rs.EOF And rs!Nombre <> Nombre
    rs.MoveNext
Loop
DameGrhIndex = rs!grhindice
If rs!ancho > 0 Then
    MOSAICO.value = vbChecked
    mAncho.Text = rs!ancho
    mLargo.Text = rs!alto
Else
    MOSAICO.value = vbUnchecked
    mAncho.Text = ""
    mLargo.Text = ""
End If
        
rs.Close

End Function


Private Sub Blockedchk_Click()

Call PlaceBlockCmd_Click

End Sub


Private Sub Check1_Click()

If DrawBlock = True Then
    DrawBlock = False
Else
    DrawBlock = True
End If

End Sub

Private Sub Command4_Click()
If MsgBox("Cuidado, con este comando podes arruinar el mapa.¿Estas seguro que queres hacer esto?", vbYesNo) = vbNo Then
        Exit Sub
End If
Dim Y As Integer
Dim X As Integer

If CurMap = 0 Then
    Exit Sub
End If

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        MapData(X, Y).Blocked = 1
    Next X
Next Y

MapInfo.Changed = 1
End Sub

Private Sub Command5_Click()
If MsgBox("Cuidado, con este comando podes arruinar el mapa.¿Estas seguro que queres hacer esto?", vbYesNo) = vbNo Then
        Exit Sub
End If
Dim Y As Integer
Dim X As Integer

If CurMap = 0 Then
    Exit Sub
End If

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        MapData(X, Y).Blocked = 0
    Next X
Next Y

MapInfo.Changed = 1
End Sub

Private Sub Command6_Click(Index As Integer)
On Error Resume Next
Select Case Index
        Case 0
            DMAncho.Text = Str(Val(DMAncho.Text) + 1)
        Case 1
            DMAncho.Text = Str(Val(DMAncho.Text) - 1)
        Case 2
            DMLargo.Text = Str(Val(DMLargo.Text) - 1)
        Case 3
            DMLargo.Text = Str(Val(DMLargo.Text) + 1)
End Select
End Sub

Private Sub DespMosaic_Click()
If DMAncho.Text = "" Then DMAncho.Text = "0"
If DMLargo.Text = "" Then DMLargo.Text = "0"
End Sub

Private Sub lCelda_Click()
Grhtxt.Text = DameGrhIndex(lCelda.List(lCelda.ListIndex))
If frmGrafico.Visible = False Then frmGrafico.Visible = True
Call PlaceGrhCmd_Click
End Sub

Private Sub mnuCargar_Click()
frmCargar.Visible = True
End Sub

Private Sub mnuMusica_Click()
frmMusica.Show
End Sub

Private Sub MOSAICO_Click()
If mAncho.Text = "" Then mAncho.Text = "1"
If mLargo.Text = "" Then mLargo.Text = "1"
End Sub

Public Sub PlaceBlockCmd_Click()

PlaceGrhCmd.Enabled = True
PlaceBlockCmd.Enabled = False
frmHerramientas.PlaceExitCmd.Enabled = True
frmHerramientas.PlaceNPCHOSTCmd.Enabled = True
frmHerramientas.PlaceNPCCmd.Enabled = True
frmHerramientas.PlaceObjCmd.Enabled = True

End Sub


Private Sub Grhtxt_Change()

If Val(Grhtxt.Text) < 1 Then
  Grhtxt.Text = NumGrhs
  Exit Sub
End If

If Val(Grhtxt.Text) > NumGrhs Then
  Grhtxt.Text = 1
  Exit Sub
End If

'Change CurrentGrh
CurrentGrh.grhIndex = Val(Grhtxt.Text)
CurrentGrh.Started = 1
CurrentGrh.FrameCounter = 1
CurrentGrh.SpeedCounter = GrhData(CurrentGrh.grhIndex).Speed

End Sub

Private Sub ClsBordMnu_Click()

'*****************************************************************
'Clears a border in a room with current GRH
'*****************************************************************

Dim Y As Integer
Dim X As Integer

If CurMap = 0 Then
    Exit Sub
End If

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then

          If frmMain.MOSAICO.value = vbChecked Then
            Dim aux As Integer
            aux = Val(frmMain.Grhtxt.Text) + _
            ((Y Mod frmMain.mLargo) * frmMain.mAncho) + (X Mod frmMain.mAncho)
             MapData(X, Y).Blocked = frmMain.Blockedchk.value
             MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)).grhIndex = aux
            'Setup GRH
            InitGrh MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)), aux
          Else
            'Else Place graphic
            MapData(X, Y).Blocked = frmMain.Blockedchk.value
            MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)).grhIndex = Val(frmMain.Grhtxt.Text)
            
            'Setup GRH
    
            InitGrh MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)), Val(frmMain.Grhtxt.Text)
        End If
             'Erase NPCs
            If MapData(X, Y).NPCIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NPCIndex = 0
            End If

            'Erase Objs
            MapData(X, Y).OBJInfo.OBJIndex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.grhIndex = 0

            'Clear exits
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0

        End If

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

Private Sub ClsRoomMnu_Click()
'*****************************************************************
'Clears all layers
'*****************************************************************

Dim Y As Integer
Dim X As Integer

If CurMap = 0 Then
    Exit Sub
End If

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        'Change blockes status
        MapData(X, Y).Blocked = frmMain.Blockedchk.value

        'Erase layer 2 and 3
        MapData(X, Y).Graphic(2).grhIndex = 0
        MapData(X, Y).Graphic(3).grhIndex = 0
        MapData(X, Y).Graphic(4).grhIndex = 0

        'Erase NPCs
        If MapData(X, Y).NPCIndex > 0 Then
            EraseChar MapData(X, Y).CharIndex
            MapData(X, Y).NPCIndex = 0
        End If

        'Erase Objs
        MapData(X, Y).OBJInfo.OBJIndex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.grhIndex = 0

        'Clear exits
        MapData(X, Y).TileExit.Map = 0
        MapData(X, Y).TileExit.X = 0
        MapData(X, Y).TileExit.Y = 0

        If frmMain.MOSAICO.value = vbChecked Then
            Dim aux As Integer
            aux = Val(frmMain.Grhtxt.Text) + _
            ((Y Mod frmMain.mLargo) * frmMain.mAncho) + (X Mod frmMain.mAncho)
             MapData(X, Y).Blocked = frmMain.Blockedchk.value
             MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)).grhIndex = aux
            'Setup GRH
            InitGrh MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)), aux
        Else
            'Else Place graphic
            MapData(X, Y).Blocked = frmMain.Blockedchk.value
            MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)).grhIndex = Val(frmMain.Grhtxt.Text)
            
            'Setup GRH
    
            InitGrh MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)), Val(frmMain.Grhtxt.Text)
        End If

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub





Private Sub Command2_Click()
If frmHerramientas.Visible Then frmHerramientas.SetFocus _
Else: frmHerramientas.Visible = Not frmHerramientas.Visible
End Sub









Private Sub DrawGridChk_Click()

If DrawGrid = True Then
    DrawGrid = False
Else
    DrawGrid = True
End If

End Sub

Private Sub EraseAllchk_Click()
Call PlaceGrhCmd_Click
Erasechk.value = False
End Sub

Private Sub Erasechk_Click()

'Set Place GRh mode
Call PlaceGrhCmd_Click

EraseAllchk.value = False

End Sub

Private Sub EraseExitChk_Click()

Call frmHerramientas.PlaceExitCmd_Click

End Sub

Private Sub EraseNPCChk_Click()

Call frmHerramientas.PlaceNPCCmd_Click

End Sub

Private Sub EraseObjChk_Click()

Call frmHerramientas.PlaceObjCmd_Click

End Sub

Private Sub HScroll1_Change()

If EngineRun Then
    If WalkMode = True Then
            MsgBox "Para moverse con las barras desactive el modo interactivo"
    Else
            UserPos.X = HScroll1.value
    End If
End If

End Sub

Private Sub HScroll1_Scroll()
If EngineRun Then
    If WalkMode = True Then
            MsgBox "Para moverse con las barras desactive el modo interactivo"
    Else
            UserPos.X = HScroll1.value
'            HScroll1.Refresh
    End If
End If
End Sub



Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Grh As Integer

'Set Place GRh mode
Call PlaceGrhCmd_Click

Grh = Val(Grhtxt.Text)

'Add to current Grh number
If Button = vbLeftButton Then
   Grh = Grh - 1
End If

If Button = vbRightButton Then
    Grh = Grh - 10
End If

'Update Grhtxt
Grhtxt.Text = Grh
Grh = Val(Grhtxt)

'If blank find next valid Grh
If GrhData(Grh).NumFrames = 0 Then
    
    Do Until GrhData(Grh).NumFrames > 0
        Grh = Grh - 1
        If Grh < 1 Then
            Grh = NumGrhs
        End If
    Loop
    
End If

'Update Grhtxt
Grhtxt.Text = Grh

End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Grh As Integer

'Set Place GRh mode
Call PlaceGrhCmd_Click

Grh = Val(Grhtxt.Text)

'Add to current Grh number
If Button = vbLeftButton Then
   Grh = Grh + 1
End If

If Button = vbRightButton Then
    Grh = Grh + 10
End If

'Update Grhtxt
Grhtxt.Text = Grh
Grh = Val(Grhtxt)

'If blank find next valid Grh
If GrhData(Grh).NumFrames = 0 Then
    
    Do Until GrhData(Grh).NumFrames > 0
        Grh = Grh + 1
        If Grh > NumGrhs Then
            Grh = 1
        End If
    Loop
    
End If

'Update Grhtxt
Grhtxt.Text = Grh

End Sub

Private Sub Layertxt_Change()

If Val(Layertxt.Text) < 1 Then
  Layertxt.Text = 1
End If

If Val(Layertxt.Text) > 4 Then
  Layertxt.Text = 4
End If

Call PlaceGrhCmd_Click

End Sub




Private Sub Form_Load()
frmMain.Caption = frmMain.Caption & " V " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim tX As Integer
Dim tY As Integer

If CurMap <= 0 Then Exit Sub

If X <= MainViewShp.Left Or X >= MainViewShp.Left + MainViewWidth Or Y <= MainViewShp.Top Or Y >= MainViewShp.Top + MainViewHeight Then
    Exit Sub
End If

ConvertCPtoTP MainViewShp.Left, MainViewShp.Top, X, Y, tX, tY

ReacttoMouseClick Button, tX, tY

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim tX As Integer
Dim tY As Integer

'Make sure map is loaded
If CurMap <= 0 Then Exit Sub

'Make sure click is in view window
If X <= MainViewShp.Left Or X >= MainViewShp.Left + MainViewWidth Or Y <= MainViewShp.Top Or Y >= MainViewShp.Top + MainViewHeight Then
    Exit Sub
End If

ConvertCPtoTP MainViewShp.Left, MainViewShp.Top, X, Y, tX, tY

ReacttoMouseClick Button, tX, tY

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Allow MainLoop to close program
If prgRun = True Then
    prgRun = False
    Cancel = 1
End If

End Sub

Private Sub nAbout_Click()
frmAbout.Show
End Sub

Public Sub PlaceGrhCmd_Click()
PlaceGrhCmd.Enabled = False
PlaceBlockCmd.Enabled = True
frmHerramientas.PlaceExitCmd.Enabled = True
frmHerramientas.PlaceNPCCmd.Enabled = True
frmHerramientas.PlaceObjCmd.Enabled = True

End Sub

Private Sub Salir_Click()
Unload Me
End Sub

Private Sub SaveMnu_Click()

If CurMap = 0 Then
    Exit Sub
End If

Call SaveMapData(CurMap)

'Set changed flag
MapInfo.Changed = 0

End Sub


Private Sub SaveNewMnu_Click()

If CurMap = 0 Then
    Exit Sub
End If

NumMaps = NumMaps + 1
Call SaveMapData(NumMaps)
frmCargar.MapLst.AddItem "Map " & NumMaps, NumMaps - 1

End Sub


Private Sub Text1_Change()
MapInfo.Name = Text1.Text
End Sub

Private Sub VScroll1_Change()
If EngineRun Then
    If WalkMode = True Then
            MsgBox "Para moverse con las barras desactive el modo interactivo"
    Else
            UserPos.Y = VScroll1.value
    End If
End If

End Sub

Private Sub WalkModeChk_Click()

ToggleWalkMode

End Sub
