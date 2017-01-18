VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor"
   ClientHeight    =   10635
   ClientLeft      =   390
   ClientTop       =   960
   ClientWidth     =   14040
   Icon            =   "frmMain2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   709
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   936
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5355
      Left            =   60
      ScaleHeight     =   355
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   295
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   6375
      Width           =   4455
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   4140
         Left            =   0
         ScaleHeight     =   4080
         ScaleWidth      =   4350
         TabIndex        =   40
         Top             =   30
         Width           =   4410
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2565
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1770
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   13305
      Begin VB.CheckBox Check2 
         BackColor       =   &H8000000D&
         Caption         =   "Loop"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   5235
         TabIndex        =   45
         Top             =   1350
         Width           =   780
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   4350
         TabIndex        =   43
         Text            =   "1"
         Top             =   1320
         Width           =   750
      End
      Begin VB.TextBox Vers 
         Height          =   375
         Left            =   2115
         TabIndex        =   41
         Top             =   1365
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000007&
         Height          =   1680
         Left            =   8505
         ScaleHeight     =   108
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   99
         TabIndex        =   28
         Top             =   45
         Width           =   1545
         Begin VB.Label Apuntador 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "O"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   30
            TabIndex        =   29
            Top             =   0
            Width           =   120
         End
      End
      Begin VB.CheckBox Completar 
         BackColor       =   &H8000000D&
         Caption         =   "AutoCompletar"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1695
         TabIndex        =   25
         Top             =   675
         Width           =   1350
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
         Height          =   240
         Index           =   3
         Left            =   4575
         TabIndex        =   24
         Top             =   885
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-"
         Height          =   240
         Index           =   2
         Left            =   3540
         TabIndex        =   23
         Top             =   885
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-"
         Height          =   240
         Index           =   1
         Left            =   3540
         TabIndex        =   22
         Top             =   480
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
         Height          =   240
         Index           =   0
         Left            =   4575
         TabIndex        =   21
         Top             =   480
         Width           =   240
      End
      Begin VB.TextBox DMLargo 
         Height          =   330
         Left            =   3765
         TabIndex        =   20
         Text            =   "0"
         Top             =   840
         Width           =   780
      End
      Begin VB.TextBox DMAncho 
         Height          =   330
         Left            =   3765
         TabIndex        =   19
         Text            =   "0"
         Top             =   435
         Width           =   780
      End
      Begin VB.CheckBox DespMosaic 
         BackColor       =   &H8000000D&
         Caption         =   "DespMosaico"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   3390
         TabIndex        =   18
         Top             =   105
         Width           =   1320
      End
      Begin VB.CommandButton Command5 
         Caption         =   "D&esbloquear mapa"
         Height          =   375
         Left            =   6960
         TabIndex        =   17
         Top             =   465
         Width           =   1500
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Bl&oquear mapa"
         Height          =   375
         Left            =   6960
         TabIndex        =   16
         Top             =   60
         Width           =   1500
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   5190
         TabIndex        =   15
         Top             =   585
         Width           =   1635
      End
      Begin VB.TextBox mLargo 
         Height          =   330
         Left            =   360
         TabIndex        =   13
         Text            =   "1"
         Top             =   855
         Width           =   1140
      End
      Begin VB.TextBox mAncho 
         Height          =   330
         Left            =   360
         TabIndex        =   12
         Text            =   "1"
         Top             =   270
         Width           =   1140
      End
      Begin VB.CheckBox MOSAICO 
         BackColor       =   &H8000000D&
         Caption         =   "Mosaico"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1695
         TabIndex        =   11
         Top             =   270
         Width           =   1230
      End
      Begin VB.TextBox StatTxt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1035
         Left            =   10200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "frmMain2.frx":030A
         Top             =   225
         Width           =   1440
      End
      Begin VB.CommandButton Command2 
         Caption         =   "H&erramientas"
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Top             =   870
         Width           =   1500
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MIDI"
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
         Left            =   3810
         TabIndex        =   44
         Top             =   1380
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version del mapa:"
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
         Left            =   180
         TabIndex        =   42
         Top             =   1440
         Width           =   1905
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Alto"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   405
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   90
         Width           =   465
      End
      Begin VB.Shape Shape4 
         Height          =   1230
         Left            =   5100
         Top             =   45
         Width           =   1830
      End
      Begin VB.Shape Shape3 
         Height          =   1230
         Left            =   3285
         Top             =   30
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del mapa"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5370
         TabIndex        =   14
         Top             =   315
         Width           =   1245
      End
      Begin VB.Shape Shape6 
         Height          =   1230
         Left            =   180
         Top             =   45
         Width           =   3030
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
         Left            =   10215
         TabIndex        =   3
         Top             =   15
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4605
      Left            =   75
      TabIndex        =   4
      Top             =   1755
      Width           =   4395
      Begin VB.CheckBox Check3 
         BackColor       =   &H8000000D&
         Caption         =   "Ver triggers"
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
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1815
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
         Left            =   2400
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3645
         Width           =   1335
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
         Left            =   2400
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3915
         Width           =   1335
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
         Left            =   2370
         TabIndex        =   34
         Text            =   "1"
         Top             =   3060
         Width           =   555
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Down"
         Height          =   255
         Left            =   3210
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Up"
         Height          =   255
         Left            =   3210
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2160
         Width           =   615
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
         Left            =   2370
         TabIndex        =   31
         Text            =   "1"
         Top             =   2280
         Width           =   795
      End
      Begin VB.CommandButton PlaceGrhCmd 
         Caption         =   "Poner Grh"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2370
         TabIndex        =   30
         Top             =   4185
         Width           =   1515
      End
      Begin VB.ListBox lCelda 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1620
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   195
         Width           =   4065
      End
      Begin VB.CheckBox Mostar4layer 
         BackColor       =   &H8000000D&
         Caption         =   "Cuarto layer"
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
         Top             =   3975
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
         Top             =   3720
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Grh"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2370
         TabIndex        =   38
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Layer"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2370
         TabIndex        =   37
         Top             =   2850
         Width           =   390
      End
      Begin VB.Shape Shape2 
         Height          =   2535
         Left            =   2310
         Top             =   1980
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         Height          =   930
         Left            =   135
         Top             =   3495
         Width           =   1725
      End
   End
   Begin VB.Shape MainViewShp 
      Height          =   6240
      Left            =   4500
      Top             =   1770
      Width           =   8160
   End
   Begin VB.Menu FileMnu 
      Caption         =   "Archivo"
      Begin VB.Menu mnuNuevo 
         Caption         =   "Nuevo mapa"
      End
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
      Caption         =   "Mapa"
      Begin VB.Menu borratri 
         Caption         =   "Borrar Triggers"
      End
      Begin VB.Menu Actmapas 
         Caption         =   "Actualizar mapas"
      End
      Begin VB.Menu mnuCarac 
         Caption         =   "Caracteristicas"
      End
      Begin VB.Menu ClsRoomMnu 
         Caption         =   "Borrar Mapa"
      End
      Begin VB.Menu ClsBordMnu 
         Caption         =   "Borrar Borde"
      End
      Begin VB.Menu mnuGrilla 
         Caption         =   "Grilla"
      End
      Begin VB.Menu mnuMusica 
         Caption         =   "Musica"
      End
      Begin VB.Menu mnuborrarArboles 
         Caption         =   "Borrar Arboles OBJ"
      End
      Begin VB.Menu mnuBloquear 
         Caption         =   "Bloquear Bordes"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
Option Explicit



Function DameGrhIndex(Nombre As String) As Integer
Dim rs As New Recordset
rs.Open "GrhIndex", Conexion, , adLockReadOnly, adCmdTable

Do While Not rs.EOF And rs!Nombre <> Nombre
    rs.MoveNext
Loop
DameGrhIndex = rs!GrhIndice
If rs!Ancho > 0 Then
    MOSAICO.value = vbChecked
    mAncho.Text = rs!Ancho
    mLargo.Text = rs!Alto
Else
    MOSAICO.value = vbUnchecked
    mAncho.Text = ""
    mLargo.Text = ""
End If
        
rs.Close

End Function


Private Sub Actmapas_Click()
Dim i As Integer
For i = 2 To 52
    Call SwitchMap("map" & i & ".map")
    Call SaveMapData("Map" & i & ".map")
Next
End Sub

Private Sub Blockedchk_Click()

Call PlaceBlockCmd_Click

End Sub

Private Sub ObtenerNombreArchivo(Guardar As Boolean)
With Dialog
    .Filter = "Mapas|*.map"
    
    If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .FileName = ""
            .flags = cdlOFNPathMustExist
            .ShowSave
           
    Else
        .DialogTitle = "Cargar"
        .FileName = ""
        
        .flags = cdlOFNFileMustExist
        .ShowOpen
    End If
End With

End Sub

Private Sub borratri_Click()
Dim Y As Integer
Dim X As Integer
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        MapData(X, Y).Trigger = 0
    Next X
Next Y

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

If Not MapaCargado Then
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

If Not MapaCargado Then
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

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lCelda_Click()
Grhtxt.Text = DameGrhIndex(lCelda.List(lCelda.ListIndex))
'If frmGrafico.Visible = False Then frmGrafico.Visible = True
Call PlaceGrhCmd_Click
End Sub

Private Sub lCelda_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
Grhtxt.SetFocus
End Sub

Private Sub lCelda_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Grhtxt.SetFocus
End Sub

Private Sub lCelda_KeyUp(KeyCode As Integer, Shift As Integer)
KeyCode = 0
Grhtxt.SetFocus
End Sub

Private Sub mnuBloquear_Click()

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
            MapData(X, Y).Blocked = 1
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

Private Sub mnuborrarArboles_Click()
Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

For Y = 1 To 100
    For X = 1 To 100
    
        If MapData(X, Y).NPCIndex > 0 Then
            
            Dim c As String
            c = GetVar(App.Path & "\npcs.dat", "NPC" & MapData(X, Y).NPCIndex, "NpcType")
            If c = "" Then c = 0
            If Val(c) = 3 Then
                If MapData(X, Y).NPCIndex = 25 Or MapData(X, Y).NPCIndex = 16 Then
                    InitGrh MapData(X, Y).ObjGrh, Val(GetVar(IniPath & "OBJ.dat", "OBJ147", "GrhIndex"))
                    MapData(X, Y).OBJInfo.objindex = 147
                    MapData(X, Y).Blocked = 1
                    MapData(X, Y).OBJInfo.Amount = 1
                ElseIf MapData(X, Y).NPCIndex = 26 Or MapData(X, Y).NPCIndex = 17 Then
                    InitGrh MapData(X, Y).ObjGrh, Val(GetVar(IniPath & "OBJ.dat", "OBJ148", "GrhIndex"))
                    MapData(X, Y).OBJInfo.objindex = 148
                    MapData(X, Y).Blocked = 1
                    MapData(X, Y).OBJInfo.Amount = 1
                ElseIf MapData(X, Y).NPCIndex = 27 Or MapData(X, Y).NPCIndex = 18 Then
                    InitGrh MapData(X, Y).ObjGrh, Val(GetVar(IniPath & "OBJ.dat", "OBJ149", "GrhIndex"))
                    MapData(X, Y).OBJInfo.objindex = 149
                    MapData(X, Y).Blocked = 1
                    MapData(X, Y).OBJInfo.Amount = 1
                ElseIf MapData(X, Y).NPCIndex = 28 Or MapData(X, Y).NPCIndex = 19 Then
                    InitGrh MapData(X, Y).ObjGrh, Val(GetVar(IniPath & "OBJ.dat", "OBJ150", "GrhIndex"))
                    MapData(X, Y).OBJInfo.objindex = 150
                    MapData(X, Y).Blocked = 1
                    MapData(X, Y).OBJInfo.Amount = 1
                ElseIf MapData(X, Y).NPCIndex = 29 Or MapData(X, Y).NPCIndex = 20 Then
                    InitGrh MapData(X, Y).ObjGrh, Val(GetVar(IniPath & "OBJ.dat", "OBJ151", "GrhIndex"))
                    MapData(X, Y).OBJInfo.objindex = 151
                    MapData(X, Y).Blocked = 1
                    MapData(X, Y).OBJInfo.Amount = 1
                ElseIf MapData(X, Y).NPCIndex = 30 Or MapData(X, Y).NPCIndex = 21 Then
                    InitGrh MapData(X, Y).ObjGrh, Val(GetVar(IniPath & "OBJ.dat", "OBJ152", "GrhIndex"))
                    MapData(X, Y).OBJInfo.objindex = 152
                    MapData(X, Y).Blocked = 1
                    MapData(X, Y).OBJInfo.Amount = 1
                ElseIf MapData(X, Y).NPCIndex = 31 Or MapData(X, Y).NPCIndex = 22 Then
                    InitGrh MapData(X, Y).ObjGrh, Val(GetVar(IniPath & "OBJ.dat", "OBJ153", "GrhIndex"))
                    MapData(X, Y).OBJInfo.objindex = 153
                    MapData(X, Y).Blocked = 1
                    MapData(X, Y).OBJInfo.Amount = 1
                ElseIf MapData(X, Y).NPCIndex = 32 Or MapData(X, Y).NPCIndex = 23 Then
                    InitGrh MapData(X, Y).ObjGrh, Val(GetVar(IniPath & "OBJ.dat", "OBJ154", "GrhIndex"))
                    MapData(X, Y).OBJInfo.objindex = 154
                    MapData(X, Y).Blocked = 1
                    MapData(X, Y).OBJInfo.Amount = 1
                End If
                MapData(X, Y).NPCIndex = 0
                Call EraseChar(MapData(X, Y).CharIndex)
            End If
'                If MapData(x, y).OBJInfo.objindex = 4 Or _
'                   MapData(x, y).OBJInfo.objindex = 5 Or _
'                   MapData(x, y).OBJInfo.objindex = 6 Then
'
'                    MapData(x, y).Blocked = 0
'                    MapData(x, y).OBJInfo.objindex = 0
'                    MapData(x, y).OBJInfo.Amount = 0
'                    MapData(x, y).ObjGrh.Grhindex = 0
'                    MapData(x, y).ObjGrh.FrameCounter = 0
'                    MapData(x, y).ObjGrh.SpeedCounter = 0
'                    MapData(x, y).ObjGrh.Started = 0
'
'
'
'                End If
        End If
        
        
    Next
Next

End Sub

Private Sub mnuCarac_Click()
frmCarac.Visible = True
End Sub

Private Sub mnuCargar_Click()
'frmCargar.Visible = True
Dialog.CancelError = True
On Error GoTo ErrHandler
Call ObtenerNombreArchivo(False)


If MapInfo.Changed = 1 Then
    If MsgBox("Este mapa há sido modificado. Vas a perder todos los cambios si no lo grabas. Lo queres grabar ahora?", vbYesNo) = vbYes Then
        Call SaveMapData(Dialog.FileName)
    End If
End If


    UserPos.X = (WindowTileWidth \ 2) + 1
    
    UserPos.Y = (WindowTileHeight \ 2) + 1
    
    Call mnuNuevo_Click
    Call SwitchMap(Dialog.FileName)
    EngineRun = True
Exit Sub

ErrHandler:
MsgBox Err.Description
End Sub

Private Sub mnuGrilla_Click()
frmGrilla.Visible = True
End Sub

Private Sub mnuMusica_Click()
frmMusica.Show
End Sub

Private Sub mnuNuevo_Click()


Dim Y As Integer
Dim X As Integer

Call borratri_Click

frmMain.MousePointer = 11
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        MapData(X, Y).Graphic(1).GrhIndex = 3
        'Change blockes status
        MapData(X, Y).Blocked = frmMain.Blockedchk.value

        'Erase layer 2 and 3
        MapData(X, Y).Graphic(2).GrhIndex = 0
        MapData(X, Y).Graphic(3).GrhIndex = 0
        MapData(X, Y).Graphic(4).GrhIndex = 0

        'Erase NPCs
        If MapData(X, Y).NPCIndex > 0 Then
            EraseChar MapData(X, Y).CharIndex
            MapData(X, Y).NPCIndex = 0
        End If

        'Erase Objs
        MapData(X, Y).OBJInfo.objindex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.GrhIndex = 0

        'Clear exits
        MapData(X, Y).TileExit.Map = 0
        MapData(X, Y).TileExit.X = 0
        MapData(X, Y).TileExit.Y = 0
         
         
        MapData(X, Y).Blocked = frmMain.Blockedchk.value
        MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = Val(frmMain.Grhtxt.Text)
            
        InitGrh MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)), Val(frmMain.Grhtxt.Text)
        
    Next X
Next Y


MapInfo.Changed = 1
MapInfo.MapVersion = 0

Text1.Text = "Nuevo Mapa"
UserPos.X = (WindowTileWidth \ 2) + 1

UserPos.Y = (WindowTileHeight \ 2) + 1

'CurMap = frmCargar.MapLst.ListCount
MapaCargado = True
EngineRun = True
frmMain.MousePointer = 0
End Sub

Private Sub MOSAICO_Click()
If mAncho.Text = "" Then mAncho.Text = "1"
If mLargo.Text = "" Then mLargo.Text = "1"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Apuntador.Move X, Y
UserPos.X = X
UserPos.Y = Y
Call ActualizaDespGrilla
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MiRadarX = X
MiRadarY = Y
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
CurrentGrh.GrhIndex = Val(Grhtxt.Text)
CurrentGrh.Started = 1
CurrentGrh.FrameCounter = 1
CurrentGrh.SpeedCounter = GrhData(CurrentGrh.GrhIndex).Speed

End Sub

Private Sub ClsBordMnu_Click()

'*****************************************************************
'Clears a border in a room with current GRH
'*****************************************************************

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
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
             MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = aux
            'Setup GRH
            InitGrh MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)), aux
          Else
            'Else Place graphic
            MapData(X, Y).Blocked = frmMain.Blockedchk.value
            MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = Val(frmMain.Grhtxt.Text)
            
            'Setup GRH
    
            InitGrh MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)), Val(frmMain.Grhtxt.Text)
        End If
             'Erase NPCs
            If MapData(X, Y).NPCIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NPCIndex = 0
            End If

            'Erase Objs
            MapData(X, Y).OBJInfo.objindex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.GrhIndex = 0

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

If Not MapaCargado Then
    Exit Sub
End If

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        MapData(X, Y).Graphic(1).GrhIndex = 3
        'Change blockes status
        MapData(X, Y).Blocked = frmMain.Blockedchk.value

        'Erase layer 2 and 3
        MapData(X, Y).Graphic(2).GrhIndex = 0
        MapData(X, Y).Graphic(3).GrhIndex = 0
        MapData(X, Y).Graphic(4).GrhIndex = 0

        'Erase NPCs
        If MapData(X, Y).NPCIndex > 0 Then
            EraseChar MapData(X, Y).CharIndex
            MapData(X, Y).NPCIndex = 0
        End If

        'Erase Objs
        MapData(X, Y).OBJInfo.objindex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.GrhIndex = 0

        'Clear exits
        MapData(X, Y).TileExit.Map = 0
        MapData(X, Y).TileExit.X = 0
        MapData(X, Y).TileExit.Y = 0

        If frmMain.MOSAICO.value = vbChecked Then
            Dim aux As Integer
            aux = Val(frmMain.Grhtxt.Text) + _
            ((Y Mod frmMain.mLargo) * frmMain.mAncho) + (X Mod frmMain.mAncho)
             MapData(X, Y).Blocked = frmMain.Blockedchk.value
             MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = aux
            'Setup GRH
            InitGrh MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)), aux
        Else
            'Else Place graphic
            MapData(X, Y).Blocked = frmMain.Blockedchk.value
            MapData(X, Y).Graphic(Val(frmMain.Layertxt.Text)).GrhIndex = Val(frmMain.Grhtxt.Text)
            
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

If Not MapaCargado Then Exit Sub

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
If Not MapaCargado Then Exit Sub

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
frmAbout1.Show
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



Call SaveMapData(Dialog.FileName)

'Set changed flag
MapInfo.Changed = 0

End Sub


Private Sub SaveNewMnu_Click()
Dialog.CancelError = True
On Error GoTo ErrHandler
Call ObtenerNombreArchivo(True)
Call SaveMapData(Dialog.FileName)
    'frmCargar.MapLst.AddItem "Map " & NumMaps, NumMaps - 1
Exit Sub

ErrHandler:
MsgBox Err.Description

End Sub


Private Sub Text1_Change()
MapInfo.Name = Text1.Text
End Sub

Private Sub WalkModeChk_Click()

'ToggleWalkMode

End Sub

Private Sub VScroll1_Scroll()

End Sub
