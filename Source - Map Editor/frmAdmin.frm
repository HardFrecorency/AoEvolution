VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admi Tool"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   1665
      TabIndex        =   67
      Top             =   3555
      Width           =   1140
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   90
      TabIndex        =   66
      Top             =   3555
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Opciones"
      Height          =   420
      Left            =   2385
      TabIndex        =   65
      Top             =   90
      Width           =   1185
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Tools"
      Height          =   420
      Left            =   3600
      TabIndex        =   64
      Top             =   90
      Width           =   1230
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   420
      Left            =   4860
      TabIndex        =   61
      Top             =   90
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   420
      Left            =   1215
      TabIndex        =   37
      Top             =   90
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   3690
      Top             =   1035
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   2670
      Left            =   45
      TabIndex        =   4
      Top             =   720
      Width           =   9420
      Begin VB.TextBox Text1 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   7
         Left            =   2115
         TabIndex        =   12
         Top             =   1530
         Width           =   915
      End
      Begin VB.TextBox AVelo 
         Height          =   330
         HideSelection   =   0   'False
         Left            =   7785
         TabIndex        =   29
         Top             =   2115
         Width           =   690
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   15
         Left            =   7020
         TabIndex        =   28
         Top             =   2115
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   14
         Left            =   6255
         TabIndex        =   27
         Top             =   2115
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   13
         Left            =   5490
         TabIndex        =   26
         Top             =   2115
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   12
         Left            =   4725
         TabIndex        =   25
         Top             =   2115
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   11
         Left            =   8550
         TabIndex        =   24
         Top             =   1395
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   10
         Left            =   7785
         TabIndex        =   23
         Top             =   1395
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   9
         Left            =   7020
         TabIndex        =   22
         Top             =   1395
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   8
         Left            =   6255
         TabIndex        =   21
         Top             =   1395
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   7
         Left            =   5490
         TabIndex        =   20
         Top             =   1395
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   6
         Left            =   4725
         TabIndex        =   19
         Top             =   1395
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   5
         Left            =   8550
         TabIndex        =   18
         Top             =   675
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   4
         Left            =   7785
         TabIndex        =   17
         Top             =   675
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   3
         Left            =   7020
         TabIndex        =   16
         Top             =   675
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   2
         Left            =   6255
         TabIndex        =   15
         Top             =   675
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   1
         Left            =   5490
         TabIndex        =   14
         Top             =   675
         Width           =   645
      End
      Begin VB.TextBox Frames 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   0
         Left            =   4725
         TabIndex        =   13
         Top             =   675
         Width           =   645
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "Insertar"
         Height          =   285
         Left            =   2205
         TabIndex        =   42
         Top             =   2325
         Width           =   960
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   285
         Left            =   1170
         TabIndex        =   41
         Top             =   2340
         Width           =   960
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Default         =   -1  'True
         Height          =   285
         Left            =   180
         TabIndex        =   40
         Top             =   2340
         Width           =   915
      End
      Begin VB.CommandButton Nav 
         Caption         =   "-"
         Height          =   285
         Index           =   1
         Left            =   3375
         TabIndex        =   39
         Top             =   2340
         Width           =   420
      End
      Begin VB.CommandButton Nav 
         Caption         =   "+"
         Height          =   285
         Index           =   0
         Left            =   3825
         TabIndex        =   38
         Top             =   2340
         Width           =   420
      End
      Begin VB.TextBox Text1 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   6
         Left            =   90
         TabIndex        =   5
         Top             =   405
         Width           =   3435
      End
      Begin VB.TextBox Text1 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   5
         Left            =   1080
         TabIndex        =   11
         Top             =   1530
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   4
         Left            =   90
         TabIndex        =   10
         Top             =   1530
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   3
         Left            =   2070
         TabIndex        =   8
         Top             =   945
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   2
         Left            =   1125
         TabIndex        =   7
         Top             =   945
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   1
         Left            =   3105
         TabIndex        =   9
         Top             =   945
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   330
         HideSelection   =   0   'False
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   945
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grh Logico"
         Height          =   195
         Index           =   7
         Left            =   2070
         TabIndex        =   60
         Top             =   1350
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Velocidad"
         Height          =   195
         Left            =   7830
         TabIndex        =   59
         Top             =   1845
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 16"
         Height          =   195
         Index           =   15
         Left            =   7020
         TabIndex        =   58
         Top             =   1845
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 15"
         Height          =   195
         Index           =   14
         Left            =   6255
         TabIndex        =   57
         Top             =   1845
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 14"
         Height          =   195
         Index           =   13
         Left            =   5490
         TabIndex        =   56
         Top             =   1845
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 13"
         Height          =   195
         Index           =   12
         Left            =   4725
         TabIndex        =   55
         Top             =   1845
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 12"
         Height          =   195
         Index           =   11
         Left            =   8550
         TabIndex        =   54
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 11"
         Height          =   195
         Index           =   10
         Left            =   7785
         TabIndex        =   53
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 10"
         Height          =   195
         Index           =   9
         Left            =   7020
         TabIndex        =   52
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 9"
         Height          =   195
         Index           =   8
         Left            =   6255
         TabIndex        =   51
         Top             =   1125
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 8"
         Height          =   195
         Index           =   7
         Left            =   5490
         TabIndex        =   50
         Top             =   1125
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 7"
         Height          =   195
         Index           =   6
         Left            =   4725
         TabIndex        =   49
         Top             =   1125
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 6"
         Height          =   195
         Index           =   5
         Left            =   8550
         TabIndex        =   48
         Top             =   405
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 5"
         Height          =   195
         Index           =   4
         Left            =   7785
         TabIndex        =   47
         Top             =   405
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 4"
         Height          =   195
         Index           =   3
         Left            =   7020
         TabIndex        =   46
         Top             =   405
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 3"
         Height          =   195
         Index           =   2
         Left            =   6255
         TabIndex        =   45
         Top             =   405
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 2"
         Height          =   195
         Index           =   1
         Left            =   5490
         TabIndex        =   44
         Top             =   405
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame 1"
         Height          =   195
         Index           =   0
         Left            =   4725
         TabIndex        =   43
         Top             =   405
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   36
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Largo"
         Height          =   195
         Index           =   5
         Left            =   1125
         TabIndex        =   35
         Top             =   1305
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ancho"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   34
         Top             =   1305
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start Y"
         Height          =   195
         Index           =   3
         Left            =   2070
         TabIndex        =   33
         Top             =   765
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start X"
         Height          =   195
         Index           =   2
         Left            =   1125
         TabIndex        =   32
         Top             =   765
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Frames"
         Height          =   195
         Index           =   1
         Left            =   3150
         TabIndex        =   31
         Top             =   765
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grh Fisico"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   30
         Top             =   765
         Width           =   705
      End
   End
   Begin VB.ListBox TXTList 
      Height          =   4155
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   4050
      Visible         =   0   'False
      Width           =   4650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar"
      Height          =   420
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   1140
   End
   Begin VB.PictureBox Visor 
      Height          =   4470
      Left            =   6435
      ScaleHeight     =   4410
      ScaleWidth      =   4320
      TabIndex        =   1
      Top             =   3465
      Width           =   4380
   End
   Begin VB.FileListBox GrhFiles 
      Height          =   4770
      Left            =   4860
      TabIndex        =   0
      Top             =   3465
      Width           =   1545
   End
   Begin VB.Label MiLastGrh 
      AutoSize        =   -1  'True
      Caption         =   "LastGrh:"
      Height          =   195
      Left            =   3015
      TabIndex        =   68
      Top             =   3615
      Width           =   600
   End
   Begin VB.Shape Shape2 
      Height          =   4785
      Left            =   45
      Top             =   3465
      Width           =   4785
   End
   Begin VB.Shape Shape1 
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   6270
   End
   Begin VB.Label YDIM 
      AutoSize        =   -1  'True
      Caption         =   "Height:"
      Height          =   195
      Left            =   9765
      TabIndex        =   63
      Top             =   8010
      Width           =   510
   End
   Begin VB.Label XDIM 
      AutoSize        =   -1  'True
      Caption         =   "Width:"
      Height          =   195
      Left            =   8100
      TabIndex        =   62
      Top             =   8010
      Width           =   465
   End
   Begin VB.Menu mnuLista 
      Caption         =   "Lista De Tiles"
      Visible         =   0   'False
      Begin VB.Menu Ordenar 
         Caption         =   "Ordenar ;-)"
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿??
'¿? ADMIN TOOL FOR TILES DATABASE OF ARGENTUM-ONLINE ¿?¿¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿??
'¿   Coded by mOrGoLoCk gulfas_morgolock@hotmail.com   ???
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿??

Option Explicit


'Bitmap header
Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Bitmap info header
Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Dim DIRECTORIO As String

Dim NumFrames As Integer

Dim TileApuntado As Integer

Private Sub ResetDatos()
Dim t
TXTList.Clear
For Each t In Text1
    t.Text = ""
Next
End Sub

Private Sub cmdEliminar_Click()
If TXTList.ListIndex > 0 Then
    TileApuntado = TXTList.ListIndex
    Dim Aux As String
    Aux = TXTList.List(TileApuntado)
    Aux = ReadField(2, Aux, 45)
    Aux = Right(Aux, Len(Aux) - 3)
    Call DelBD(Aux)
    TXTList.List(TXTList.ListIndex) = "ELIMINADO"
    
    TXTList.RemoveItem TXTList.ListIndex
    
    If TileApuntado < TXTList.ListCount And TileApuntado > 0 Then
            TXTList.ListIndex = TileApuntado
            
    End If
End If

End Sub



Private Sub cmdInsert_Click()
If Not TXTList.Visible Then Exit Sub

Dim cad As String

Dim j
For Each j In Text1
    j.Text = ""
Next

cad = "NUEVO" & "-" & "Grh" & BiggestVal & "-" & 1 & "-" & 0
cad = cad & "-" & 0 & "-" & 0 & "-" & 0 & "-" & 0
   
TXTList.AddItem cad
Call BuscarDescrip("NUEVO")

'Update la BD
If frmOptions.Check4 = vbChecked Then Call AddBD(ReadField(1, cad, 45), ReadField(2, cad, 45), 0, 0)

MiLastGrh.Caption = "LastGrh:" & BiggestVal - 1
End Sub

Function BuscarElemento(s As String) As Integer
Dim i As Integer
For i = 0 To TXTList.ListCount - 1
    If UCase(TXTList.List(i)) = UCase(s) Then
            BuscarElemento = i
            Exit For
    End If
Next i

End Function

Private Sub cmdUpdate_Click()
On Error Resume Next
NumFrames = Val(Text1(1).Text)
Dim cad As String
Dim Aux As String
If NumFrames = 1 Then
        If Not InfoOk Then
               MsgBox "Info invalida."
               TXTList.ListIndex = TileApuntado
               Exit Sub
        End If
        
        
        cad = UCase(Text1(6).Text) & "-" & _
        "Grh" & Text1(7).Text & "-" & Text1(1).Text & "-" & _
        Text1(0).Text & "-" & Text1(2).Text
        cad = cad & "-" & Text1(3).Text & "-" & Text1(4).Text & "-" & Text1(5).Text
                   
        TXTList.List(TileApuntado) = cad
        Changed = False
        TXTList.RemoveItem TileApuntado
        TXTList.AddItem cad
        TXTList.ListIndex = BuscarElemento(cad)
        
Else
        Dim j As Integer
        For j = 0 To 16 - 1
            If Frames(j).Text = "" Or j > NumFrames - 1 Then Exit For
            cad = cad & "-" & Frames(j).Text
        Next
        If Frames(j).Text = "" Then Text1(1).Text = j
        cad = UCase(Text1(6).Text) & "-" & "Grh" & Text1(7).Text & "-" & Text1(1).Text & cad
        cad = cad & "-" & AVelo.Text
        TXTList.List(TileApuntado) = cad
        Changed = False
        TXTList.RemoveItem TileApuntado
        TXTList.AddItem cad
        TXTList.ListIndex = BuscarElemento(cad)
        
End If
Call TXTList_Click
Aux = ReadField(2, cad, 45)
Aux = Right(Aux, Len(Aux) - 3)
Call UpdateBD(Aux, cad, 0, 0)
End Sub

Private Sub Command1_Click()
Dialog.CancelError = True
On Error GoTo ErrHandler
Call ObtenerNombreArchivo(False)
Call ResetDatos
frmProgreso.Progreso.Value = 0
TXTList.Visible = False

If frmOptions.Check5 = vbChecked Then LastGrh = Val(GetVar(App.Path & "\Grh.ini", "INIT", "NumGrhs"))

frmProgreso.Progreso.max = LastGrh
frmProgreso.Contador.Caption = "Cargando base de datos, por favor espere..."
frmProgreso.Show
frmProgreso.MousePointer = 11
Call CargarListBox(Dialog.FileName)
frmProgreso.MousePointer = 0
frmProgreso.Visible = False
TXTList.Visible = True
MiLastGrh.Caption = "LastGrh:" & BiggestVal - 1

ErrHandler:

End Sub

Private Sub Command2_Click()
Dialog.CancelError = True
On Error GoTo ErrHandler
Call ObtenerNombreArchivo(True)

frmProgreso.Progreso.Value = 0
TXTList.Visible = False

LastGrh = TXTList.ListCount

frmProgreso.Progreso.max = LastGrh
frmProgreso.Contador.Caption = "Guardando base de datos, por favor espere..."
frmProgreso.Show
frmProgreso.MousePointer = 11

Call SaveTXTListToFile(Dialog.FileName)

frmProgreso.MousePointer = 0
frmProgreso.Visible = False
TXTList.Visible = True

ErrHandler:

End Sub

Function FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

If Dir(File, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If

End Function

Private Sub SaveTXTListToFile(fn As String)
'This sub writes the info in the ListBox to an archive!


Dim sX As Integer
Dim sY As Integer
Dim pixelWidth As Integer
Dim pixelHeight As Integer
Dim FileNum As Integer

Dim Frames(1 To 16) As Integer
Dim Speed As Integer
Dim cad As String
Dim Descrip As String

Dim TempInt As Integer
Dim Grh As Integer
Dim Frame As Integer
Dim MisFrames As String
Dim StartX As String
Dim StartY As String
Dim Ancho As String
Dim Largo As String
Dim GrhFisico As String
Dim GrhLogico As String
Dim Velocidad As String
Dim ln As String
Dim n As Integer
Dim j As Integer
Dim f As File

On Error Resume Next
Set f = FSO.GetFile(fn)
Dim tau As String
tau = App.Path & "\[BackUP] " & f.Name
FSO.CopyFile fn, tau, True
If Err Then MsgBox Err.Description
If FileExist(fn, vbArchive) Then Kill (fn)

For Grh = 0 To LastGrh - 1
    DoEvents
    NumFrames = ReadField(3, TXTList.List(Grh), 45)
    ln = NumFrames
    If NumFrames = 1 Then
           'Descrip
            Descrip = ReadField(1, TXTList.List(Grh), 45)
            
            'GrhFisico
            GrhFisico = ReadField(4, TXTList.List(Grh), 45)
            
           
            'GrhLogico
            cad = ReadField(2, TXTList.List(Grh), 45)
            cad = Right(cad, Len(cad) - 3)
            GrhLogico = cad
        
            
            StartX = ReadField(5, TXTList.List(Grh), 45)
            StartY = ReadField(6, TXTList.List(Grh), 45)
            Ancho = ReadField(7, TXTList.List(Grh), 45)
            Largo = ReadField(8, TXTList.List(Grh), 45)
            ln = ln & "-" & GrhFisico
            ln = ln & "-" & StartX
            ln = ln & "-" & StartY
            ln = ln & "-" & Ancho
            ln = ln & "-" & Largo
            ln = ln & "-" & Descrip
            
    Else
            
           'Desc
            Descrip = ReadField(1, TXTList.List(Grh), 45)
            
            Dim t As String
            t = ReadField(2, TXTList.List(Grh), 45)
            t = Right(t, Len(t) - 3)
            GrhLogico = Val(t)
            
                            
            'Frames
            For j = 0 To NumFrames - 1
                  cad = cad & "-" & ReadField(4 + j, TXTList.List(Grh), 45)
            Next j
            
            Velocidad = ReadField(4 + NumFrames, TXTList.List(Grh), 45)
            ln = ln & cad & "-" & Velocidad & "-" & Descrip
            
   End If
   
   Call WriteVar(fn, "Graphics", "Grh" & GrhLogico, ln)
   frmProgreso.Progreso.Value = frmProgreso.Progreso.Value + 1
   ln = ""
   cad = ""
       
Next Grh



Exit Sub

errorhandler:
 MsgBox "Error mientras se escriben los datos al TXT"

End Sub

Private Sub Command3_Click()
frmAbout1.Show
End Sub

Private Sub Command4_Click()
frmTools.Visible = True
End Sub

Private Sub Command5_Click()
frmOptions.Visible = True
End Sub

Private Sub Command6_Click()
Call BuscarDescrip(Text2.Text)
End Sub

Private Sub Form_Load()
LastGrh = Val(GetVar(App.Path & "\Grh.ini", "INIT", "NumGrhs"))
DIRECTORIO = App.Path & "\Grh\"
Dialog.InitDir = App.Path & "\"
    
'GrhFiles.FileName
GrhFiles.Path = DIRECTORIO
Call AbrirBD

End Sub



Private Sub GetBitmapDimensions(BmpFile As String, bmWidth As Long, bmHeight As Long)
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open App.Path & "\GRH\" & BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Sub
Private Sub AbrirBD()
Err.Clear
On Error GoTo fin

With Conexion
    .Provider = "Microsoft.Jet.OLEDB.3.51"
    .ConnectionString = "Data Source=" & App.Path & "\grhindex\grhindex.mdb"
    .Open
End With

Exit Sub



fin:
If Err Then
    MsgBox "No se puede abrir la base de datos"
    End
End If

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim f As Form
For Each f In Forms
    Unload f
Next
End
End Sub



Private Sub Frames_GotFocus(Index As Integer)
Frames(Index).SelStart = 0
Frames(Index).SelLength = Len(Frames(Index).Text)
End Sub

Private Sub Frames_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    Else
        Changed = True
    End If
End If
End Sub

Private Sub Frames_LostFocus(Index As Integer)
Frames(Index).SelStart = 0
Frames(Index).SelLength = 0
End Sub

Private Sub GrhFiles_Click()

Dim Nombre As String
Dim X As Long, Y As Long

Nombre = GrhFiles.FileName
Call GetBitmapDimensions(Nombre, X, Y)
XDIM.Caption = "Width:" & X
YDIM.Caption = "Height:" & Y
If Nombre <> "" Then
    Visor.Picture = LoadPicture(DIRECTORIO & Nombre)
End If

If frmOptions.Check2 = vbChecked Then
        Dim s As String
        s = Left(Nombre, Len(Nombre) - 4)
        s = Right(s, Len(s) - 3)
        If Text1(0).Text <> s Then Changed = True
        Text1(0).Text = s
End If

End Sub

Private Sub CargarListBox(f As String)
On Error GoTo errorhandler

Dim sX As Integer
Dim sY As Integer
Dim pixelWidth As Integer
Dim pixelHeight As Integer
Dim FileNum As Integer

Dim Frames(1 To 16) As Integer
Dim Speed As Integer
Dim cad As String
Dim Descrip As String

Dim TempInt As Integer
Dim Grh As Integer
Dim Frame As Integer
Dim ln As String

For Grh = 1 To LastGrh
    DoEvents
    ln = GetVar(f, "Graphics", "Grh" & Grh)
    If ln <> "" Then
        NumFrames = Val(ReadField(1, ln, 45))
        If NumFrames <= 0 Then GoTo errorhandler
        '¿ES UNA ANIMACIóN?
        If NumFrames > 1 Then
            cad = UCase(ReadField(NumFrames + 3, ln, 45)) & "-" & NumFrames & "-"
            'Lee la animación
            For Frame = 1 To NumFrames
                Frames(Frame) = Val(ReadField(Frame + 1, ln, 45))
                If Frames(Frame) <= 0 Or Frames(Frame) > LastGrh Then
                        MsgBox "If Frames(Frame) <= 0 Or Frames(Frame) > LastGrh Then GoTo errorhandler"
                        GoTo errorhandler
                End If
                cad = cad & Frames(Frame) & "-"
            Next Frame
            Speed = Val(ReadField(NumFrames + 2, ln, 45))
            If Speed <= 0 Then GoTo errorhandler
            Descrip = ReadField(NumFrames + 3, ln, 45)
            If Descrip = "" Then Descrip = "Sin Desc"
            cad = Descrip & "-" & "Grh" & Grh & "-" & ln
            'cad = Left(cad, Len(cad) - Len(Descrip) - 1)
            TXTList.AddItem cad
        Else
            cad = UCase(Descrip) & "-" & FileNum & "-" & NumFrames & "-" & sX
            cad = cad & "-" & sY & "-" & pixelWidth & "-" & pixelHeight
            Descrip = ReadField(7, ln, 45)
            If Descrip = "" Then Descrip = "Sin Desc"
            cad = Descrip & "-" & "Grh" & Grh & "-" & ln
            If ReadField(7, ln, 45) = "" Then
                cad = Left(cad, Len(cad))
            Else
                cad = Left(cad, Len(cad) - Len(Descrip))
            End If
            TXTList.AddItem cad
            
        End If
  End If
  frmProgreso.Progreso.Value = frmProgreso.Progreso.Value + 1
Next Grh




Exit Sub

errorhandler:

End Sub

Private Sub Nav_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0
        If TXTList.ListIndex < TXTList.ListCount Then TXTList.ListIndex = TXTList.ListIndex + 1
    Case 1
        If TXTList.ListIndex > 0 Then TXTList.ListIndex = TXTList.ListIndex - 1
End Select
End Sub
Private Sub ClearTextBoxes()
Dim j, i As Integer
For Each j In Frames
        j.Text = ""
Next

For i = 2 To 5
        Text1(i).Text = ""
Next

AVelo.Text = ""

End Sub

Private Sub DisableFrames()
Dim j
For Each j In Frames
    j.Enabled = False
Next
End Sub
Private Sub EnableFrames()
Dim j
For Each j In Frames
    j.Enabled = True
Next
End Sub

Private Sub EnableOneTileInfo()
Dim j
For Each j In Text1
    j.Enabled = True
Next
End Sub
Private Sub DisableOneTileInfo()
Dim j
For Each j In Text1
    j.Enabled = False
Next
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Function InfoOk() As Boolean
InfoOk = Text1(0).Text <> "" And _
Text1(1).Text <> "" And _
Text1(2).Text <> "" And _
Text1(3).Text <> "" And _
Text1(4).Text <> "" And _
Text1(5).Text <> ""
End Function






Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (Index <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    Else
        Changed = True
    End If
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = 0
End Sub

Private Sub TXTList_Click()
Dim cad As String
Dim i As Integer
Dim j

If TXTList.List(TXTList.ListIndex) <> "ELIMINADO" Then
    If Changed Then
         If MsgBox("La info del actual tile fue cambiada, ¿Reflejar los cambios en el archivo?", vbYesNo) = vbYes Then
            If InfoOk Then
               Call cmdUpdate_Click
            Else
                MsgBox "Info invalida."
                TXTList.ListIndex = TileApuntado
                Exit Sub
            End If
        End If
    End If
End If

Changed = False
NumFrames = ReadField(3, TXTList.List(TXTList.ListIndex), 45)

Call ClearTextBoxes


If NumFrames = 1 Then

Call DisableFrames
Call EnableOneTileInfo

AVelo.Enabled = False

        If TXTList.List(TXTList.ListIndex) <> "ELIMINADO" Then '_
            'Descrip
            Text1(6).Text = ReadField(1, TXTList.List(TXTList.ListIndex), 45)
            
            'GrhFisico
              
            Text1(0).Text = ReadField(4, TXTList.List(TXTList.ListIndex), 45)
            
           
            'GrhLogico
            cad = ReadField(2, TXTList.List(TXTList.ListIndex), 45)
            cad = Right(cad, Len(cad) - 3)
            Text1(7).Text = cad
        
            cad = "GRH" & Text1(0).Text & ".bmp"
            i = 0
            Do While i < GrhFiles.ListCount
             If UCase(GrhFiles.List(i)) = UCase(cad) Then Exit Do
             i = i + 1
            Loop
            
            
            If UCase(GrhFiles.List(i)) = UCase(cad) Then GrhFiles.ListIndex = i
            
            
            Text1(1).Text = ReadField(3, TXTList.List(TXTList.ListIndex), 45)
            Text1(2).Text = ReadField(5, TXTList.List(TXTList.ListIndex), 45)
            Text1(3).Text = ReadField(6, TXTList.List(TXTList.ListIndex), 45)
            Text1(4).Text = ReadField(7, TXTList.List(TXTList.ListIndex), 45)
            Text1(5).Text = ReadField(8, TXTList.List(TXTList.ListIndex), 45)
        End If

Else
'*********************************************
    'Activamos casilleros apropiados
    EnableFrames
    For Each j In Frames
        j.Text = ""
    Next
    For i = 2 To 5
        Text1(i).Enabled = False
    Next
    Text1(0).Enabled = False
    AVelo.Enabled = True
'*********************************************
    
    If TXTList.List(TXTList.ListIndex) <> "ELIMINADO" Then '_
            'Desc
            Text1(6).Text = ReadField(1, TXTList.List(TXTList.ListIndex), 45)
            
            
            Text1(0).Text = "" 'ReadField(4, TXTList.List(TXTList.ListIndex), 45)
            Text1(7).Text = ReadField(2, TXTList.List(TXTList.ListIndex), 45)
            Text1(7).Text = Right(Text1(7).Text, Len(Text1(7).Text) - 3)
            i = 0
            Do While i < GrhFiles.ListCount
             If UCase(GrhFiles.List(i)) = UCase(cad) Then Exit Do
             i = i + 1
            Loop
            
            If UCase(GrhFiles.List(i)) = UCase(cad & ".bmp") Then GrhFiles.ListIndex = i
            
            
            Text1(1).Text = NumFrames
            
            
            For j = 0 To NumFrames - 1
                Frames(j).Text = ReadField(4 + j, TXTList.List(TXTList.ListIndex), 45)
            Next
            
            AVelo.Text = ReadField(4 + NumFrames, TXTList.List(TXTList.ListIndex), 45)
            
     End If

End If




TileApuntado = TXTList.ListIndex


If frmOptions.Check1 = vbChecked Then Text1(6).SetFocus

End Sub

Public Sub ObtenerNombreArchivo(Guardar As Boolean)


With Dialog
    .Filter = "Tiles indexes|*.txt"
    If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .FileName = ""
            .Flags = cdlOFNPathMustExist
            .ShowSave
    Else
        .DialogTitle = "Cargar"
        .FileName = ""
        
        .Flags = cdlOFNFileMustExist
        .ShowOpen
    End If
End With

End Sub

Private Sub TXTList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = vbRightButton Then PopupMenu mnuLista
End Sub
