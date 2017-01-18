Attribute VB_Name = "Declares"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

'Tipos de objetos


Public Type ObjData
    
    Name As String 'Nombre del obj
    
    ObjType As Integer 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    MinHP As Integer ' Minimo puntos de vida
    MaxHP As Integer ' Maximo puntos de vida
    
    info As String
    
    MinHIT As Integer 'Minimo golpe
    MaxHIT As Integer 'Maximo golpe
    
    DEF As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    
    Vendible As Integer   ' �Se puede vender o comprar?
    Valor As Long      ' Precio
    
    Cerrada As Integer
    Llave As Byte
    Clave As Integer 'si clave=llave la puerta se abre o cierra
    Resistencia As Long
    
    Texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida1 As String
    ClaseProhibida2 As String
    ClaseProhibida3 As String
    ClaseProhibida4 As String
    ClaseProhibida5 As String
    ClaseProhibida6 As String
    ClaseProhibida7 As String
End Type

Public Type Obj
    objindex As Integer
    Amount As Integer
End Type


Public Conexion As New Connection
Public NumMidi As Integer
Public prgRun As Boolean
Public CurrentGrh As Grh
Public ENDL As String
Public Play As Boolean
Public MapaCargado As Boolean
Public MiRadarX As Integer
Public MiRadarY As Integer
Public ObjData() As ObjData


Sub LoadOBJData2()

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim numobjdatas

'obtiene el numero de obj
numobjdatas = Val(GetVar(IniPath & "Obj.dat", "INIT", "NumObjs"))
ReDim ObjData(1 To numobjdatas) As ObjData
  
'Llena la lista
For Object = 1 To numobjdatas
    
    ObjData(Object).Name = GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Name")
    
    ObjData(Object).GrhIndex = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "GrhIndex"))
    
    ObjData(Object).ObjType = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "ObjType"))
    
    ObjData(Object).Ropaje = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "NumRopaje"))
    
    ObjData(Object).info = GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Info")
    
    ObjData(Object).WeaponAnim = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Anim"))
    
    
    ObjData(Object).MaxHIT = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
    ObjData(Object).MinHIT = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
    ObjData(Object).MaxHP = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "MinHP"))
 
    ObjData(Object).DEF = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "DEF"))
    
    ObjData(Object).Vendible = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Vendible"))
    ObjData(Object).Valor = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Valor"))
    
    ObjData(Object).Cerrada = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Llave"))
    End If
    
    
    ObjData(Object).Texto = GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "GrhSec"))
    ObjData(Object).Clave = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Clave"))
    
            
    ObjData(Object).Resistencia = Val(GetVar(IniPath & "Obj.dat", "OBJ" & Object, "Resistencia"))
    
    

Next Object

End Sub

