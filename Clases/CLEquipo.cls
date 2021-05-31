VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLEquipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private i_Cod As Long
Private s_Desc As String
Private s_Grupo As String
Private s_Turno As String
Private s_Tipo As String
Private s_GT As Integer
Private s_Rubros As String
Private s_CodDep As String
Private s_CodSdep As String

Private l_legajo As Long
Private l_Cajero As Long

Private s_TipoTime As String


Public Property Get cod() As Variant
cod = i_Cod
End Property

Public Property Let cod(ByVal vNewValue As Variant)
i_Cod = vNewValue
End Property

Public Function Delete_Equipos() As Boolean
Dim sql As String

sql = "DELETE FROM Estadis.Equipos WHERE Cod_equipo = " & cod

Delete_Equipos = Aplicacion.EjecutarDAO(sql)

End Function
Public Function Delete_Emp() As Boolean
Dim sql As String

sql = "DELETE FROM Estadis.Persona_equipos WHERE Legajo = " & cod

Delete_Emp = Aplicacion.EjecutarDAO(sql)

End Function

Public Property Get desc() As Variant
desc = s_Desc
End Property

Public Property Let desc(ByVal vNewValue As Variant)
s_Desc = vNewValue
End Property

Public Property Get Grupo() As Variant
Grupo = s_Grupo
End Property
Public Property Get Turno() As Variant
Turno = s_Turno
End Property
Public Property Get GT() As Variant
GT = s_GT
End Property

Public Property Let Grupo(ByVal vNewValue As Variant)
s_Grupo = vNewValue
End Property
Public Property Let Turno(ByVal vNewValue As Variant)
s_Turno = vNewValue
End Property
Public Property Let GT(ByVal vNewValue As Variant)
s_GT = vNewValue
End Property


Public Property Get Rubros() As Variant
Rubros = s_Rubros
End Property

Public Property Let Rubros(ByVal vNewValue As Variant)
s_Rubros = vNewValue
End Property

Public Property Get CodDep() As Variant
CodDep = s_CodDep
End Property

Public Property Let CodDep(ByVal vNewValue As Variant)
s_CodDep = vNewValue
End Property

Public Property Get CodSdep() As Variant
CodSdep = s_CodSdep
End Property

Public Property Let CodSdep(ByVal vNewValue As Variant)
s_CodSdep = vNewValue
End Property
Public Function Insert_Equipo() As Boolean
Dim sql As String

sql = "INSERT INTO  estadis.Equipos("
sql = sql & " Cod_equipo,"
sql = sql & " descrip,"
sql = sql & " cod_depn,"
sql = sql & " cod_sdep,"
sql = sql & " grupo,"
sql = sql & " rubros"
sql = sql & " ) "
sql = sql & " VALUES ("
sql = sql & "" & cod & ", "
sql = sql & "'" & desc & "', "
sql = sql & "'" & CodDep & "', "
sql = sql & "'" & CodSdep & "', "
sql = sql & "'" & Grupo & "', "
sql = sql & "'" & Rubros & "' "
sql = sql & " ) "

Insert_Equipo = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Insert_Emp() As Boolean
Dim sql As String

sql = "INSERT INTO  estadis.persona_equipos("
sql = sql & " legajo,"
sql = sql & " cod_depn,"
sql = sql & " cod_sdep,"
sql = sql & " grupo,"
sql = sql & " turno,"
sql = sql & " rubro,time,horas,tipo"
sql = sql & " ) "
sql = sql & " VALUES ("
sql = sql & "" & cod & ", "
sql = sql & "'" & CodDep & "', "
sql = sql & "'" & CodSdep & "', "
sql = sql & "'" & Grupo & "', "
sql = sql & "'" & Turno & "', "
sql = sql & "'" & Rubros & "', "
sql = sql & "'" & TipoTime & "', "
'If TipoTime = FULL Then
    sql = sql & "" & GT & ", "
'Else
'    sql = sql & "'" & Grupo & "' "
'End If
sql = sql & "'" & Tipo & "' "
sql = sql & " ) "

Insert_Emp = Aplicacion.EjecutarDAO(sql)

End Function


Public Function Update_Equipos() As Boolean
Dim sql As String

sql = "UPDATE estadis.Equipos"
sql = sql & " SET "
sql = sql & " descrip  =" & "'" & desc & "',"
sql = sql & " cod_depn  =" & "'" & CodDep & "',"
sql = sql & " cod_sdep  =" & "'" & CodSdep & "',"
sql = sql & " grupo  =" & "'" & Grupo & "',"
sql = sql & " rubros  =" & "'" & Rubros & "' "
sql = sql & " WHERE COD_equipo = " & cod & ""

Update_Equipos = Aplicacion.EjecutarDAO(sql)

End Function
Public Function Update_Emp() As Boolean
Dim sql As String

sql = "UPDATE estadis.persona_equipos"
sql = sql & " SET "
sql = sql & " cod_depn  =" & "'" & CodDep & "',"
sql = sql & " cod_sdep  =" & "'" & CodSdep & "',"
sql = sql & " grupo  =" & "'" & Grupo & "',"
sql = sql & " Turno  =" & "'" & Turno & "',"
sql = sql & " rubro  =" & "'" & Rubros & "',"
sql = sql & " Time  =" & "'" & TipoTime & "', "
sql = sql & " Tipo  =" & "'" & Tipo & "', "
'If TipoTime = FULL Then
'    sql = sql & " grupo_turno  =" & "'" & Grupo & "' "
'Else
    sql = sql & " horas  = " & GT & " "
'End If

sql = sql & " WHERE legajo = " & cod & ""

Update_Emp = Aplicacion.EjecutarDAO(sql)

End Function



Public Property Get Legajo() As Variant
Legajo = l_legajo
End Property

Public Property Let Legajo(ByVal vNewValue As Variant)
l_legajo = vNewValue
End Property

Public Function Insert_Persona() As Boolean
Dim sql As String

sql = "INSERT INTO  estadis.Persona_Equipos("
sql = sql & "Legajo,"
sql = sql & "COD_Equipo, "
sql = sql & "time "
sql = sql & " ) "
sql = sql & " VALUES ("
sql = sql & "" & Legajo & ", "
sql = sql & "" & cod & ", "
sql = sql & "'" & TipoTime & "' "
sql = sql & " ) "

Insert_Persona = Aplicacion.EjecutarDAO(sql)

End Function
Public Function Insert_Ausent(f As Date) As Boolean
Dim sql As String

sql = "INSERT INTO  personal.ausencia ("
sql = sql & "Legajo,"
sql = sql & "causa, "
sql = sql & "fecha "
sql = sql & " ) "
sql = sql & " VALUES ("
sql = sql & "" & Legajo & ", "
sql = sql & "'" & desc & "', "
sql = sql & "" & func_ToDate(Format$(f, FTOFECHA)) & " "
sql = sql & " ) "

Insert_Ausent = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Insert_Cajero() As Boolean
Dim sql As String

sql = "INSERT INTO  VENTAS.Cajeros("
sql = sql & "Cod_Cajero, "
sql = sql & "Legajo"
sql = sql & " ) "
sql = sql & " VALUES ("
sql = sql & "" & cajero & ", "
sql = sql & "" & Legajo & " "
sql = sql & " ) "

Insert_Cajero = Aplicacion.EjecutarDAO(sql)

End Function
Public Function Insert_User() As Boolean
Dim sql As String

sql = "INSERT INTO  estadis.perfiles ("
sql = sql & "usuario, "
sql = sql & "perfil,"
sql = sql & "nivel"
sql = sql & " ) "
sql = sql & " VALUES ("
sql = sql & "'" & desc & "' , "
sql = sql & "'" & Tipo & "', "
sql = sql & "" & cod & " "
sql = sql & " ) "

Insert_User = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Update_Persona() As Boolean
Dim sql As String

sql = "UPDATE estadis.Persona_Equipos "
sql = sql & " SET "
sql = sql & " cod_Equipo  = " & "" & cod & ", "
sql = sql & " time  = " & "'" & TipoTime & "' "
sql = sql & " WHERE legajo = " & Legajo & ""

Update_Persona = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Update_Cajero() As Boolean
Dim sql As String

sql = "UPDATE ventas.cajeros "
sql = sql & " SET "
sql = sql & " legajo  =" & "" & Legajo & ""
sql = sql & " WHERE cod_cajero = " & cajero & " "

Update_Cajero = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Update_User() As Boolean
Dim sql As String

sql = "UPDATE estadis.perfiles "
sql = sql & " SET "
sql = sql & " nivel  =" & "" & cod & ", "
sql = sql & " perfil  =" & "'" & Tipo & "'"
sql = sql & " WHERE usuario = '" & desc & "' "

Update_User = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Update_Ausen() As Boolean
Dim sql As String

sql = "UPDATE personal.ausencia "
sql = sql & " SET "
sql = sql & " causa  = " & "'" & desc & "' "
sql = sql & " WHERE fecha = " & func_ToDate(TipoTime) & " "
sql = sql & " And  Legajo = " & Legajo & " "

Update_Ausen = Aplicacion.EjecutarDAO(sql)

End Function
Public Function Delete_Ausen() As Boolean
Dim sql As String

sql = "Delete from personal.ausencia "
sql = sql & " WHERE fecha = " & func_ToDate(TipoTime) & " "
sql = sql & " And  Legajo = " & Legajo & " "

Delete_Ausen = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Delete_Persona() As Boolean
Dim sql As String

sql = "DELETE FROM estadis.Persona_Equipos "
sql = sql & " WHERE legajo = " & Legajo & ""

Delete_Persona = Aplicacion.EjecutarDAO(sql)

End Function


Public Function Delete_Cajero() As Boolean
Dim sql As String

sql = "DELETE FROM VENTAS.Cajeros "
sql = sql & " WHERE legajo = " & Legajo & " And cod_cajero = " & cajero

Delete_Cajero = Aplicacion.EjecutarDAO(sql)

End Function
Public Function Delete_User() As Boolean
Dim sql As String

sql = "DELETE FROM estadis.perfiles "
sql = sql & " WHERE usuario = '" & desc & "'  "

Delete_User = Aplicacion.EjecutarDAO(sql)

End Function


Public Property Get cajero() As Variant
cajero = l_Cajero
End Property

Public Property Let cajero(ByVal vNewValue As Variant)
l_Cajero = vNewValue
End Property

Public Property Get TipoTime() As Variant
TipoTime = s_TipoTime
End Property

Public Property Let TipoTime(ByVal vNewValue As Variant)
s_TipoTime = vNewValue
End Property

Public Property Get Tipo() As Variant
Tipo = s_Tipo
End Property

Public Property Let Tipo(ByVal vNewValue As Variant)
s_Tipo = vNewValue
End Property
