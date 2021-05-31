VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLEstimado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_Id As String
Private s_Descrip As String
Private i_Anio As Integer
Private i_Mes As Integer
Private si_Imp As Double
Private l_Pax As Long
Private l_Tick As Long

Private s_Depn As String
Private s_Sdep As String
Private l_Sec As Long
Private s_Tipo As String

Public col_PorcRub As Collection
Public col_PorcDia As Collection
Public col_PorcComi As Collection
Public col_PorcEsp As Collection
Public col_PorcLoc As Collection

Public col_PrnPorcModG As Collection
Public col_PrnLPG As Collection
Public Function Delete_Espigon() As Boolean
Dim sql As String

sql = "DELETE FROM Estadis.porciento_espigon "
sql = sql & " WHERE anio = " & anio
sql = sql & " AND mes = " & mes

Delete_Espigon = Aplicacion.EjecutarDAO(sql)

End Function
Public Function Delete_Comi() As Boolean
Dim sql As String

sql = "DELETE FROM Estadis.porciento_comi "
sql = sql & " WHERE anio = " & anio
sql = sql & " AND mes = " & mes
sql = sql & " AND secuencia = " & Sec
sql = sql & " AND tipo_porc = '" & tipo & "'"

Delete_Comi = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Delete_Rubro() As Boolean
Dim sql As String

sql = "DELETE FROM Estadis.porciento_rubro "
sql = sql & " WHERE anio = " & anio
sql = sql & " AND mes = " & mes
sql = sql & " AND secuencia = " & Sec
sql = sql & " AND tipo_porc = '" & tipo & "'"

Delete_Rubro = Aplicacion.EjecutarDAO(sql)

End Function
Public Function Delete_Local() As Boolean
Dim sql As String

sql = "DELETE FROM Estadis.porciento_local "
sql = sql & " WHERE anio = " & anio
sql = sql & " AND mes = " & mes
sql = sql & " AND secuencia = " & Sec
sql = sql & " AND tipo_porc = '" & tipo & "'"

Delete_Local = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Delete_Porciento() As Boolean
Dim sql As String

sql = "DELETE FROM Estadis.porciento_cabezera "
sql = sql & " WHERE anio = " & anio
sql = sql & " AND mes = " & mes
sql = sql & " AND secuencia = " & Sec
sql = sql & " AND tipo_porc = '" & tipo & "'"
sql = sql & " AND cod_depn = '" & Depn & "'"
sql = sql & " AND cod_sdep = '" & Sdep & "'"

Delete_Porciento = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Delete_Dia() As Boolean
Dim sql As String

sql = "DELETE FROM Estadis.porciento_diario "
sql = sql & " WHERE anio = " & anio
sql = sql & " AND mes = " & mes
sql = sql & " AND secuencia = " & Sec
sql = sql & " AND tipo_porc = '" & tipo & "'"

Delete_Dia = Aplicacion.EjecutarDAO(sql)

End Function


Public Function Insert_ColEspigon() As Boolean
Dim sql As String
Dim cl_esp As CLGeneric
Dim salir As Boolean

salir = True
    
    For Each cl_esp In col_PorcEsp
        sql = "INSERT INTO estadis.porciento_espigon ("
        sql = sql & "anio,"
        sql = sql & "mes,"
        sql = sql & "cod_depn,"
        sql = sql & "cod_sdep,"
        sql = sql & "tipo_porc,"
        sql = sql & "porcentaje)"
        sql = sql & "values ("
        sql = sql & "" & anio & ", "
        sql = sql & "" & mes & ", "
        sql = sql & "'" & cl_esp.Depn & "',"
        sql = sql & "'" & cl_esp.Identif & "',"
        sql = sql & "'" & cl_esp.tipo & "',"
        sql = sql & "" & cl_esp.Porc & " )"
        
        If Not Aplicacion.EjecutarDAO(sql) Then
            salir = False
            Exit For
        End If
    Next

Insert_ColEspigon = salir

End Function

Public Function Insert_ColComi() As Boolean
Dim sql As String
Dim cl_esp As CLGeneric
Dim salir As Boolean

salir = True

    For Each cl_esp In col_PorcComi
        sql = "INSERT INTO estadis.porciento_comi ("
        sql = sql & "anio,"
        sql = sql & "mes,"
        sql = sql & "tipo_porc,"
        sql = sql & "secuencia,"
        sql = sql & "comi,"
        sql = sql & "porcentaje)"
        sql = sql & "values ("
        sql = sql & "" & anio & ", "
        sql = sql & "" & mes & ", "
        sql = sql & "'" & tipo & "',"
        sql = sql & "" & Sec & ","
        sql = sql & "'" & cl_esp.Identif & "',"
        sql = sql & "" & cl_esp.Porc & " )"
        
        If Not Aplicacion.EjecutarDAO(sql) Then
            salir = False
            Exit For
        End If
    Next
    Insert_ColComi = salir
End Function

Public Function Insert_ColRubro() As Boolean
Dim sql As String
Dim cl_esp As CLGeneric
Dim salir As Boolean

salir = True

    For Each cl_esp In col_PorcRub
        sql = "INSERT INTO estadis.porciento_rubro ("
        sql = sql & "anio,"
        sql = sql & "mes,"
        sql = sql & "tipo_porc,"
        sql = sql & "secuencia,"
        sql = sql & "cod_rubro,"
        sql = sql & "comi,"
        sql = sql & "porcentaje)"
        sql = sql & "values ("
        sql = sql & "" & anio & ", "
        sql = sql & "" & mes & ", "
        sql = sql & "'" & tipo & "',"
        sql = sql & "" & Sec & ","
        sql = sql & "'" & cl_esp.Identif & "',"
        sql = sql & "'" & cl_esp.tipo & "',"
        sql = sql & "" & cl_esp.Porc & " )"
        
        If Not Aplicacion.EjecutarDAO(sql) Then
            salir = False
            Exit For
        End If
    Next
    Insert_ColRubro = salir
End Function

Public Function Insert_ColLocal() As Boolean
Dim sql As String
Dim cl_esp As CLGeneric
Dim salir As Boolean

salir = True

    For Each cl_esp In col_PorcLoc
        sql = "INSERT INTO estadis.porciento_local ("
        sql = sql & "anio,"
        sql = sql & "mes,"
        sql = sql & "tipo_porc,"
        sql = sql & "secuencia,"
        sql = sql & "cod_rubro,"
        sql = sql & "cod_local,"
        sql = sql & "porcentaje)"
        sql = sql & "values ("
        sql = sql & "" & anio & ", "
        sql = sql & "" & mes & ", "
        sql = sql & "'" & tipo & "',"
        sql = sql & "" & Sec & ","
        sql = sql & "'" & cl_esp.Identif & "',"
        sql = sql & "'" & cl_esp.Locale & "',"
        sql = sql & "" & cl_esp.Porc & " )"
        
        If Not Aplicacion.EjecutarDAO(sql) Then
            salir = False
            Exit For
        End If
    Next
    Insert_ColLocal = salir
End Function


Public Function Insert_ColDia() As Boolean
Dim sql As String
Dim cl_esp As CLGeneric
Dim salir As Boolean

salir = True

    For Each cl_esp In col_PorcDia
        sql = "INSERT INTO estadis.porciento_diario ("
        sql = sql & "anio,"
        sql = sql & "mes,"
        sql = sql & "tipo_porc,"
        sql = sql & "secuencia,"
        sql = sql & "fch_dia,"
        sql = sql & "grupo,"
        sql = sql & "porcentaje)"
        sql = sql & "values ("
        sql = sql & "" & anio & ", "
        sql = sql & "" & mes & ", "
        sql = sql & "'" & tipo & "',"
        sql = sql & "" & Sec & ","
        sql = sql & "" & func_ToDate(cl_esp.Identif) & ","
        sql = sql & "'" & cl_esp.Grupo & "',"
        sql = sql & "" & cl_esp.Porc & " )"
        
        If Not Aplicacion.EjecutarDAO(sql) Then
            salir = False
            Exit For
        End If
    Next
    Insert_ColDia = salir
End Function



Public Function Insert_PrintModelo()
Dim sql As String
Dim cl_esp As CLGeneric
Dim salir As Boolean

salir = True

Call Aplicacion.EjecutarDAO("Delete from estadis.Prn_Modelo")

For Each cl_esp In col_PrnPorcModG

    sql = "INSERT INTO  ESTADIS.Prn_Modelo("
    sql = sql & "Anio,"
    sql = sql & "Mes,"
    sql = sql & "Tipo_porc,"
    sql = sql & "Cod_depn,"
    sql = sql & "Cod_sdep,"
    sql = sql & "Descrip,"
    sql = sql & "Fch_dia,"
    sql = sql & "Porc_dia,"
    sql = sql & "Porc_ga,"
    sql = sql & "Porc_gb,"
    sql = sql & "Porc_gc,"
    sql = sql & "Sec"
    sql = sql & " ) "
    sql = sql & " VALUES ("
    sql = sql & "" & anio & ", "
    sql = sql & "" & mes & ", "
    sql = sql & "'" & tipo & "', "
    sql = sql & "'" & Depn & "', "
    sql = sql & "'" & Sdep & "', "
    sql = sql & "'" & Descrip & "', "
    sql = sql & "" & func_ToDate(cl_esp.Identif) & ","
    sql = sql & "" & cl_esp.Porc & ", "
    sql = sql & "" & cl_esp.PorcGA & ", "
    sql = sql & "" & cl_esp.PorcGB & ", "
    sql = sql & "" & cl_esp.PorcGC & ", "
    sql = sql & "" & Sec & " "
    sql = sql & " ) "

    If Not Aplicacion.EjecutarDAO(sql) Then
        salir = False
        Exit For
    End If

Next
Insert_PrintModelo = salir
End Function
Public Function Insert_PrintRubros()
Dim sql As String
Dim cl_esp As CLGeneric
Dim salir As Boolean

salir = True

Call Aplicacion.EjecutarDAO("Delete from estadis.Prn_rubros")

For Each cl_esp In col_PrnPorcModG

    sql = "INSERT INTO  estadis.Prn_Rubros("
    sql = sql & "Anio,"
    sql = sql & "Mes,"
    sql = sql & "Tipo_porc,"
    sql = sql & "Cod_depn,"
    sql = sql & "Cod_sdep,"
    sql = sql & "Descrip,"
    sql = sql & "cod,"
    sql = sql & "porc_1,"
    sql = sql & "porc_2,"
    sql = sql & "porc_3,"
    sql = sql & "porc_4"
    sql = sql & " ) "
    sql = sql & " VALUES ("
    sql = sql & "" & anio & ", "
    sql = sql & "" & mes & ", "
    sql = sql & "'" & cl_esp.tipo & "', "
    sql = sql & "'" & Depn & "', "
    sql = sql & "'" & Sdep & "', "
    sql = sql & "'" & Descrip & "', "
    sql = sql & "'" & cl_esp.Identif & "',"
    sql = sql & "" & cl_esp.Porc & ", "
    sql = sql & "" & cl_esp.PorcGA & ", "
    sql = sql & "" & cl_esp.PorcGB & ", "
    sql = sql & "" & cl_esp.PorcGC & " "
    sql = sql & " ) "

    If Not Aplicacion.EjecutarDAO(sql) Then
        salir = False
        Exit For
    End If

Next
Insert_PrintRubros = salir
End Function

Public Property Get tipo() As Variant
tipo = s_Tipo
End Property

Public Property Let tipo(ByVal vNewValue As Variant)
s_Tipo = vNewValue
End Property

Public Property Get Sec() As Variant
Sec = l_Sec
End Property

Public Property Let Sec(ByVal vNewValue As Variant)
l_Sec = vNewValue
End Property

Public Function Delete_Estim(anio As Integer, mes As Integer) As Boolean
Dim sql As String

sql = "DELETE FROM Estadis.Modelo_Estim "
sql = sql & " WHERE anio = " & anio
sql = sql & " AND mes = " & mes

Delete_Estim = Aplicacion.EjecutarDAO(sql)

End Function

Public Property Get Id() As Variant
Id = s_Id
End Property
Public Function Insert_Estim() As Boolean
Dim sql As String, salir As Boolean
Dim cl_esp As CLGeneric

salir = True


sql = "INSERT INTO  estadis.modelo_estim("
sql = sql & "anio,"
sql = sql & "mes,"
sql = sql & "descrip,"
sql = sql & "importe,"
sql = sql & "ticket,"
sql = sql & "pax"
sql = sql & " ) "
sql = sql & " VALUES ("
sql = sql & "" & anio & ", "
sql = sql & "" & mes & ", "
sql = sql & "'" & Descrip & "', "
sql = sql & "" & Importe & ", "
sql = sql & "" & Ticket & ", "
sql = sql & "" & Pax & " "
sql = sql & " ) "

If Aplicacion.EjecutarDAO(sql) Then
    salir = Insert_ColEspigon
Else
    salir = False
End If

Insert_Estim = salir

End Function

Public Function Update_Estim() As Boolean
Dim sql As String
Dim salida As Boolean

salida = True

sql = "UPDATE "
sql = sql & " estadis.modelo_estim SET "
sql = sql & " importe =" & "" & Importe & ","
sql = sql & " pax  =" & "" & Pax & ","
sql = sql & " ticket  =" & "" & Ticket & ""
sql = sql & " WHERE anio = " & anio
sql = sql & " and mes = " & mes

If Aplicacion.EjecutarDAO(sql) Then
    salida = Delete_Espigon
    If salida Then
        salida = Insert_ColEspigon
    End If
Else
    salida = False
End If

Update_Estim = salida

End Function

Public Function Update_Porciento() As Boolean
Dim sql As String
Dim salida As Boolean

salida = (Delete_Comi And Delete_Rubro And Delete_Dia And Delete_Local)
If salida Then
    salida = (Insert_ColComi And Insert_ColRubro And Insert_ColDia And Insert_ColLocal)
End If

Update_Porciento = salida

End Function

Public Function Insert_Porciento() As Boolean
Dim sql As String, salir As Boolean
Dim rs As Recordset
Dim seguir As Boolean

sql = "SELECT descrip FROM estadis.porciento_cabezera "
sql = sql & " WHERE anio =" & anio
sql = sql & " AND mes =" & mes
sql = sql & " AND cod_depn = '" & Depn & "'"
sql = sql & " AND cod_sdep = '" & Sdep & "'"
sql = sql & " AND tipo_porc = '" & tipo & "'"
sql = sql & " AND nivel = 0 "

If Aplicacion.Nivel = 0 Then
    If Aplicacion.ObtenerRsDAO(sql, rs) Then
        If rs.RecordCount = 0 Then
            seguir = True
        Else
            MsgBox "Ya existe un modelo inicical", vbOKOnly + vbExclamation, "ATENCION"
            seguir = False
        End If
    End If
Else
    seguir = True
End If

If seguir Then
    salir = True
    
    sql = "INSERT INTO  estadis.porciento_cabezera("
    sql = sql & "anio,"
    sql = sql & "mes,"
    sql = sql & "tipo_porc,"
    sql = sql & "secuencia,"
    sql = sql & "cod_depn,"
    sql = sql & "cod_sdep,"
    sql = sql & "descrip,"
    sql = sql & "NIVEL"
    sql = sql & " ) "
    sql = sql & " VALUES ("
    sql = sql & "" & anio & ", "
    sql = sql & "" & mes & ", "
    sql = sql & "'" & tipo & "', "
    sql = sql & "" & Sec & ", "
    sql = sql & "'" & Depn & "', "
    sql = sql & "'" & Sdep & "', "
    sql = sql & "'" & Descrip & "', "
    sql = sql & Aplicacion.Nivel
    sql = sql & " ) "
    
    If Aplicacion.EjecutarDAO(sql) Then
        If Insert_ColComi Then
            If Insert_ColRubro Then
                If Insert_ColLocal Then
                    salir = Insert_ColDia
                Else
                    salir = False
                End If
            Else
                salir = False
            End If
        Else
            salir = False
        End If
    Else
        salir = False
    End If
    
    Insert_Porciento = salir
    
End If
End Function

Public Property Let Id(ByVal vNewValue As Variant)
s_Id = vNewValue
End Property


Public Property Get anio() As Variant
anio = i_Anio
End Property

Public Property Let anio(ByVal vNewValue As Variant)
i_Anio = vNewValue
End Property

Public Property Get mes() As Variant
mes = i_Mes
End Property

Public Property Let mes(ByVal vNewValue As Variant)
i_Mes = vNewValue
End Property

Public Property Get Descrip() As Variant
Descrip = s_Descrip
End Property

Public Property Let Descrip(ByVal vNewValue As Variant)
s_Descrip = vNewValue
End Property

Public Property Get Importe() As Variant
Importe = si_Imp
End Property

Public Property Let Importe(ByVal vNewValue As Variant)
si_Imp = vNewValue
End Property


Private Sub Class_Initialize()
    Set col_PorcRub = New Collection
    Set col_PorcDia = New Collection
    Set col_PorcComi = New Collection
    Set col_PorcEsp = New Collection
End Sub



Public Property Get Ticket() As Variant
Ticket = l_Tick
End Property

Public Property Let Ticket(ByVal vNewValue As Variant)
l_Tick = vNewValue
End Property

Public Property Get Pax() As Variant
Pax = l_Pax
End Property

Public Property Let Pax(ByVal vNewValue As Variant)
l_Pax = vNewValue
End Property

Public Property Get Depn() As Variant
Depn = s_Depn
End Property

Public Property Let Depn(ByVal vNewValue As Variant)
s_Depn = vNewValue
End Property

Public Property Get Sdep() As Variant
Sdep = s_Sdep
End Property

Public Property Let Sdep(ByVal vNewValue As Variant)
s_Sdep = vNewValue
End Property


