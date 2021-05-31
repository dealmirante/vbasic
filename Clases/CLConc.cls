VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLConc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private v_CodConc As Variant
Private v_CodProv As Variant
Private s_Descrip As String
Private s_fchDesde As String
Private s_fchHasta As String
Private s_Activo As String

Private s_Tipo As String
Private si_MontoProp As Single
Private si_MontoPaga As Single

Private i_sec As Integer
Private i_Mes As Long
Private s_CodRubr   As String
'Private i_Objetivo As Integer

Public col_producto As Collection
Public col_Plus As Collection

Public Function Delete_Col_Discriminado_PxC() As Boolean
Dim sql As String
Dim CodigoProveedor As CLProdPrec, setProveedor As String
Dim rs As Recordset
Dim salir As Boolean
    
salir = True

setProveedor = "( "

For Each CodigoProveedor In col_producto

    sql = " SELECT cod_prov FROM ESTADIS.Concurso_Prov " _
    & " WHERE id_concurso = '" & codConc & "' " _
    & " And cod_prov = '" & CodigoProveedor.CodProv & "' " _
    & " And SECUENCIA = '" & CodigoProveedor.sec & "' "
    setProveedor = setProveedor & "'" & CodigoProveedor.CodProv & "',"
    If Aplicacion.ObtenerRsDAO(sql, rs) Then
        If rs.RecordCount = 0 Then
        'No existe el proveedor -> insertar
            sql = "INSERT INTO  ESTADIS.Concurso_Prov ("
            sql = sql & "id_concurso,"
            sql = sql & "cod_prov, "
            sql = sql & "SECUENCIA "
            sql = sql & " ) "
            sql = sql & " VALUES ("
            sql = sql & "'" & codConc & "', "
            sql = sql & "'" & CodigoProveedor.CodProv & "', "
            sql = sql & CodigoProveedor.sec
            sql = sql & " ) "
            If Not Aplicacion.EjecutarDAO(sql) Then
                Exit For
                salir = False
            End If
        End If
    End If
Next
setProveedor = Left(setProveedor, Len(setProveedor) - 1) & ")"

'Borrar todos los proveedores que no pertenecen a la coleccion
If salir Then
    If setProveedor = "()" Then
        salir = Delete_COL_PxC
    Else
        sql = " DELETE FROM ESTADIS.Concurso_prov " _
        & " WHERE id_concurso = '" & codConc & "'" _
        & " And cod_prov NOT IN " & setProveedor
        salir = Aplicacion.EjecutarDAO(sql)
    End If
End If
Delete_Col_Discriminado_PxC = salir

End Function

Public Function Delete_COL_PxC() As Boolean
Dim sql As String

sql = "DELETE FROM ESTADIS.Concurso_Prov WHERE id_concurso = '" & codConc & "'"

Delete_COL_PxC = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Delete_COL_Prod() As Boolean
Dim sql As String

sql = "DELETE FROM ESTADIS.Concurso_d " _
& " WHERE id_concurso = '" & codConc & "'" _
& " And cod_prov = '" & CodProv & "' " _
& " And secuencia = " & sec

Delete_COL_Prod = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Delete_Productividad() As Boolean
Dim sql As String

sql = "DELETE FROM ESTADIS.codigos_productividad " _
& " WHERE aniomes = " & Mes & "" _

Delete_Productividad = Aplicacion.EjecutarDAO(sql)

End Function


Public Function Delete_Lanzamientos() As Boolean
Dim sql As String

sql = "DELETE FROM ESTADIS.codigos_lanzamientos " _
& " WHERE aniomes = " & Mes & "" _

Delete_Lanzamientos = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Delete_MonLinea() As Boolean
Dim sql As String

sql = "DELETE FROM ESTADIS.monitoreo_linea " _
& " WHERE aniomes = " & Mes & ""

Delete_MonLinea = Aplicacion.EjecutarDAO(sql)

End Function



Public Function Delete_COL_Plus() As Boolean
Dim sql As String

sql = "DELETE FROM ESTADIS.Concurso_Plus WHERE id_concurso = '" & codConc & "'" _
& " And cod_prov = '" & CodProv & "' And secuencia = " & sec

Delete_COL_Plus = Aplicacion.EjecutarDAO(sql)

End Function


Public Function Delete_PxC() As Boolean
Dim sql As String

sql = "DELETE FROM ESTADIS.Concurso_H WHERE id_concurso = '" & codConc & "'"

Delete_PxC = Aplicacion.EjecutarDAO(sql)

End Function

Public Function Insert_PxC() As Boolean
Dim sql As String
Dim cod As Variant

    sql = "INSERT INTO  ESTADIS.Concurso_h ("
    sql = sql & "id_concurso, "
    sql = sql & "fch_vdesde, "
    sql = sql & "fch_vhasta, "
    sql = sql & "descrip, "
    sql = sql & "activo, "
    sql = sql & "mes, "
    sql = sql & "cod_rubr "
    sql = sql & " ) "
    sql = sql & " VALUES ("
    sql = sql & "'" & codConc & "', "
    sql = sql & Func.func_ToDate(fchDesde) & ", "
    sql = sql & Func.func_ToDate(fchHasta) & ", "
    sql = sql & "'" & Descrip & "', "
    sql = sql & "'" & activo & "', "
    sql = sql & "" & Mes & ", "
    sql = sql & " '" & CodRubr & "' ) "
    
If Aplicacion.EjecutarDAO(sql) Then
    Insert_PxC = Insertar_Col_Prov
Else
    Insert_PxC = False
End If

End Function

Public Function Insert_Prod() As Boolean
Dim sql As String
Dim salir As Boolean
Dim clProd As CLProdPrec

salir = True
        sql = " Update estadis.concurso_prov Set Tipo = " & tipo
        sql = sql & " , monto_prop = " & MontoProp
        sql = sql & " , monto_paga = " & MontoPaga
        sql = sql & " Where id_concurso = '" & codConc & "' "
        sql = sql & " And cod_prov = '" & CodProv & "' "
        sql = sql & " and Secuencia = " & sec & " "
        
        Call Aplicacion.EjecutarDAO(sql)
        
    For Each clProd In col_producto
        sql = "INSERT INTO  ESTADIS.Concurso_d ("
        sql = sql & "id_concurso,"
        sql = sql & "cod_prov, "
        sql = sql & "cod_prod, "
        sql = sql & "paga, "
        sql = sql & "SECUENCIA "
        sql = sql & " ) "
        sql = sql & " VALUES ("
        sql = sql & "'" & codConc & "', "
        sql = sql & "'" & CodProv & "', "
        sql = sql & clProd.codProd & ", "
        sql = sql & clProd.Precio & ", "
        sql = sql & sec & " "
        sql = sql & " ) "
    
        If Not Aplicacion.EjecutarDAO(sql) Then
            Exit For
            salir = False
        End If
    Next
    If salir Then
        For Each clProd In col_Plus
            sql = "INSERT INTO  ESTADIS.Concurso_Plus ("
            sql = sql & "id_concurso,"
            sql = sql & "cod_prov, "
            sql = sql & "limite_rango, "
            sql = sql & "plus, "
            sql = sql & "secuencia, "
            sql = sql & "tipo_pluss, "
            sql = sql & "tipo_objetivo "
            sql = sql & " ) "
            sql = sql & " VALUES ("
            sql = sql & "'" & codConc & "', "
            sql = sql & "'" & CodProv & "', "
            sql = sql & clProd.codProd & ", "
            sql = sql & clProd.Precio & ", "
            sql = sql & sec & ", "
            sql = sql & clProd.tipo & ", "
            sql = sql & clProd.Objetivo
            sql = sql & " ) "
        
            If Not Aplicacion.EjecutarDAO(sql) Then
                Exit For
                salir = False
            End If
        Next
    End If
Insert_Prod = salir

End Function

Public Property Get codConc() As Variant
codConc = v_CodConc
End Property

Public Property Let codConc(ByVal vNewValue As Variant)
v_CodConc = vNewValue
End Property

Public Property Get CodProv() As Variant
CodProv = v_CodProv
End Property

Public Property Let CodProv(ByVal vNewValue As Variant)
v_CodProv = vNewValue
End Property

Public Property Get Descrip() As Variant
Descrip = s_Descrip
End Property

Public Property Let Descrip(ByVal vNewValue As Variant)
s_Descrip = vNewValue
End Property


Public Function Insertar_Col_Prov() As Boolean
Dim sql As String
Dim cod As Variant
Dim salir As Boolean
    
    
    salir = True
    For Each cod In col_producto
        sql = "INSERT INTO  ESTADIS.Concurso_Prov ("
        sql = sql & "id_concurso,"
        sql = sql & "cod_prov, "
        sql = sql & "secuencia "
        sql = sql & " ) "
        sql = sql & " VALUES ("
        sql = sql & "'" & codConc & "', "
        sql = sql & "'" & cod.CodProv & "', "
        sql = sql & "" & cod.sec & " "
        sql = sql & " ) "
    
        If Not Aplicacion.EjecutarDAO(sql) Then
            Exit For
            salir = False
        End If
    Next

Insertar_Col_Prov = salir

End Function

Public Function Insertar_Porductividad() As Boolean
Dim sql As String
Dim cod As Variant
Dim salir As Boolean
    
    
    salir = True
    For Each cod In col_producto
        sql = "INSERT INTO  ESTADIS.codigos_productividad ("
        sql = sql & "aniomes,"
        sql = sql & "fch_desde, "
        sql = sql & "fch_hasta, "
        sql = sql & "cod_prod, "
        sql = sql & "multiplo "
        sql = sql & " ) "
        sql = sql & " VALUES ("
        sql = sql & "" & Mes & ", "
        sql = sql & "" & func_ToDate(cod.fdesde) & ", "
        sql = sql & "" & func_ToDate(cod.fhasta) & ", "
        sql = sql & "" & cod.codProd & ", "
        sql = sql & "" & cod.Precio & " "
        sql = sql & " ) "
    
        If Not Aplicacion.EjecutarDAO(sql) Then
            Exit For
            salir = False
        End If
    Next

Insertar_Porductividad = salir

End Function


Public Function Insertar_Lanzamientos() As Boolean
Dim sql As String
Dim cod As Variant
Dim salir As Boolean
    
    
    salir = True
    For Each cod In col_producto
        sql = "INSERT INTO  ESTADIS.codigos_lanzamientos ("
        sql = sql & "aniomes,"
        sql = sql & "fch_desde, "
        sql = sql & "fch_hasta, "
        sql = sql & "cod_prod, objetivo "
        sql = sql & " ) "
        sql = sql & " VALUES ("
        sql = sql & "" & Mes & ", "
        sql = sql & "" & func_ToDate(cod.fdesde) & ", "
        sql = sql & "" & func_ToDate(cod.fhasta) & ", "
        sql = sql & "" & cod.codProd & ", "
        sql = sql & "" & cod.Objetivo & " "
        sql = sql & " ) "
    
        If Not Aplicacion.EjecutarDAO(sql) Then
            Exit For
            salir = False
        End If
    Next

Insertar_Lanzamientos = salir

End Function

Public Function Insertar_MonLinea() As Boolean
Dim sql As String
Dim cod As Variant
Dim salir As Boolean
    
    
    salir = True
    For Each cod In col_producto
        sql = "INSERT INTO  ESTADIS.Monitoreo_Linea ("
        sql = sql & "aniomes,"
        sql = sql & "cod_rubr, "
        sql = sql & "cod_prov, "
        sql = sql & "estimado_$ "
        sql = sql & " ) "
        sql = sql & " VALUES ("
        sql = sql & "" & Mes & ", "
        sql = sql & "'" & cod.CodRubr & "', "
        sql = sql & "'" & cod.CodProv & "', "
        sql = sql & "" & cod.Objetivo & " "
        sql = sql & " ) "
    
        If Not Aplicacion.EjecutarDAO(sql) Then
            Exit For
            salir = False
        End If
    Next

Insertar_MonLinea = salir

End Function


Public Function Update_PxC(ModGrilla As Boolean) As Boolean
Dim sql As String

sql = "UPDATE "
sql = sql & " ESTADIS.concurso_h SET "
sql = sql & " descrip  =" & "'" & Descrip & "',"
sql = sql & " fch_VDesde  =" & func_ToDate(fchDesde) & ","
sql = sql & " fch_VHasta  =" & func_ToDate(fchHasta) & ","
sql = sql & " Activo = '" & activo & "', "
sql = sql & " Mes = " & Mes & ", "
sql = sql & " Cod_Rubr = '" & CodRubr & "' "
sql = sql & " WHERE id_concurso = '" & codConc & "'"

If Aplicacion.EjecutarDAO(sql) Then
    If ModGrilla Then
        Update_PxC = Delete_Col_Discriminado_PxC
    Else
        Update_PxC = True
    End If
Else
    Update_PxC = False
End If
End Function
Public Function Update_Prod() As Boolean

Update_Prod = Delete_COL_Prod And Delete_COL_Plus And Insert_Prod

End Function
Public Function Update_Productividad() As Boolean

Update_Productividad = Delete_Productividad And Insertar_Porductividad

End Function

Public Function Update_Lanzamientos() As Boolean

Update_Lanzamientos = Delete_Lanzamientos And Insertar_Lanzamientos

End Function

Public Function Update_MonLinea() As Boolean

Update_MonLinea = Delete_MonLinea And Insertar_MonLinea

End Function


Private Sub Class_Initialize()
Set col_producto = New Collection
Set col_Plus = New Collection
End Sub
Public Property Get fchDesde() As Variant
fchDesde = s_fchDesde
End Property
Public Property Let fchDesde(ByVal vNewValue As Variant)
s_fchDesde = vNewValue
End Property
Public Property Get fchHasta() As Variant
fchHasta = s_fchHasta
End Property
Public Property Let fchHasta(ByVal vNewValue As Variant)
s_fchHasta = vNewValue
End Property
Public Property Get activo() As Variant
activo = s_Activo
End Property
Public Property Let activo(ByVal vNewValue As Variant)
s_Activo = vNewValue
End Property
Public Property Get sec() As Variant
sec = i_sec
End Property
Public Property Let sec(ByVal vNewValue As Variant)
i_sec = vNewValue
End Property
Public Property Get Mes() As Variant
Mes = i_Mes
End Property
Public Property Let Mes(ByVal vNewValue As Variant)
i_Mes = vNewValue
End Property
Public Property Get CodRubr() As Variant
CodRubr = s_CodRubr
End Property
Public Property Let CodRubr(ByVal vNewValue As Variant)
s_CodRubr = vNewValue
End Property
'Public Property Get Objetivo() As Variant
'Objetivo = i_Objetivo
'End Property
'Public Property Let Objetivo(ByVal vNewValue As Variant)
'i_Objetivo = vNewValue
'End Property

Public Property Get tipo() As Variant
tipo = s_Tipo
End Property

Public Property Let tipo(ByVal vNewValue As Variant)
s_Tipo = vNewValue
End Property

Public Property Get MontoProp() As Variant
MontoProp = si_MontoProp
End Property

Public Property Let MontoProp(ByVal vNewValue As Variant)
si_MontoProp = vNewValue
End Property

Public Property Get MontoPaga() As Variant
MontoPaga = si_MontoPaga
End Property

Public Property Let MontoPaga(ByVal vNewValue As Variant)
si_MontoPaga = vNewValue
End Property
