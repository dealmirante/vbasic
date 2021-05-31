Attribute VB_Name = "FuncCbos"
Option Explicit

Public Sub FuncCbo_LlenarCbosLST(ByRef Cbo As Control, ByRef lst As Control, Tabla As String)
Dim RS As Recordset
Dim sql As String

Select Case Tabla
    Case "DEPENDENCIA"
        sql = " SELECT cod_depn,descrip FROM dependencia ORDER BY cod_depn"
    Case "COMITENTE"
        sql = " SELECT cod_COMI,descrip FROM COMITENTE ORDER BY cod_COMI"
End Select

If Aplicacion.ObtenerRsDAO(sql, RS) Then
    
    Func_LlenarComboLst Cbo, lst, RS
    If Cbo.ListCount > 0 Then
        Cbo.ListIndex = -1
        lst.ListIndex = -1
    End If
    Aplicacion.CerrarDAO RS
    
End If

End Sub
Public Sub FuncCbos_LlenarCboLst(ByRef Cbo As Control, lst As Control, sql As String)
Dim RS As Recordset

If Aplicacion.ObtenerRsDAO(sql, RS) Then
   Func_LlenarComboLst Cbo, lst, RS
   Aplicacion.CerrarDAO RS
End If

End Sub

Public Sub FuncCbos_LlenarCboiTEM(ByRef Cbo As Control, sql As String)
Dim RS As Recordset

If Aplicacion.ObtenerRsDAO(sql, RS) Then
   Func_LlenarCombo Cbo, RS
   Aplicacion.CerrarDAO RS
End If

End Sub

Public Sub FuncCbos_LlenarCbo(ByRef Cbo As Control, sql As String)

Dim RS As Recordset

If Aplicacion.ObtenerRsDAO(sql, RS) Then
   Func_LlenarComboNoItem Cbo, RS
   Aplicacion.CerrarDAO RS
End If

End Sub


Public Sub FuncCbo_LlenarCbosCondic(ByRef Cbo As Control, ByRef lst As Control, sql As String)
Dim RS As Recordset

If Aplicacion.ObtenerRsDAO(sql, RS) Then
    
    Func_LlenarComboLst Cbo, lst, RS

End If

End Sub
