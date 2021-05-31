Attribute VB_Name = "Func"
Option Explicit

Public Sub Func_FormatRange(ColDesde As Integer, ColHasta As Integer, RowDesde As Integer, RowHasta As Integer, spr As control, fto As String, FS As Boolean)

    'Select a block of cells
    spr.col = ColDesde
    spr.Row = RowDesde
    spr.Col2 = ColHasta
    spr.Row2 = RowHasta
    spr.BlockMode = True

    'Define cells as type FLOAT
    spr.CellType = SS_CELL_TYPE_FLOAT
    spr.TypeHAlign = SS_CELL_H_ALIGN_RIGHT
    spr.TypeFloatMin = "-" & fto
    spr.TypeFloatMax = fto

    spr.TypeFloatDecimalPlaces = 0
    spr.TypeFloatCurrencyChar = Asc("%")
    spr.TypeFloatSepChar = Asc(",")
    spr.TypeFloatDecimalChar = Asc(".")
    spr.TypeFloatSeparator = FS
    'Turn block mode off
    spr.BlockMode = False


End Sub


Public Function Func_ObtenerDesc(ByVal SQL As String, ByRef desc As String) As Boolean
Dim rs As Recordset

If Aplicacion.ObtenerRsDAO(SQL$, rs) Then
    If Aplicacion.CantReg(rs) > 0 Then
        Func_ObtenerDesc = True
        desc = rs!Descrip
    Else
        Func_ObtenerDesc = False
        desc = ""
    End If
    Aplicacion.CerrarDAO rs
End If

End Function


Public Function Func_CantidadDias(fch As String) As Integer
Dim fchDesde As String
Dim fchHasta As String

fchDesde = func_Dia1SegunMes_Anio(Month(Format$(fch, FTOFECHA)), Year(Format$(fch, FTOFECHA)))
fchHasta = func_Dia30SegunMes_Anio(Month(Format$(fch, FTOFECHA)), Year(Format$(fch, FTOFECHA)))

Func_CantidadDias = CDate(fchHasta) - CDate(fchDesde) + 1

End Function

Public Function func_Dia1SegunMes_Anio(Mes As Integer, anio As Integer)
func_Dia1SegunMes_Anio = "01-" & Format(Mes, "0#") & "-" & Trim(str(anio))
End Function

Public Function func_Dia30SegunMes_Anio(Mes As Integer, anio As Integer)
Dim MesAux, AnioAux
Dim fecha

If Mes = 12 Then
    MesAux = 1
    AnioAux = anio + 1
Else
    MesAux = Mes + 1
    AnioAux = anio
End If

fecha = "01-" & Format(MesAux, "0#") & "-" & Trim(str(AnioAux))

func_Dia30SegunMes_Anio = Format$(CDate(fecha) - 1, FTOFECHA)

End Function
Public Function func_ToDate(fch As String) As String

func_ToDate = "To_date ('" & fch & "','" & FTOFECHA & "')"

End Function

Public Function Update_() As Boolean
Dim SQL As String

SQL = "UPDATE "
SQL = SQL & " XXXX SET "
SQL = SQL & " xxx  =" & "'" & "" & "',"
SQL = SQL & " xxx  =" & "'" & "" & "',"
SQL = SQL & " xxx  =" & "'" & "" & "',"
SQL = SQL & " xxx  =" & "'" & "" & "',"
SQL = SQL & " xxx  =" & "'" & "" & "',"
SQL = SQL & " WHERE CODIGO = '" & "" & "'"

Update_ = Aplicacion.EjecutarDAO(SQL)

End Function

Public Sub Func_LlenarCombo(Cbo As control, rs As Recordset)
Dim i As Integer
Dim cant As Integer

cant% = Aplicacion.CantReg(rs)
Cbo.Clear

If cant > 0 Then
    For i% = 0 To cant - 1
        Cbo.AddItem rs.Fields(1).Value
        Cbo.ItemData(Cbo.NewIndex) = rs.Fields(0).Value
        rs.MoveNext
    Next
End If

End Sub


Public Sub Func_LlenarComboLst(ByRef Cbo As control, ByRef lst As control, rs As Recordset)
Dim i As Integer
Dim cant As Integer
cant% = Aplicacion.CantReg(rs)
Cbo.Clear
lst.Clear
If cant > 0 Then
    For i% = 0 To cant - 1
        Cbo.AddItem rs.Fields(1).Value
        lst.AddItem rs.Fields(0).Value
        rs.MoveNext
    Next
End If

End Sub
Public Sub Func_LlenarComboNoItem(ByRef Cbo As control, rs As Recordset)
Dim i As Integer
Dim cant As Integer
cant% = Aplicacion.CantReg(rs)
Cbo.Clear
If cant > 0 Then
    For i% = 0 To cant - 1
        Cbo.AddItem rs.Fields(0).Value
        rs.MoveNext
    Next
End If

End Sub



Private Sub Aux()
Dim SQL$
Dim rs As Recordset

SQL$ = ""
SQL$ = SQL$ & ""

If Aplicacion.ObtenerRsDAO(SQL$, rs) Then
    If Aplicacion.CantReg(rs) > 0 Then
    Else
    End If
    Aplicacion.CerrarDAO rs
End If

Aplicacion.ComienzoTrans

If Aplicacion.EjecutarDAO(SQL$) Then
    Aplicacion.TerminarConExitoTrans
Else
    Aplicacion.TerminarConErrorTrans
End If

End Sub

Public Function Insert_() As Boolean
Dim SQL As String

SQL = "INSERT INTO  ("
SQL = SQL & ","
SQL = SQL & ","
SQL = SQL & ","
SQL = SQL & ","
SQL = SQL & ","
SQL = SQL & ","
SQL = SQL & ","
SQL = SQL & ","
SQL = SQL & ","
SQL = SQL & ","
SQL = SQL & ","
SQL = SQL & " ) "
SQL = SQL & " VALUES ("
SQL = SQL & "'" & "" & "', "
SQL = SQL & "'" & "" & "', "
SQL = SQL & "'" & "" & "', "
SQL = SQL & "'" & "" & "', "
SQL = SQL & "'" & "" & "', "
SQL = SQL & "'" & "" & "', "
SQL = SQL & "'" & "" & "', "
SQL = SQL & "'" & "" & "', "
SQL = SQL & "'" & "" & "', "
SQL = SQL & "'" & "" & "', "
SQL = SQL & "'" & "" & "', "
SQL = SQL & "'" & "" & "', "
SQL = SQL & " ) "

Insert_ = Aplicacion.EjecutarDAO(SQL)

End Function
Public Function Func_ProximoValor(sec As String) As Long
Dim SQL As String
Dim rs As Recordset

SQL = "SELECT " & sec & ".nextval as Numero FROM Dual "

If Aplicacion.ObtenerRsDAO(SQL, rs) Then
    Func_ProximoValor = rs!numero
Else
    MsgBox "Existe algún problema con los números de códigos", vbOKOnly + vbExclamation, "ATENCION"
    
    Func_ProximoValor = -1
    
End If

End Function


Public Sub Func_MoverPrimero(ByRef rs As Recordset, ByRef pos As String)
rs.MoveFirst
pos = 1
End Sub
Public Sub Func_MoverUltimo(ByRef rs As Recordset, ByRef pos As String)
rs.MoveLast
pos = rs.RecordCount
End Sub
Public Sub Func_MoverSiguiente(ByRef rs As Recordset, ByRef pos As String)
rs.MoveNext
pos = pos + 1
End Sub
Public Sub Func_MoverAnterior(ByRef rs As Recordset, ByRef pos As String)
rs.MovePrevious
pos = pos - 1
End Sub


Public Sub Func_SetearCboINT(Cbo As control, dato As Variant)
Dim i%

For i% = 0 To Cbo.ListCount
    If Val(Cbo.List(i%)) = dato Then
        Cbo.ListIndex = i%
        Exit For
    End If
Next

End Sub
Public Sub Func_SetearCboSTR(Cbo As control, dato As Variant)
Dim i%
Dim ex As Boolean

ex = False
For i% = 0 To Cbo.ListCount
    If Cbo.List(i%) = dato Then
        Cbo.ListIndex = i%
        ex = True
        Exit For
    End If
Next
If Not ex Then
    Cbo.ListIndex = -1
End If
End Sub

Public Sub Func_SetearCboConLst(Cbo As control, lst As control, dato As Variant)
Dim i%

For i% = 0 To Cbo.ListCount - 1
    If lst.List(i%) = dato Then
        Cbo.ListIndex = i%
        lst.ListIndex = i%
        Exit For
    End If
Next

End Sub

Public Sub Func_SetearCboItem(Cbo As control, dato As Variant)
Dim i%

For i% = 0 To Cbo.ListCount
    If Val(Cbo.ItemData(i%)) = dato Then
        Cbo.ListIndex = i%
        Exit For
    End If
Next

End Sub

Public Sub Func_Setearlst(lst As control, dato As Variant)
Dim i%

If dato = lst.List(0) Then
    lst.ListIndex = 0
Else
    For i% = 1 To 10
        If dato - lst.List(i%) < 0 Then
            lst.ListIndex = i% - 1
            Exit For
        ElseIf dato - lst.List(i%) = 0 Then
            lst.ListIndex = i%
            Exit For
        End If
    Next
    If i% > 10 Then
        lst.ListIndex = 10
    End If
End If

'For i% = 0 To 10
'    If lst.ItemData(i%) = dato Then
'    If lst.List(i%) = dato Then
'        lst.ListIndex = i%
'        Exit For
'    End If
'Next

End Sub


