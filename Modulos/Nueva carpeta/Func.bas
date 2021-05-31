Attribute VB_Name = "Func"
Option Explicit

Public Function Func_ObtenerDesc(ByVal sql As String, ByRef desc As String) As Boolean
Dim Rs As Recordset

If Aplicacion.ObtenerRsDAO(sql$, Rs) Then
    If Aplicacion.CantReg(Rs) > 0 Then
        Func_ObtenerDesc = True
        desc = Rs!Descrip
    Else
        Func_ObtenerDesc = False
        desc = ""
    End If
    Aplicacion.CerrarDAO Rs
End If

End Function


Public Function Func_CantidadDias(fch As String) As Integer
Dim fchdesde As String
Dim fchHasta As String

fchdesde = func_Dia1SegunMes_Anio(Month(Format$(fch, FTOFECHA)), Year(Format$(fch, FTOFECHA)))
fchHasta = func_Dia30SegunMes_Anio(Month(Format$(fch, FTOFECHA)), Year(Format$(fch, FTOFECHA)))

Func_CantidadDias = CDate(fchHasta) - CDate(fchdesde) + 1

End Function

Public Function func_Dia1SegunMes_Anio(mes As Integer, anio As Integer)
func_Dia1SegunMes_Anio = "01-" & Format(mes, "0#") & "-" & Trim(str(anio))
End Function

Public Function func_Dia30SegunMes_Anio(mes As Integer, anio As Integer)
Dim MesAux, AnioAux
Dim fecha

If mes = 12 Then
    MesAux = 1
    AnioAux = anio + 1
Else
    MesAux = mes + 1
    AnioAux = anio
End If

fecha = "01-" & Format(MesAux, "0#") & "-" & Trim(str(AnioAux))

func_Dia30SegunMes_Anio = Format$(CDate(fecha) - 1, FTOFECHA)

End Function


Public Function func_ToDate(fch As String) As String

func_ToDate = "To_date ('" & fch & "','" & FTOFECHA & "')"

End Function

Public Function Update_() As Boolean
Dim sql As String

sql = "UPDATE "
sql = sql & " XXXX SET "
sql = sql & " xxx  =" & "'" & "" & "',"
sql = sql & " xxx  =" & "'" & "" & "',"
sql = sql & " xxx  =" & "'" & "" & "',"
sql = sql & " xxx  =" & "'" & "" & "',"
sql = sql & " xxx  =" & "'" & "" & "',"
sql = sql & " WHERE CODIGO = '" & "" & "'"

Update_ = Aplicacion.EjecutarDAO(sql)

End Function

Public Sub Func_LlenarCombo(Cbo As Control, Rs As Recordset)
Dim i As Integer
Dim cant As Integer

cant% = Aplicacion.CantReg(Rs)
Cbo.Clear

If cant > 0 Then
    For i% = 0 To cant - 1
        Cbo.AddItem Rs.Fields(1).Value
        Cbo.ItemData(Cbo.NewIndex) = Rs.Fields(0).Value
        Rs.MoveNext
    Next
End If

End Sub


Public Sub Func_LlenarComboLst(ByRef Cbo As Control, ByRef lst As Control, Rs As Recordset)
Dim i As Integer
Dim cant As Integer
cant% = Aplicacion.CantReg(Rs)
Cbo.Clear
lst.Clear
If cant > 0 Then
    For i% = 0 To cant - 1
        Cbo.AddItem Rs.Fields(1).Value
        lst.AddItem Rs.Fields(0).Value
        Rs.MoveNext
    Next
End If

End Sub
Public Sub Func_LlenarComboNoItem(ByRef Cbo As Control, Rs As Recordset)
Dim i As Integer
Dim cant As Integer
cant% = Aplicacion.CantReg(Rs)
Cbo.Clear
If cant > 0 Then
    For i% = 0 To cant - 1
        Cbo.AddItem Rs.Fields(0).Value
        Rs.MoveNext
    Next
End If

End Sub



Private Sub Aux()
Dim sql$
Dim Rs As Recordset

sql$ = ""
sql$ = sql$ & ""

If Aplicacion.ObtenerRsDAO(sql$, Rs) Then
    If Aplicacion.CantReg(Rs) > 0 Then
    Else
    End If
    Aplicacion.CerrarDAO Rs
End If

Aplicacion.ComienzoTrans

If Aplicacion.EjecutarDAO(sql$) Then
    Aplicacion.TerminarConExitoTrans
Else
    Aplicacion.TerminarConErrorTrans
End If

End Sub

Public Function Insert_() As Boolean
Dim sql As String

sql = "INSERT INTO  ("
sql = sql & ","
sql = sql & ","
sql = sql & ","
sql = sql & ","
sql = sql & ","
sql = sql & ","
sql = sql & ","
sql = sql & ","
sql = sql & ","
sql = sql & ","
sql = sql & ","
sql = sql & " ) "
sql = sql & " VALUES ("
sql = sql & "'" & "" & "', "
sql = sql & "'" & "" & "', "
sql = sql & "'" & "" & "', "
sql = sql & "'" & "" & "', "
sql = sql & "'" & "" & "', "
sql = sql & "'" & "" & "', "
sql = sql & "'" & "" & "', "
sql = sql & "'" & "" & "', "
sql = sql & "'" & "" & "', "
sql = sql & "'" & "" & "', "
sql = sql & "'" & "" & "', "
sql = sql & "'" & "" & "', "
sql = sql & " ) "

Insert_ = Aplicacion.EjecutarDAO(sql)

End Function
Public Function Func_ProximoValor(Sec As String) As Long
Dim sql As String
Dim Rs As Recordset

sql = "SELECT " & Sec & ".nextval as Numero FROM Dual "

If Aplicacion.ObtenerRsDAO(sql, Rs) Then
    Func_ProximoValor = Rs!numero
Else
    MsgBox "Existe algún problema con los números de códigos", vbOKOnly + vbExclamation, "ATENCION"
    
    Func_ProximoValor = -1
    
End If

End Function


Public Sub Func_MoverPrimero(ByRef Rs As Recordset, ByRef pos As String)
Rs.MoveFirst
pos = 1
End Sub
Public Sub Func_MoverUltimo(ByRef Rs As Recordset, ByRef pos As String)
Rs.MoveLast
pos = Rs.RecordCount
End Sub
Public Sub Func_MoverSiguiente(ByRef Rs As Recordset, ByRef pos As String)
Rs.MoveNext
pos = pos + 1
End Sub
Public Sub Func_MoverAnterior(ByRef Rs As Recordset, ByRef pos As String)
Rs.MovePrevious
pos = pos - 1
End Sub


Public Sub Func_SetearCboINT(Cbo As Control, dato As Variant)
Dim i%

For i% = 0 To Cbo.ListCount
    If Val(Cbo.List(i%)) = dato Then
        Cbo.ListIndex = i%
        Exit For
    End If
Next

End Sub
Public Sub Func_SetearCboSTR(Cbo As Control, dato As Variant)
Dim i%

For i% = 0 To Cbo.ListCount
    If Cbo.List(i%) = dato Then
        Cbo.ListIndex = i%
        Exit For
    End If
Next

End Sub

Public Sub Func_SetearCboConLst(Cbo As Control, lst As Control, dato As Variant)
Dim i%

For i% = 0 To Cbo.ListCount - 1
    If lst.List(i%) = dato Then
        Cbo.ListIndex = i%
        lst.ListIndex = i%
        Exit For
    End If
Next

End Sub

Public Sub Func_SetearCboItem(Cbo As Control, dato As Variant)
Dim i%

For i% = 0 To Cbo.ListCount
    If Val(Cbo.ItemData(i%)) = dato Then
        Cbo.ListIndex = i%
        Exit For
    End If
Next

End Sub

Public Sub Func_Setearlst(lst As Control, dato As Variant)
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


