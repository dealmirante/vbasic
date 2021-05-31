Attribute VB_Name = "FuncLocal"
Option Explicit



Public Sub FuncLocal_SacarForm(nom As String)
Dim i

For i = 0 To frmPrincipal.lstForms.ListCount - 1
    If frmPrincipal.lstForms.List(i) = nom Then
        frmPrincipal.lstForms.RemoveItem (i)
        Exit For
    End If
Next

End Sub


Public Sub FuncLocal_Porcentages(ByRef sprPorc As control, sprDato As control, Orientacion As String)
Dim i As Integer, j As Integer
Dim valor As Variant
Dim tot As Variant, fecha As Variant
Dim Result As Double

On Error GoTo ErrPorc:

sprPorc.MaxRows = 0

Select Case Orientacion
    Case "FILAS"
        For i = 1 To sprDato.MaxRows
            sprDato.GetText sprDato.MaxCols, i, tot
            sprDato.GetText 1, i, fecha
            Result = 0
            sprPorc.MaxRows = i
            
            sprPorc.SetText 1, i, Format$(fecha, "dd-mm-yy")
            sprPorc.SetText TotalINTA, i, "100"
            
            For j = 2 To sprDato.MaxCols - 1
                sprDato.GetText j, i, valor
                If tot <> "" Then
                    Result = valor * 100 / tot
                End If
                sprPorc.SetText j, i, Format$(Result, "% 0.00")
            Next
            
        Next
    Case "COL"
        For i = 1 To sprDato.MaxRows
         If i = sprDato.MaxRows Then
            sprPorc.MaxRows = i
            For j = 2 To sprDato.MaxCols
                sprPorc.SetText j, i, "100"
            Next
         Else
            sprDato.GetText 1, i, fecha
            Result = 0
            sprPorc.MaxRows = i
        
            sprPorc.SetText 1, i, Format$(fecha, "dd-mm-yy")
            
            For j = 2 To sprDato.MaxCols
                sprDato.GetText j, sprDato.MaxRows, tot
                sprDato.GetText j, i, valor
                If tot <> 0 Then
                    Result = valor * 100 / tot
                End If
                sprPorc.SetText j, i, Format$(Result, "% 0.00")
            Next
         End If
        Next

End Select
        Spread_TotalesLinea sprPorc
ErrPorc:
    Exit Sub
End Sub

Public Sub Func_CargarGraficos(GR As control, col As Collection)
Dim i
Dim cl_dato As CLlgi

'Graph1.Visible = False

GR.NumPoints = col.Count
GR.NumSets = 1

'GR.YAxisMax = 1
'GR.YAxisTicks = 2

GR.ThisSet = 1
i = 1
For Each cl_dato In col
    GR.ThisPoint = i

    GR.LegendText = cl_dato.Locale
    GR.GraphData = cl_dato.DatoGral
    i = i + 1
Next

GR.DrawMode = 2

End Sub




Public Sub FuncLocal_PromediosFilaSPR(ByVal SprDendo As control, ByVal SprSor As control, ByRef SprResult As control, ColHasta As String)
Dim i
Dim ValorSor As Variant
Dim ValorDendo As Variant
Dim Result As Double

If SprDendo.MaxCols = SprSor.MaxCols Then
    SprResult.MaxRows = SprResult.MaxRows + 1
    
    Spread_TotalesLinea SprResult

    For i = 2 To ColHasta
        SprDendo.GetText i, SprDendo.MaxRows, ValorDendo
        SprSor.GetText i, SprSor.MaxRows, ValorSor
        
        If Val(ValorSor) > 0 Then
            Result = ValorDendo / ValorSor
        Else
            Result = 0
        End If
        SprResult.SetText i, SprResult.MaxRows, Format$(Result, "#.0")
    Next
    
End If

End Sub

Public Function FuncLocal_Secuencia(anio As Integer, Mes As Integer, Tipo As String) As Long
Dim sql As String
Dim rs As Recordset


sql = "SELECT Estadis.Sec_Modelo ("
sql = sql & anio & ","
sql = sql & Mes & ","
sql = sql & "'" & Tipo & "') sec from dual "

If Aplicacion.ObtenerRsDAO(sql, rs) Then
    FuncLocal_Secuencia = rs!sec
    Aplicacion.CerrarDAO rs
End If


End Function

Public Sub FuncLocal_SeteoTABS(tabs As control)
Dim i

'tabs.TabVisible(0) = False
Select Case Aplicacion.Perfil
    Case "INTA"
        tabs.Tab = 1
        tabs.TabVisible(2) = False
        tabs.TabVisible(3) = False
    Case "INTB"
        tabs.Tab = 2
        tabs.TabVisible(1) = False
        tabs.TabVisible(3) = False
    Case "AEP"
        tabs.Tab = 3
        tabs.TabVisible(1) = False
        tabs.TabVisible(2) = False
    Case "INT"
        tabs.Tab = 4
        tabs.TabVisible(1) = False
        tabs.TabVisible(2) = False
End Select

End Sub

Public Function funcLocal_Vista(Prefijo As String, anio As Integer) As String
Dim Hist As String

Hist = ""
If anio <> Year(Date) Then
    Hist = "_hist"
End If

'If Aplicacion.Perfil = "TODOS" Then
    funcLocal_Vista = ESQUEMA & Prefijo & Hist
'Else
'    funcLocal_Vista = ESQUEMA & Prefijo & "_" & Aplicacion.Perfil & Hist
'End If

End Function



Public Function funcLocal_VistaTicket(Prefijo As String, Mes As Integer, anio As Integer) As String
Dim Hist As String

Hist = ""
If Mes <> Month(Date) Or anio <> Year(Date) Then
    Hist = "_Hist"
End If

'If Aplicacion.Perfil = "TODOS" Then
    funcLocal_VistaTicket = ESQUEMA & Prefijo & Hist
'Else
'    funcLocal_VistaTicket = ESQUEMA & Prefijo & "_" & Aplicacion.Perfil & Hist
'End If

End Function



