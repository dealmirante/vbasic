Attribute VB_Name = "Exl"
Option Explicit

'Hor
Global Const Exl_Izq = 1
Global Const Exl_Centro = 7
Global Const Exl_CentroCol = 3

'Vert
Global Const Exl_CentroVert = 2
Global Const Exl_TopVert = 1
Global Const Exl_DownVert = 3


Global Const Exl_Gris = 15
Global Const Exl_Negro = 1
Global Const Exl_Ros = 38
Global Const Exl_Cel = 24
Global Const Exl_Rojo = 3
Global Const Exl_Blanco = 2
Global Const Exl_AmaClaro = 19
Global Const Exl_AguaClaro = 20

Global Const NEGRITA = 1

Global Const Exl_LinDoble = 9
Global Const Exl_Linsimple = 1

Global Const Exl_Carta = 1
Global Const Exl_Legal = 5
Global Const Exl_A4 = 9

Global Const Exl_TopMargen = 40
Global Const Exl_BotMargen = 50

Global Const Exl_LArr = 3
Global Const Exl_LAba = 4
Global Const Exl_LIzq = 1
Global Const Exl_LDer = 2

Public Function Exl_AnchoCol(apl As Object, colD As Integer, colH As Integer, Ancho As Integer)

    apl.Application.Columns(Chr(colD + 64) & ":" & Chr(colH + 64)).ColumnWidth = Ancho

End Function
Public Sub Exl_Format(aplic As Object, rango As String)
    aplic.Application.Range(rango).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
End Sub

Public Sub Exl_Justificacion(apl As Object, rango As String, TipoH As Integer, TipoV As Integer, War As Boolean)
    apl.Application.Range(rango).HorizontalAlignment = TipoH
    apl.Application.Range(rango).VerticalAlignment = TipoV
    apl.Application.Range(rango).WrapText = War
End Sub
Public Sub Exl_ColorInt(apl As Object, rango As String, Tipo As Integer)
    apl.Application.Range(rango).Interior.ColorIndex = Tipo
End Sub

Public Sub Exl_Letra(apl As Object, rango As String, Tipo As Integer, tam As Integer, nombre As String)
    Select Case Tipo
        Case NEGRITA
            apl.Application.Range(rango).Font.Bold = True
            
    End Select
    apl.Application.Range(rango).Font.Size = tam
    apl.Application.Range(rango).Font.Name = nombre
    'apl.application.range(rango).Font.indexcolor = color
End Sub
Public Sub Exl_LetraColor(apl As Object, rango As String, color As Integer)
    apl.Application.Range(rango).Font.ColorIndex = color
End Sub

Public Sub Exl_PonerValor(apl As Object, F As Integer, c As Integer, ByVal dato As String)
If IsDate(dato) Then
    apl.Application.Cells(F, c).Value = " " & dato
Else
    apl.Application.Cells(F, c).Value = dato
End If

'AppExcel.application.Columns("C:C").PageBreak = True

End Sub

Public Function Exl_rangos(LI As Integer, LF As Integer, CI As Integer, CF As Integer) As String
Exl_rangos = Chr(64 + CI) & LI & ":" & Chr(64 + CF) & LF
End Function

Public Sub Exl_Lineas(apl As Object, rango As String, Tipo As String)

    apl.Application.Range(rango).Borders(1).LineStyle = Tipo 'xlContinuous
    apl.Application.Range(rango).Borders(2).LineStyle = Tipo
    apl.Application.Range(rango).Borders(3).LineStyle = Tipo
    apl.Application.Range(rango).Borders(4).LineStyle = Tipo

End Sub


Public Sub Exl_LineasPart(apl As Object, rango As String, Tipo As Integer, pos As Integer)
apl.Application.Range(rango).Borders(pos).LineStyle = Tipo
End Sub
Public Sub Exl_LineasAbj(apl As Object, rango As String, Tipo As Integer)
apl.Application.Range(rango).Borders(4).LineStyle = Tipo
End Sub

Public Sub Exl_BajarGrillaExel(spr As control, apl As Object, FilaInit As Integer, ColInit As Integer, titCol() As String)
Dim gr_fila As Integer
Dim gr_col As Integer, gr_aux_col As Integer
Dim dato As Variant
Dim rango As String



For gr_fila = 0 To spr.MaxRows
    gr_aux_col = 1
    For gr_col = 1 To spr.MaxCols
        spr.col = gr_col
        If Not spr.ColHidden Then
            If gr_fila = 0 Then
                apl.Application.Cells(FilaInit + gr_fila, gr_aux_col).Value = titCol(gr_col)
            Else
                spr.GetText gr_col, gr_fila, dato
                Exl_PonerValor apl, FilaInit + gr_fila, gr_aux_col, dato
                If IsNumeric(dato) Then
                If dato < 0 Then
                    rango = Exl_rangos(FilaInit + gr_fila, FilaInit + gr_fila, gr_aux_col, gr_aux_col)
                    Exl_LetraColor apl, rango, Exl_Rojo
                End If
                End If
            End If
            gr_aux_col = gr_aux_col + 1
        Else
            
        End If
    Next
Next
rango = Exl_rangos(FilaInit, FilaInit, ColInit, spr.MaxCols)

Exl_Justificacion apl, rango, Exl_CentroCol, Exl_CentroVert, False

Exl_ColorInt apl, rango, Exl_Gris

rango = Exl_rangos(FilaInit, spr.MaxRows + FilaInit, ColInit, spr.MaxCols)

Exl_Lineas apl, rango, Exl_Linsimple

End Sub



Public Function Exl_SeteoPagina(apl As Object)

With apl.Application.ActiveSheet.PageSetup
    .PrintTitleRows = "$1:$6"
    .PrintTitleColumns = ("$A:$A")
    .PaperSize = Exl_A4
    .CenterHorizontally = True
    .TopMargin = Exl_TopMargen
    .BottomMargin = Exl_BotMargen
    .CenterFooter = "Página &P de &N"
    .PrintGridlines = False
End With
'    ActiveSheet.PageSetup.PrintArea = ""
'    With ActiveSheet.PageSetup
'        .LeftHeader = ""
'        .CenterHeader = ""
'        .RightHeader = ""
'        .LeftFooter = ""
'        .CenterFooter = ""
'        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0)
'        .RightMargin = Application.InchesToPoints(0)
'        .HeaderMargin = Application.InchesToPoints(0)
'        .FooterMargin = Application.InchesToPoints(0)
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .PrintQuality = 300
'        .CenterHorizontally = False
'        .CenterVertically = False
'        .Orientation = xlPortrait
'        .Draft = False
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False
'        .Zoom = 100
'    End With

End Function


Private Sub OBJETO()
'Dim xlApp As excel.Application
'Dim xlBook As excel.Workbook
'Dim xlSheet As excel.Worksheet

'Set xlApp = CreateObject("Excel.Application")
'Set xlBook = xlApp.Workbooks.Add
'Set xlSheet = xlBook.Worksheets(1)

'Exl_PonerValor xlSheet, 2, 1, Titulo
    
'xlSheet.PageSetup.CenterHorizontally = True
'xlSheet.PageSetup.TopMargin = Exl_TopMargen
    
'    xlBook.SaveAs nombre & ".xls"
'    xlBook.Close
'
'    Set xlSheet = Nothing
'    Set xlBook = Nothing
'    Set xlApp = Nothing
End Sub


