VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmResumenDiarioIFL 
   Caption         =   "Resumen día a día INFLIGHT"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   11580
   Begin VB.CommandButton botEjecutar 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   9540
      Picture         =   "frmResumenDiarioIFL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   150
      Width           =   600
   End
   Begin VB.CommandButton botEjecutar 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   8265
      Picture         =   "frmResumenDiarioIFL.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   600
   End
   Begin VB.CommandButton botEjecutar 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   0
      Left            =   7635
      Picture         =   "frmResumenDiarioIFL.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   150
      Width           =   615
   End
   Begin VB.CommandButton botExcel 
      Caption         =   "Excel"
      Height          =   525
      Left            =   8910
      Picture         =   "frmResumenDiarioIFL.frx":0A26
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   150
      Width           =   600
   End
   Begin VB.Frame frCab 
      Caption         =   "Datos de Cabecera"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   825
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   10275
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2535
         Picture         =   "frmResumenDiarioIFL.frx":0FB8
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   5850
         Picture         =   "frmResumenDiarioIFL.frx":112A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   330
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1365
         TabIndex        =   9
         Top             =   330
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFHasta 
         Height          =   285
         Left            =   4695
         TabIndex        =   10
         Top             =   345
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Desde"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Hasta"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3435
         TabIndex        =   11
         Top             =   345
         Width           =   1185
      End
   End
   Begin TabDlg.SSTab tabEst 
      Height          =   5070
      Left            =   30
      TabIndex        =   5
      Top             =   900
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8943
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   441
      ForeColor       =   255
      TabCaption(0)   =   "Ventas Diarias"
      TabPicture(0)   =   "frmResumenDiarioIFL.frx":129C
      Tab(0).ControlCount=   5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "sprIFL"
      Tab(0).Control(4).Enabled=   0   'False
      TabCaption(1)   =   "Ventas por Carro"
      TabPicture(1)   =   "frmResumenDiarioIFL.frx":12B8
      Tab(1).ControlCount=   0
      Tab(1).ControlEnabled=   0   'False
      Begin FPSpread.vaSpread sprIFL 
         Height          =   4140
         Left            =   75
         OleObjectBlob   =   "frmResumenDiarioIFL.frx":12D4
         TabIndex        =   6
         Top             =   630
         Width           =   11400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   2
         X1              =   4635
         X2              =   8715
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   1
         X1              =   8700
         X2              =   8700
         Y1              =   615
         Y2              =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   0
         X1              =   4620
         X2              =   4620
         Y1              =   600
         Y2              =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "MOVIMIENTOS DE STOCK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5370
         TabIndex        =   13
         Top             =   285
         Width           =   2385
      End
   End
End
Attribute VB_Name = "frmResumenDiarioIFL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsData As Recordset

Private Sub L_Grilla(spr As control, valor As Single)
    spr.MaxRows = spr.MaxRows + 1
    spr.SetText 1, spr.MaxRows, Trim(RsData!DM)
    spr.SetText 2, spr.MaxRows, str(RsData!Act)
    spr.SetText 3, spr.MaxRows, str(valor)

    Spread_TotalesGrillaAcum spr, 2, 6, spr.MaxRows
    Spread_TotalesGrillaAcum spr, 3, 7, spr.MaxRows
    
    spread_ResaltarCelda spr, 4, spr.MaxRows
    spread_ResaltarCelda spr, 5, spr.MaxRows
    spread_ResaltarCelda spr, 8, spr.MaxRows
    spread_ResaltarCelda spr, 9, spr.MaxRows

End Sub
Private Sub L_LlenarGrillas()
Dim linea As Integer

On Error GoTo ErrGLC:

linea = 1

Do While Not RsData.EOF
 sprIFL.MaxRows = sprIFL.MaxRows + 1
 sprIFL.SetText 1, linea, str(RsData!fch_vuelo)
 sprIFL.SetText 2, linea, str(IIf(IsNull(RsData!salidos), 0, RsData!salidos))
 sprIFL.SetText 3, linea, str(IIf(IsNull(RsData!arrivados), 0, RsData!arrivados))
 sprIFL.SetText 4, linea, str(IIf(IsNull(RsData!procesados), 0, RsData!procesados))
 sprIFL.SetText 5, linea, str(IIf(IsNull(RsData!despachados), 0, RsData!despachados))
 sprIFL.SetText 6, linea, str(IIf(IsNull(RsData!Importe), 0, RsData!Importe))
 sprIFL.SetText 7, linea, str(IIf(IsNull(RsData!Discrepancia), 0, RsData!Discrepancia))
 sprIFL.SetText 9, linea, str(IIf(IsNull(RsData!Recaudacion), 0, RsData!Recaudacion))
 RsData.MoveNext
 linea = linea + 1
Loop

ErrGLC:
     Exit Sub
End Sub
Private Sub L_LlenarAcumulados()
Dim n As Integer
Dim Ventas As Variant
Dim VentasAnteriores As Variant
Dim Discrepancias As Variant
Dim DiscrepanciasAnteriores As Variant
Dim Recaudacion As Variant
Dim RecaudacionAnteriores As Variant

On Error GoTo ErrGLC:

sprIFL.GetText 6, 1, Ventas
sprIFL.GetText 8, 1, Recaudacion
sprIFL.GetText 10, 1, Discrepancias
sprIFL.SetText 7, 1, str(Ventas)
sprIFL.SetText 9, 1, str(Recaudacion)
sprIFL.SetText 11, 1, str(Discrepancias)

For n = 2 To sprIFL.MaxRows
 sprIFL.GetText 6, n, Ventas
 sprIFL.GetText 8, n, Recaudacion
 sprIFL.GetText 10, n, Discrepancias
 sprIFL.GetText 7, n - 1, VentasAnteriores
 sprIFL.GetText 9, n - 1, RecaudacionAnteriores
 sprIFL.GetText 11, n - 1, DiscrepanciasAnteriores
 sprIFL.SetText 7, n, str(Val(Ventas) + Val(VentasAnteriores))
 sprIFL.SetText 9, n, str(Val(Recaudacion) + Val(RecaudacionAnteriores))
 sprIFL.SetText 11, n, str(Val(Discrepancias) + Val(DiscrepanciasAnteriores))
Next n

ErrGLC:
     Exit Sub
End Sub
Private Sub L_LimpiarGrillas()

sprIFL.MaxRows = 0
 
End Sub
Private Sub L_PlanillaExcel(Titulo As String, Tipo As String)
Dim AppExcel As Excel.Application
Dim libroExcel As Excel.Workbook
Dim hojaExcel As Excel.Worksheet
Dim i As Integer
Dim NOMBRE As String
Dim tit(1 To 2) As String

On Error GoTo ErrorExl:

tit(1) = " Período : " & Format$(mskFDesde.FormattedText, "dd-mmm-yyyy") & " - " & Format$(mskFHasta.FormattedText, "dd-mmm-yyyy")
tit(2) = " DIA A DIA POR CARRO "

NOMBRE = frmDir.NombreArchivo()
DoEvents

If NOMBRE <> "" Then
    frmResumenDiarioIFL.caption = Aplicacion.SeteoProceso(frmResumenDiarioIFL.caption)
    Set AppExcel = CreateObject("Excel.application")
    Set libroExcel = AppExcel.Workbooks.Add
    'AppExcel.Visible = True
    Select Case Tipo
        Case "DIARIO"
           Set hojaExcel = libroExcel.Worksheets(1)  'libroExcel.Worksheets(i)
           hojaExcel.Activate
           L_TratarExel hojaExcel, sprIFL, Titulo & tit(1), Tipo
        Case "CARROS"
           Set hojaExcel = libroExcel.Worksheets(1)  'libroExcel.Worksheets(i)
           hojaExcel.Activate
           L_TratarExel hojaExcel, sprIFL, Titulo & tit(2), Tipo
    End Select
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
      libroExcel.PrintOut
    End If
    
    libroExcel.SaveAs (NOMBRE & ".xls")
    AppExcel.Quit
    Set AppExcel = Nothing
    frmResumenDiarioIFL.caption = Aplicacion.SeteoFin
End If
Exit Sub

ErrorExl:
    frmResumenDiarioIFL.caption = Aplicacion.SeteoFin
    MsgBox "Planilla no generada", vbOKOnly + vbCritical, "ATENCION"
    Exit Sub
End Sub

Private Sub L_TitulosyFormatos(xls As Object, filaD As Integer, filaH As Integer, Tipo As String)
Dim rango As String
Dim i

    Exl.Exl_PonerValor xls, filaD, 1, " Fecha "
    Exl.Exl_PonerValor xls, filaD, 2, " Carros "
    Exl.Exl_PonerValor xls, filaD, 3, " Arrivados "
    Exl.Exl_PonerValor xls, filaD, 4, " Procesados "
    Exl.Exl_PonerValor xls, filaD, 5, " Tránsitos "
    Exl.Exl_PonerValor xls, filaD, 6, " Facturado "
    Exl.Exl_PonerValor xls, filaD, 7, " Discrepancias "
    Exl.Exl_PonerValor xls, filaD, 8, " Total "
    Exl.Exl_PonerValor xls, filaD, 9, " Recaudado S/Com "
    Exl.Exl_PonerValor xls, filaD, 10, " Sobr/Falt. $ "

    rango = Exl_rangos(filaD, filaH, 2, 2)
    Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LIzq


    'For i = 0 To 2
        rango = Exl_rangos(filaD, filaD, 2, 5)
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LArr
        rango = Exl_rangos(filaD, filaH, 5, 5)
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LDer
        rango = Exl_rangos(filaH, filaH, 2, 5)
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LAba
    
        rango = Exl_rangos(filaD, filaD, 5, 8)
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LArr
        rango = Exl_rangos(filaD, filaH, 8, 8)
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LDer
        rango = Exl_rangos(filaH, filaH, 5, 8)
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LAba
    
        rango = Exl_rangos(filaD, filaD, 8, 10)
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LArr
        rango = Exl_rangos(filaD, filaH, 10, 10)
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LDer
        rango = Exl_rangos(filaH, filaH, 8, 10)
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LAba
    
    'Next
    'Lineas y formatos de letras
        rango = Exl_rangos(filaD, filaD, 2, 10)
        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LAba
        Exl.Exl_Justificacion xls, rango, Exl_CentroCol, Exl_DownVert, True
        Exl.Exl_Letra xls, rango, NEGRITA, 8, "Ms Serif"
        rango = Exl_rangos(filaD, filaH, 1, 1)
        Exl.Exl_Letra xls, rango, NEGRITA, 8, "Ms Serif"
        rango = Exl_rangos(sprIFL.MaxRows + 4, sprIFL.MaxRows + 4, 2, 10)
        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LArr

'        rango = Exl_rangos(filaD, filaH, 4, 4)
'        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LDer
'        rango = Exl_rangos(filaD, filaH, 10, 10)
'        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LDer
'        rango = Exl_rangos(filaD, filaH, 16, 16)
'        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LDer
    'Achos de columnas
        Exl.Exl_AnchoCol xls, 6, 6, 13
        Exl.Exl_AnchoCol xls, 7, 7, 13
        Exl.Exl_AnchoCol xls, 8, 8, 13
        Exl.Exl_AnchoCol xls, 9, 9, 13
        Exl.Exl_AnchoCol xls, 10, 10, 13
'    '    Exl.Exl_AnchoCol xls, 16, 16, 6
'    '    Exl.Exl_AnchoCol xls, 19, 19, 6
'        Exl.Exl_AnchoCol xls, 8, 9, 9
'        Exl.Exl_AnchoCol xls, 11, 12, 9
'        Exl.Exl_AnchoCol xls, 14, 19, 7
    
'        rango = Exl_rangos(filaD + 1, filaH, 2, 19)
'        xls.Range(rango).Font.Size = 8
    'formato de columnas de difencias
'        rango = Exl_rangos(filaD + 1, filaH, 4, 4)
'        Exl.Exl_Format xls, rango
'        rango = Exl_rangos(filaD + 1, filaH, 7, 7)
'        Exl.Exl_Format xls, rango
'        rango = Exl_rangos(filaD + 1, filaH, 10, 10)
'        Exl.Exl_Format xls, rango
'        rango = Exl_rangos(filaD + 1, filaH, 13, 13)
'        Exl.Exl_Format xls, rango
'        rango = Exl_rangos(filaD + 1, filaH, 16, 16)
'        Exl.Exl_Format xls, rango
'        rango = Exl_rangos(filaD + 1, filaH, 19, 19)
'        Exl.Exl_Format xls, rango
    'formato de las columnas de importes
        rango = Exl_rangos(filaD + 1, filaH, 6, 10)
        Exl.Exl_Format xls, rango
'        rango = Exl_rangos(filaD + 1, filaH, 5, 6)
'        Exl.Exl_Format xls, rango


End Sub

Private Sub L_TratarExel(hojaExcel As Object, spr As control, Titulo As String, Tipo As String)
Dim i As Integer, fila As Integer
Dim rango As String
Dim valor As Variant


fila = 2
Exl_PonerValor hojaExcel, fila, 1, Titulo
rango = Exl_rangos(fila, fila, 1, 10)
Exl_Justificacion hojaExcel, rango, Exl_Centro, Exl_CentroVert, False
rango = Exl_rangos(fila, fila, 1, 1)
Exl.Exl_Letra hojaExcel, rango, NEGRITA, 12, "Ms Serif"

fila = fila + 2

For i = 1 To spr.MaxRows
    spr.GetText 1, i, valor
    hojaExcel.Cells(fila + i, 1).Value = "" & Format(valor, "dd-mmm-yyyy")
Next
hojaExcel.Cells(fila + i - 1, 1).Value = "TOTAL"

L_TitulosyFormatos hojaExcel, fila, i + fila - 1, Tipo

For i = 1 To spr.MaxRows
    spr.GetText 2, i, valor
    Exl_PonerValor hojaExcel, fila + i, 2, valor
    spr.GetText 3, i, valor
    Exl_PonerValor hojaExcel, fila + i, 3, valor
    spr.GetText 4, i, valor
    Exl_PonerValor hojaExcel, fila + i, 4, valor
    spr.GetText 5, i, valor
    Exl_PonerValor hojaExcel, fila + i, 5, Format$(valor, "#.00")
    spr.GetText 6, i, valor
    Exl_PonerValor hojaExcel, fila + i, 6, valor
    spr.GetText 7, i, valor
    Exl_PonerValor hojaExcel, fila + i, 7, valor
    spr.GetText 8, i, valor
    Exl_PonerValor hojaExcel, fila + i, 8, valor
    spr.GetText 9, i, valor
    Exl_PonerValor hojaExcel, fila + i, 9, Format$(valor, "#.00")
    spr.GetText 10, i, valor
    Exl_PonerValor hojaExcel, fila + i, 10, Format$(valor, "#.00")

Next
hojaExcel.PageSetup.Orientation = 2
hojaExcel.PageSetup.Zoom = False

End Sub

Private Sub botEjecutar_Click(Index As Integer)
Dim i

Select Case Index
    Case 0
      frmResumenDiarioIFL.caption = Aplicacion.SeteoProceso(frmResumenDiarioIFL.caption)
        
      L_Refrescar
      L_PintarfinSemana_Esp sprIFL
      Spread.Spread_TotalesGrillas sprIFL, 9, 1
       
      botEjecutar(0).Enabled = False
      frmResumenDiarioIFL.caption = Aplicacion.SeteoFin
    Case 1
      frCab.Enabled = True
      botEjecutar(0).Enabled = True
      L_LimpiarGrillas
    Case 2
      Unload Me
End Select

End Sub

Private Function L_Armarcondicion()
Dim Cond As String
Dim fechaDesde As String
Dim fechaHasta As String

fechaDesde = mskFDesde.FormattedText
fechaHasta = mskFHasta.FormattedText

Cond = " WHERE V.fch_vuelo between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

L_Armarcondicion = Cond


End Function

Private Sub L_Refrescar()
Dim sql As String, sqlT As String

On Error GoTo ErrEstim:

sqlT = " SELECT V.FCH_VUELO,A.SALIDOS,B.DESPACHADOS,C.ARRIVADOS,V.PROCESADOS,IMPORTE,D.RECAUDACION,DISCREPANCIA"
sqlT = sqlT & " FROM   INFLIGHT.VIEW_CARROS_DIARIOS V,"
sqlT = sqlT & "        (SELECT FCH_VUELO,COUNT(*) SALIDOS"
sqlT = sqlT & "         FROM INFLIGHT.FL_TRANSITO"
sqlT = sqlT & "         GROUP BY FCH_VUELO) A,"
sqlT = sqlT & "        (SELECT FCH_VUELO,COUNT(*) DESPACHADOS"
sqlT = sqlT & "         FROM INFLIGHT.FL_CARRO"
sqlT = sqlT & "         WHERE COD_ESTADO = 3"
sqlT = sqlT & "         GROUP BY FCH_VUELO) B,"
sqlT = sqlT & "        (SELECT FCH_VUELO,COUNT(*) ARRIVADOS"
sqlT = sqlT & "         FROM INFLIGHT.FL_CARRO"
sqlT = sqlT & "         WHERE  COD_ESTADO IN(4,5)"
sqlT = sqlT & "         GROUP BY FCH_VUELO) C, "
sqlT = sqlT & "        (SELECT FECHA,SUM(RECAUDADO_USD) RECAUDACION"
sqlT = sqlT & "         FROM INFLIGHT.VIEW_RECAUDACION"
sqlT = sqlT & "         GROUP BY FECHA) D"
sqlT = sqlT & L_Armarcondicion
sqlT = sqlT & "  AND  V.FCH_VUELO = A.FCH_VUELO(+)"
sqlT = sqlT & "  AND  V.FCH_VUELO = B.FCH_VUELO(+)"
sqlT = sqlT & "  AND  V.FCH_VUELO = C.FCH_VUELO(+)"
sqlT = sqlT & "  AND  V.FCH_VUELO = D.FECHA(+)"
sqlT = sqlT & "  AND  V.ESTADO = 6"

If Aplicacion.ObtenerRsDAO(sqlT, RsData) Then
    
    L_LlenarGrillas
    Aplicacion.CerrarDAO RsData

End If

ErrEstim:
    frmResumenDiarioIFL.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub


Private Sub botExcel_Click()

Select Case tabEst.Tab
    Case 0
        L_PlanillaExcel " VENTA A BORDO DIA A DIA ", "DIARIO"
    Case 1
        L_PlanillaExcel " VENTA A BORDO POR DIA Y POR CARRO ", "CARROS"
End Select

End Sub

Private Sub botHelpFD_Click()
Dim fch As Date

If mskFDesde.Text <> "" Then
    fch = mskFDesde.FormattedText
Else
    fch = Date
End If

frmFecha.MuestroFormFecha fch

mskFDesde.Text = Format$(fch, FTOFECHA)

mskFDesde.SetFocus
End Sub

Private Sub botHelpFH_Click()
Dim fch As Date

If mskFHasta.Text <> "" Then
    fch = mskFHasta.FormattedText
Else
    fch = Date
End If

frmFecha.MuestroFormFecha fch

mskFHasta.Text = Format$(fch, FTOFECHA)

mskFHasta.SetFocus
End Sub

Private Sub Form_Load()
Dim sql As String

Top = 30
Left = 210
Height = 7000
Width = 11700

tabEst.TabVisible(1) = False

mskFDesde.Text = func_Dia1SegunMes_Anio(Month(Date), Year(Date))
If CDate(mskFDesde.FormattedText) = Date Then
    mskFHasta.Text = Format$(Date, FTOFECHA)
Else
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
End If

L_LimpiarGrillas

End Sub
Public Sub L_PintarfinSemana_Esp(ByRef spr As control)
Dim i As Integer
Dim valor As Variant

For i = 1 To spr.MaxRows
    spr.GetText 1, i, valor
    If IsDate(valor) Then
        If WeekDay(CDate(valor)) = DOMINGO Or WeekDay(CDate(valor)) = SABADO Then
            Spread_PintaLinea spr, i, 1, i, spr.MaxCols
        End If
    End If
Next

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FuncLocal_SacarForm "frmVsDia"
End Sub

Private Sub mskFDesde_LostFocus()
    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
     Else

    End If
    
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
    If mskFHasta.Text <> "" Then
        If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
            mskFHasta.Text = mskFDesde.Text
        End If
    End If

End Sub

Private Sub mskFHasta_LostFocus()
If Not IsDate(mskFHasta.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
ElseIf CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Or Year(mskFHasta.FormattedText) <> Year(mskFDesde.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
End If

mskFHasta.Text = Format$(mskFHasta.FormattedText, FTOFECHA)

End Sub

Private Sub tabEst_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    Select Case tabEst.Tab
        Case 0
'            sprTotal(tabTotal.Tab).SetFocus
        Case 1
'            sprGA(tabGA.Tab).SetFocus
        Case 2
'            sprGB(tabGB.Tab).SetFocus
        Case 3
'            sprGC(tabGC.Tab).SetFocus
    End Select
    
    
ErrT:
    Exit Sub

End Sub
