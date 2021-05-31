VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPuntosNac 
   Caption         =   "Puntos por Nacionalidad"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   9195
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
      Height          =   480
      Index           =   2
      Left            =   8400
      Picture         =   "frmPuntosNac.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   180
      Width           =   570
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
      Height          =   480
      Index           =   1
      Left            =   7170
      Picture         =   "frmPuntosNac.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   180
      Width           =   570
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
      Height          =   495
      Index           =   0
      Left            =   6555
      Picture         =   "frmPuntosNac.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   180
      Width           =   570
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   3915
      Left            =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   6906
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   459
      Enabled         =   0   'False
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "COMPAÑIA"
      TabPicture(0)   =   "frmPuntosNac.frx":0A26
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frNum(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SprTot"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "EZEIZA"
      TabPicture(1)   =   "frmPuntosNac.frx":0A42
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sprEzeA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frNum(0)"
      Tab(1).Control(1).Enabled=   0   'False
      TabCaption(2)   =   "AEROPARQUE"
      TabPicture(2)   =   "frmPuntosNac.frx":0A5E
      Tab(2).ControlCount=   2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "sprAEP"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frNum(1)"
      Tab(2).Control(1).Enabled=   0   'False
      TabCaption(3)   =   "INTERIOR"
      TabPicture(3)   =   "frmPuntosNac.frx":0A7A
      Tab(3).ControlCount=   2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "sprInt"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "frNum(2)"
      Tab(3).Control(1).Enabled=   0   'False
      Begin FPSpread.vaSpread sprEzeA 
         Height          =   2775
         Left            =   -74790
         OleObjectBlob   =   "frmPuntosNac.frx":0A96
         TabIndex        =   15
         Top             =   585
         Width           =   8670
      End
      Begin FPSpread.vaSpread sprAEP 
         Height          =   2775
         Left            =   -74790
         OleObjectBlob   =   "frmPuntosNac.frx":0F7E
         TabIndex        =   16
         Top             =   615
         Width           =   8670
      End
      Begin FPSpread.vaSpread sprInt 
         Height          =   2775
         Left            =   -74805
         OleObjectBlob   =   "frmPuntosNac.frx":1466
         TabIndex        =   17
         Top             =   570
         Width           =   8670
      End
      Begin FPSpread.vaSpread SprTot 
         Height          =   2775
         Left            =   225
         OleObjectBlob   =   "frmPuntosNac.frx":194E
         TabIndex        =   18
         Top             =   570
         Width           =   8670
      End
      Begin VB.Frame frNum 
         Height          =   3105
         Index           =   3
         Left            =   135
         TabIndex        =   14
         Top             =   390
         Width           =   8865
      End
      Begin VB.Frame frNum 
         Height          =   3150
         Index           =   1
         Left            =   -74900
         TabIndex        =   12
         Top             =   405
         Width           =   8865
      End
      Begin VB.Frame frNum 
         Height          =   3150
         Index           =   2
         Left            =   -74900
         TabIndex        =   13
         Top             =   375
         Width           =   8880
      End
      Begin VB.Frame frNum 
         Height          =   3150
         Index           =   0
         Left            =   -74895
         TabIndex        =   11
         Top             =   405
         Width           =   8910
      End
   End
   Begin VB.Frame frdatos 
      Height          =   945
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   9030
      Begin VB.CommandButton BotExcelResumen 
         Caption         =   "Excel"
         Height          =   480
         Left            =   7770
         Picture         =   "frmPuntosNac.frx":1E36
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   570
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   5895
         Picture         =   "frmPuntosNac.frx":23C8
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   285
         Width           =   375
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2625
         Picture         =   "frmPuntosNac.frx":253A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   315
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1455
         TabIndex        =   4
         Top             =   345
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFHasta 
         Height          =   285
         Left            =   4740
         TabIndex        =   5
         Top             =   300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
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
         Left            =   3540
         TabIndex        =   7
         Top             =   300
         Width           =   1140
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
         Left            =   255
         TabIndex        =   6
         Top             =   345
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmPuntosNac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset




Private Sub L_LimpiarGrillas()
Dim i

sprEzeA.MaxRows = 0
sprInt.MaxRows = 0
sprAep.MaxRows = 0
SprTot.MaxRows = 0

End Sub

Private Sub L_Participacion(ByRef spr As control)
Dim i As Integer
Dim valor As Variant
Dim tot As Variant, fecha As Variant
Dim Result As Double

On Error GoTo ErrPorc:

spr.GetText 2, spr.MaxRows, tot

For i = 1 To spr.MaxRows
    spr.GetText 2, i, valor
    
    spr.SetText 6, i, str(valor / tot * 100)
Next


ErrPorc:
    Exit Sub
End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrEA:

frmPuntosNac.caption = Aplicacion.SeteoProceso(frmPuntosNac.caption)

sql = " SELECT COD_DEPN, "
sql = sql & " NACIONALIDAD, "
sql = sql & " DESCRIP, "
sql = sql & " sum(CANT_TICKET) TICKET, "
sql = sql & " sum(IMPORTE) IMPORTE, "
sql = sql & " sum(DESCUENTOS) DESCUENTOs"
sql = sql & " FROM ESTADIS.VENTAS_PUNTOS P, VENTAS.PAISES S "
sql = sql & " Where nacionalidad = COD_PAIS "
sql = sql & " AND FCH_TICKET BETWEEN " & Func.func_ToDate(mskFDesde.FormattedText) & " AND " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " GROUP BY COD_DEPN, NACIONALIDAD,DESCRIP "
sql = sql & " Union "
sql = sql & " SELECT 'TOT', "
sql = sql & " NACIONALIDAD, "
sql = sql & " DESCRIP, "
sql = sql & " sum(CANT_TICKET), "
sql = sql & " sum(IMPORTE), "
sql = sql & " Sum (DESCUENTOS)"
sql = sql & " FROM ESTADIS.VENTAS_PUNTOS P, VENTAS.PAISES S"
sql = sql & " Where nacionalidad = COD_PAIS "
sql = sql & " AND FCH_TICKET BETWEEN " & Func.func_ToDate(mskFDesde.FormattedText) & " AND " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " GROUP BY NACIONALIDAD,DESCRIP "


If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        L_Llenargrilla
        'L_Porcentages sprPorcA, sprEzeA, "COL"
    End If

    Aplicacion.CerrarDAO RsData
    
End If

ErrEA:
    frmPuntosNac.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub

Private Sub L_SeteosGral()
Dim i

    frdatos.Enabled = False
    botEjecutar(0).Enabled = False
    tabEspigon.Enabled = True

End Sub

Private Sub botEjecutar_Click(Index As Integer)

Select Case Index
    
    Case 0
        L_Refrescar
        tabEspigon.Enabled = True
        botEjecutar(0).Enabled = False
    Case 1
        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        tabEspigon.Enabled = False
        L_LimpiarGrillas
    Case 2
        Unload Me
End Select

End Sub




Private Sub L_TratarExcel(titulo As String, subTit As String, Info As String)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim i
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmPuntosNac.caption = Aplicacion.SeteoProceso(frmPuntosNac.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.application.Visible = True
    
    ReDim titCol(SprTot.MaxCols)
    Col = 1
    fila = 4
    
    For i = 1 To SprTot.MaxCols
        SprTot.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 2, 1, titulo
    rango = Exl_rangos(2, 2, 1, SprTot.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, SprTot.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, 1, "EZEIZA"
    rango = Exl_rangos(fila, fila, 1, 1)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    fila = fila + 1
    Exl_BajarGrillaExel sprEzeA, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeA.MaxRows, Col + 1, sprEzeA.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeA.MaxRows
    rango = Exl_rangos(fila, fila, Col, sprEzeA.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, 1, "AEROPARQUE"
    rango = Exl_rangos(fila, fila, 1, 1)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    fila = fila + 1
    Exl_BajarGrillaExel sprAep, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprAep.MaxRows, Col + 1, sprAep.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprAep.MaxRows
    rango = Exl_rangos(fila, fila, Col, sprAep.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, 1, "INTERIOR"
    rango = Exl_rangos(fila, fila, 1, 1)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    fila = fila + 1
    Exl_BajarGrillaExel sprInt, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprInt.MaxRows, Col + 1, sprInt.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprInt.MaxRows
    rango = Exl_rangos(fila, fila, Col, sprInt.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, 1, "TOTAL COMPAÑIA"
    rango = Exl_rangos(fila, fila, 1, 1)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    fila = fila + 1
    Exl_BajarGrillaExel SprTot, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + SprTot.MaxRows, Col + 1, SprTot.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + SprTot.MaxRows
    rango = Exl_rangos(fila, fila, Col, SprTot.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    'AppExcel.application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If

    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmPuntosNac.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Function L_NombreDato(i As Integer) As String
Select Case i
    Case 1
        L_NombreDato = "Importes"
    Case 2
        L_NombreDato = "Tickets"
    Case 3
        L_NombreDato = "Pasajeros"
    Case 4
        L_NombreDato = "Promedios por Ticket"
    Case 5
        L_NombreDato = "Promedios por Pasajeros"
End Select
End Function

Private Sub BotExcelResumen_Click()
    L_TratarExcel "Informe Estadístico Duty Free Point", "por Nacionalidad (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", Total
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



Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabEspigon


End Sub

Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6000
Me.Width = 9300


mskFDesde.Text = func_Dia1SegunMes_Anio(Month(Date), Year(Date))
If CDate(mskFDesde.FormattedText) = Date Then
    mskFHasta.Text = Format$(Date, FTOFECHA)
Else
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
End If


L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmLocNac"

End Sub




Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmGLC"
End Sub


Private Sub mskFDesde_LostFocus()
    
    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Date
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
ElseIf CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
End If

mskFHasta.Text = Format$(mskFHasta.FormattedText, FTOFECHA)

End Sub


Private Sub L_Llenargrilla()
Dim fecha As String
Dim i As Integer
Dim Total As Variant
Dim TOTImp As Long
Dim totCant As Long


Do While Not RsData.EOF
    Select Case RsData!cod_depn
           Case "EZE"
            sprEzeA.MaxRows = sprEzeA.MaxRows + 1
            
            sprEzeA.SetText 1, sprEzeA.MaxRows, Trim(RsData!Descrip)
            sprEzeA.SetText 2, sprEzeA.MaxRows, str(RsData!Importe)
            sprEzeA.SetText 3, sprEzeA.MaxRows, str(RsData!ticket)
            sprEzeA.SetText 4, sprEzeA.MaxRows, str(RsData!descuentos)
           
           Case "AEP"
            sprAep.MaxRows = sprAep.MaxRows + 1
           
            sprAep.SetText 1, sprAep.MaxRows, Trim(RsData!Descrip)
            sprAep.SetText 2, sprAep.MaxRows, str(RsData!Importe)
            sprAep.SetText 3, sprAep.MaxRows, str(RsData!ticket)
            sprAep.SetText 4, sprAep.MaxRows, str(RsData!descuentos)
           
           Case "INT"
            sprInt.MaxRows = sprInt.MaxRows + 1
           
            sprInt.SetText 1, sprInt.MaxRows, Trim(RsData!Descrip)
            sprInt.SetText 2, sprInt.MaxRows, str(RsData!Importe)
            sprInt.SetText 3, sprInt.MaxRows, str(RsData!ticket)
            sprInt.SetText 4, sprInt.MaxRows, str(RsData!descuentos)
           
           Case "TOT"
            SprTot.MaxRows = SprTot.MaxRows + 1
    
            SprTot.SetText 1, SprTot.MaxRows, Trim(RsData!Descrip)
            SprTot.SetText 2, SprTot.MaxRows, str(RsData!Importe)
            SprTot.SetText 3, SprTot.MaxRows, str(RsData!ticket)
            SprTot.SetText 4, SprTot.MaxRows, str(RsData!descuentos)
    
    End Select
        
    RsData.MoveNext
Loop

Spread_TotalesGrillas sprEzeA, sprEzeA.MaxCols - 3, 2
Spread_TotalesGrillas sprAep, sprAep.MaxCols - 3, 2
Spread_TotalesGrillas sprInt, sprInt.MaxCols - 3, 2
Spread_TotalesGrillas SprTot, SprTot.MaxCols - 3, 2

L_Participacion sprEzeA
L_Participacion sprAep
L_Participacion SprTot
L_Participacion sprInt

End Sub

