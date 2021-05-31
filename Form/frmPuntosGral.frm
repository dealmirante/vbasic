VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPuntosGral 
   Caption         =   "Estadísticas Generales"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   9195
   Begin VB.TextBox txtSocios 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   360
      Left            =   2730
      TabIndex        =   34
      Top             =   5055
      Width           =   2565
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
      Index           =   2
      Left            =   8400
      Picture         =   "frmPuntosGral.frx":0000
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
      Picture         =   "frmPuntosGral.frx":0822
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
      Picture         =   "frmPuntosGral.frx":0924
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
      TabPicture(0)   =   "frmPuntosGral.frx":0A26
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frNum(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "sprGral(0)"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "EZEIZA"
      TabPicture(1)   =   "frmPuntosGral.frx":0A42
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frNum(0)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "AEROPARQUE"
      TabPicture(2)   =   "frmPuntosGral.frx":0A5E
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frNum(1)"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "INTERIOR"
      TabPicture(3)   =   "frmPuntosGral.frx":0A7A
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frNum(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Begin FPSpread.vaSpread sprGral 
         Height          =   1815
         Index           =   0
         Left            =   225
         OleObjectBlob   =   "frmPuntosGral.frx":0A96
         TabIndex        =   15
         Top             =   1290
         Width           =   8370
      End
      Begin VB.Frame frNum 
         Height          =   3015
         Index           =   3
         Left            =   165
         TabIndex        =   14
         Top             =   390
         Width           =   8565
         Begin VB.TextBox txtPorc 
            Height          =   315
            Index           =   0
            Left            =   3975
            TabIndex        =   35
            Top             =   435
            Width           =   1290
         End
         Begin VB.TextBox txtRep 
            Height          =   315
            Index           =   0
            Left            =   2055
            TabIndex        =   32
            Top             =   435
            Width           =   1725
         End
         Begin VB.CommandButton Command1 
            Height          =   360
            Left            =   7740
            Picture         =   "frmPuntosGral.frx":12D8
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   390
            Width           =   480
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "sobre total de socios"
            Height          =   300
            Index           =   0
            Left            =   5340
            TabIndex        =   39
            Top             =   480
            Width           =   1845
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Repeticiones"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   18
            Top             =   450
            Width           =   1665
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   0
            Left            =   60
            TabIndex        =   17
            Top             =   285
            Width           =   8370
         End
      End
      Begin VB.Frame frNum 
         Height          =   3000
         Index           =   1
         Left            =   -74820
         TabIndex        =   12
         Top             =   405
         Width           =   8550
         Begin FPSpread.vaSpread sprGral 
            Height          =   1815
            Index           =   2
            Left            =   45
            OleObjectBlob   =   "frmPuntosGral.frx":13DA
            TabIndex        =   20
            Top             =   900
            Width           =   8370
         End
         Begin VB.TextBox txtPorc 
            Height          =   315
            Index           =   2
            Left            =   3960
            TabIndex        =   37
            Top             =   435
            Width           =   1290
         End
         Begin VB.TextBox txtRep 
            Height          =   315
            Index           =   2
            Left            =   2010
            TabIndex        =   29
            Top             =   435
            Width           =   1785
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "sobre total de socios"
            Height          =   300
            Index           =   2
            Left            =   5355
            TabIndex        =   41
            Top             =   480
            Width           =   1845
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Repeticiones"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   225
            TabIndex        =   25
            Top             =   435
            Width           =   1650
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   2
            Left            =   45
            TabIndex        =   24
            Top             =   285
            Width           =   8370
         End
      End
      Begin VB.Frame frNum 
         Height          =   2985
         Index           =   2
         Left            =   -74820
         TabIndex        =   13
         Top             =   420
         Width           =   8550
         Begin FPSpread.vaSpread sprGral 
            Height          =   1815
            Index           =   3
            Left            =   45
            OleObjectBlob   =   "frmPuntosGral.frx":1C1C
            TabIndex        =   21
            Top             =   900
            Width           =   8370
         End
         Begin VB.TextBox txtPorc 
            Height          =   315
            Index           =   3
            Left            =   3975
            TabIndex        =   38
            Top             =   420
            Width           =   1290
         End
         Begin VB.TextBox txtRep 
            Height          =   315
            Index           =   3
            Left            =   1980
            TabIndex        =   30
            Top             =   420
            Width           =   1785
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "sobre total de socios"
            Height          =   300
            Index           =   3
            Left            =   5370
            TabIndex        =   42
            Top             =   465
            Width           =   1845
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Repeticiones"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   225
            TabIndex        =   27
            Top             =   420
            Width           =   1650
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   3
            Left            =   45
            TabIndex        =   26
            Top             =   270
            Width           =   8370
         End
      End
      Begin VB.Frame frNum 
         Height          =   2985
         Index           =   0
         Left            =   -74820
         TabIndex        =   11
         Top             =   405
         Width           =   8550
         Begin FPSpread.vaSpread sprGral 
            Height          =   1815
            Index           =   1
            Left            =   60
            OleObjectBlob   =   "frmPuntosGral.frx":245E
            TabIndex        =   19
            Top             =   900
            Width           =   8370
         End
         Begin VB.TextBox txtPorc 
            Height          =   315
            Index           =   1
            Left            =   3975
            TabIndex        =   36
            Top             =   435
            Width           =   1290
         End
         Begin VB.TextBox txtRep 
            Height          =   315
            Index           =   1
            Left            =   2025
            TabIndex        =   28
            Top             =   435
            Width           =   1755
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "sobre total de socios"
            Height          =   300
            Index           =   1
            Left            =   5400
            TabIndex        =   40
            Top             =   480
            Width           =   1845
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Repeticiones"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   23
            Top             =   435
            Width           =   1650
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   1
            Left            =   60
            TabIndex        =   22
            Top             =   285
            Width           =   8370
         End
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
         Picture         =   "frmPuntosGral.frx":2CA0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   570
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   5895
         Picture         =   "frmPuntosGral.frx":3232
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   285
         Width           =   375
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2625
         Picture         =   "frmPuntosGral.frx":33A4
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
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad Total de Socios"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   45
      TabIndex        =   33
      Top             =   5055
      Width           =   2625
   End
End
Attribute VB_Name = "frmPuntosGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset



Private Sub L_CalcularTotales()
Dim sql As String
Dim rs As Recordset


sql = " SELECT COUNT(*) cantidad "
sql = sql & " FROM ventas.cliente_point "
sql = sql & " Where "
sql = sql & " puntos_historicos >= 0 "

If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        txtSocios.Text = rs!cantidad
    rs.MoveNext
    Loop
    Aplicacion.CerrarDAO rs
End If


End Sub

Public Sub L_PonerItems()
Dim i

For i = 0 To 3
sprGral(i).MaxRows = sprGral(i).MaxRows + 1
sprGral(i).SetText 1, sprGral(i).MaxRows, "Facturación"
sprGral(i).MaxRows = sprGral(i).MaxRows + 1
sprGral(i).SetText 1, sprGral(i).MaxRows, "Tickets"
sprGral(i).MaxRows = sprGral(i).MaxRows + 1
sprGral(i).SetText 1, sprGral(i).MaxRows, "Pr. x Tickets"
sprGral(i).MaxRows = sprGral(i).MaxRows + 1
sprGral(i).SetText 1, sprGral(i).MaxRows, "Pr x Ticket"
sprGral(i).MaxRows = sprGral(i).MaxRows + 1
sprGral(i).SetText 1, sprGral(i).MaxRows, "Desc / Importe"
Next


End Sub


Private Sub L_LimpiarGrillas()
Dim i, j, h
For i = 0 To 0
    For j = 2 To 7
        For h = 1 To 4
            sprGral(i).SetText j, h, ""
        Next
    Next
Next

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

frmPuntosGral.caption = Aplicacion.SeteoProceso(frmPuntosGral.caption)

sql = " SELECT cod_depn COD_SDEP"
sql = sql & " ,sum(Imp) ImpTotal ,sum(tick) TickTotal"
sql = sql & " ,sum(impPtos) ImpSocios ,sum(TickPtos) TickSocios"
sql = sql & " ,sum(imp-impPtos) ImpNoSocios ,sum(tick-TickPtos) TickNoSocios"
sql = sql & " ,sum(Descuent) Descuentos"
sql = sql & " From("
sql = sql & " SELECT cod_depn,COD_SDEP,0 imp, 0 tick, SUM(CANT_TICKET) TickPtos,SUM(IMPORTE) ImpPtos,SUM(DESCUENTOS) Descuent"
sql = sql & " From estadis.VENTAS_PUNTOS"
sql = sql & " WHERE FCH_TICKET BETWEEN " & Func.func_ToDate(mskFDesde.FormattedText) & " AND " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " GROUP BY cod_depn,COD_SDEP"
sql = sql & " Union"
sql = sql & " SELECT cod_depn,COD_SDEP,SUM(IMPORTE),SUM(CANT_TICKETS),0,0,0"
sql = sql & " From estadis.PAX_ESPIGON"
sql = sql & " WHERE FCH_VTA BETWEEN " & Func.func_ToDate(mskFDesde.FormattedText) & " AND " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " GROUP BY cod_depn,COD_SDEP"
sql = sql & " Union"
sql = sql & " SELECT cod_depn,COD_SDEP,SUM(IMPORTE),SUM(CANT_TICKETS),0,0,0"
sql = sql & " From ESTADIS.PAX_ESPIGON_HIST"
sql = sql & " WHERE FCH_VTA BETWEEN " & Func.func_ToDate(mskFDesde.FormattedText) & " AND " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " GROUP BY cod_depn,COD_SDEP)"
sql = sql & " GROUP BY cod_depn "

sql = sql & " Union"

sql = sql & " SELECT 'TOT'"
sql = sql & " ,sum(Imp) ImpTotal ,sum(tick) TickTotal"
sql = sql & " ,sum(impPtos) ImpSocios ,sum(TickPtos) TickSocios"
sql = sql & " ,sum(imp-impPtos) ImpNoSocios ,sum(tick-TickPtos) TickNoSocios"
sql = sql & " ,sum(Descuent) Descuentos"
sql = sql & " From("
sql = sql & " SELECT 'TOT',0 imp, 0 tick, SUM(CANT_TICKET) TickPtos,SUM(IMPORTE) ImpPtos,SUM(DESCUENTOS) Descuent"
sql = sql & " From ESTADIS.VENTAS_PUNTOS"
sql = sql & " WHERE FCH_TICKET BETWEEN " & Func.func_ToDate(mskFDesde.FormattedText) & " AND " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " Union"
sql = sql & " SELECT 'TOT',SUM(IMPORTE),SUM(CANT_TICKETS),0,0,0"
sql = sql & " From estadis.PAX_ESPIGON"
sql = sql & " WHERE FCH_VTA BETWEEN  " & Func.func_ToDate(mskFDesde.FormattedText) & " AND " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " Union"
sql = sql & " SELECT 'TOT',SUM(IMPORTE),SUM(CANT_TICKETS),0,0,0"
sql = sql & " From ESTADIS.PAX_ESPIGON_HIST"
sql = sql & " WHERE FCH_VTA BETWEEN " & Func.func_ToDate(mskFDesde.FormattedText) & " AND " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " )"

If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        L_Llenargrilla
        'L_Porcentages sprPorcA, sprEzeA, "COL"
    End If

    Aplicacion.CerrarDAO RsData
    
End If

ErrEA:
    frmPuntosGral.caption = Aplicacion.SeteoFin
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
Dim i As Integer
Dim tit As Variant
Dim NOMBRE As String, depn As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmPuntosGral.caption = Aplicacion.SeteoProceso(frmPuntosGral.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.Application.Visible = True
    
    ReDim titCol(sprGral(0).MaxCols)
    Col = 1
    fila = 4
    
    For i = 1 To sprGral(0).MaxCols
        sprGral(0).GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 2, 1, titulo
    rango = Exl_rangos(2, 2, 1, sprGral(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, sprGral(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    rango = Exl_rangos(fila, fila, 1, sprGral(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, "Cantidad de asociados : " & txtSocios.Text
    rango = Exl_rangos(fila, fila, 1, 1)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        
    For i = 0 To 3
        fila = fila + 2
        Select Case i
            Case 0
                depn = "TOTAL COMPAÑIA"
            Case 1
                depn = "EZEIZA"
            Case 2
                depn = "AEROPARQUE"
            Case 3
                depn = "INTERIOR"
        End Select
        Exl_PonerValor AppExcel, fila, 1, depn
        rango = Exl_rangos(fila, fila, 1, 1)
        Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
        
        If txtRep(i).Text <> "" Then
            fila = fila + 1
            Exl_PonerValor AppExcel, fila, 1, "Clientes con repeticiones = " & txtRep(i).Text & "  - Porc sobre total :" & txtPorc(i).Text
            rango = Exl_rangos(fila, fila, 1, 1)
            Exl_Letra AppExcel, rango, NEGRITA, 10, "Ms Serif"
        
        End If
        
        fila = fila + 1
        Exl_BajarGrillaExel sprGral(i), AppExcel, fila, Col, titCol
        rango = Exl_rangos(fila + 1, fila + sprGral(i).MaxRows, Col + 1, sprGral(i).MaxCols)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila + 3, fila + 4, 3, 3)
        Exl_ColorInt AppExcel, rango, Exl_Gris
        rango = Exl_rangos(fila + 3, fila + 4, 5, 5)
        Exl_ColorInt AppExcel, rango, Exl_Gris
        rango = Exl_rangos(fila + 4, fila + 4, 4, 4)
        Exl_ColorInt AppExcel, rango, Exl_Gris
        rango = Exl_rangos(fila + 4, fila + 4, 6, 6)
        Exl_ColorInt AppExcel, rango, Exl_Gris
        
        fila = fila + sprGral(0).MaxRows + 1
        rango = Exl_rangos(fila, fila, Col, sprGral(0).MaxCols)
        Exl_ColorInt AppExcel, rango, Exl_Gris
    Next
    
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

    frmPuntosGral.caption = Aplicacion.SeteoFin
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
    L_TratarExcel "Informe Estadístico Duty Free Point", "Indicadores generales (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", Total
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



Private Sub Command1_Click()
Dim sql As String
Dim rs As Recordset

frmPuntosGral.caption = Aplicacion.SeteoProceso(frmPuntosGral.caption)

sql = " SELECT cod_depn,COUNT(*) c "
sql = sql & " FROM("
sql = sql & " SELECT cod_depn,NRO_TARJETA,COUNT(*)"
sql = sql & " FROM ventas.TICKET_PUNTOS_MES P, ventas.TICKET_H_MES H"
sql = sql & " Where h.FCH_TICKET = P.FCH_TICKET"
sql = sql & " AND H.COD_CAJA = P.COD_CAJA"
sql = sql & " AND H.NRO_TICKET = P.NRO_TICKET"
sql = sql & " AND NRO_TARJETA <> 0"
sql = sql & " AND H.FCH_TICKET  BETWEEN " & Func.func_ToDate(mskFDesde.FormattedText) & " AND " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " and exists"
sql = sql & " (select * from vtol.clientes_puntos@link_webib_freeshop v where v.nro_tarjeta = p.nro_tarjeta)"
sql = sql & " GROUP BY cod_depn,NRO_TARJETA"
sql = sql & " having count(*)> 1)"
sql = sql & " GROUP BY cod_depn"
sql = sql & " Union"
sql = sql & " SELECT cod_depn,COUNT(*)"
sql = sql & " FROM("
sql = sql & " SELECT 'TOT' cod_depn,NRO_TARJETA,COUNT(*)"
sql = sql & " FROM ventas.TICKET_PUNTOS_MES P, ventas.TICKET_H_MES H"
sql = sql & " Where h.FCH_TICKET = P.FCH_TICKET"
sql = sql & " AND H.COD_CAJA = P.COD_CAJA"
sql = sql & " AND H.NRO_TICKET = P.NRO_TICKET"
sql = sql & " AND NRO_TARJETA <> 0"
sql = sql & " AND H.FCH_TICKET  BETWEEN " & Func.func_ToDate(mskFDesde.FormattedText) & " AND " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " and exists"
sql = sql & " (select * from vtol.clientes_puntos@link_webib_freeshop v where v.nro_tarjeta = p.nro_tarjeta)"
sql = sql & " GROUP BY NRO_TARJETA"
sql = sql & " having count(*)> 1)"
sql = sql & " GROUP BY cod_depn"

If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
    Select Case rs!cod_depn
           Case "TOT"
                txtRep(0).Text = rs!c
                txtPorc(0).Text = Format(txtRep(0).Text / txtSocios.Text * 100, "###.00") & " %"
           Case "EZE"
                txtRep(1).Text = rs!c
                txtPorc(1).Text = Format(txtRep(1).Text / txtSocios.Text * 100, "###.00") & " %"
           Case "AEP"
                txtRep(2).Text = rs!c
                txtPorc(2).Text = Format(txtRep(2).Text / txtSocios.Text * 100, "###.00") & " %"
           Case "INT"
                txtRep(3).Text = rs!c
                txtPorc(3).Text = Format(txtRep(3).Text / txtSocios.Text * 100, "###.00") & " %"
    End Select
    
    rs.MoveNext
    Loop
    Aplicacion.CerrarDAO rs
End If

    frmPuntosGral.caption = Aplicacion.SeteoFin

End Sub

Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabEspigon


End Sub

Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6200
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
Dim prtSocio As Single, prtNoSocio As Single


Do While Not RsData.EOF
    Select Case RsData!Cod_Sdep
           Case "TOT"
                i = 0
           Case "EZE"
                i = 1
           Case "INT"
                i = 3
           Case "AEP"
                i = 2
    End Select
        'Datos del socio
        sprGral(i).SetText 2, 1, Format(RsData!impsocios, "###,###,###.00")
        sprGral(i).SetText 2, 2, str(RsData!ticksocios)
        If RsData!ticksocios > 0 Then
            sprGral(i).SetText 2, 3, Format(RsData!impsocios / RsData!ticksocios, "###,###.00")
            prtSocio = RsData!impsocios / RsData!ticksocios
        End If
        
        sprGral(i).SetText 2, 4, Format(RsData!descuentos / RsData!impsocios * 100, "###,###.00")
        
        sprGral(i).SetText 3, 1, Format(RsData!impsocios / RsData!imptotal * 100, "##.00")
        sprGral(i).SetText 3, 2, Format(RsData!ticksocios / RsData!ticktotal * 100, "##.00")
        
        'Datos  no socio
        sprGral(i).SetText 4, 1, Format(RsData!impnosocios, "###,###,###.00")
        sprGral(i).SetText 4, 2, str(RsData!ticknosocios)
       
        If RsData!ticknosocios > 0 Then
            sprGral(i).SetText 4, 3, Format(RsData!impnosocios / RsData!ticknosocios, "###,###.00")
            prtNoSocio = RsData!impnosocios / RsData!ticknosocios
        End If
        
        sprGral(i).SetText 5, 1, Format(RsData!impnosocios / RsData!imptotal * 100, "##.00")
        sprGral(i).SetText 5, 2, Format(RsData!ticknosocios / RsData!ticktotal * 100, "##.00")


        If RsData!impnosocios > 0 Then
            sprGral(i).SetText 6, 1, Format(100 * ((RsData!impsocios / RsData!impnosocios) - 1), "###,###.00")
            sprGral(i).SetText 6, 2, Format(100 * ((RsData!ticksocios / RsData!ticknosocios) - 1), "###,###.00")
        End If
        If prtNoSocio > 0 Then
            sprGral(i).SetText 6, 3, Format(100 * ((prtSocio / prtNoSocio) - 1), "###,###.00")
        End If

        
        'Datos totales
        sprGral(i).SetText 7, 1, Format(RsData!imptotal, "###,###,###.00")
        sprGral(i).SetText 7, 2, str(RsData!ticktotal)
        If RsData!ticktotal > 0 Then
            sprGral(i).SetText 7, 3, Format(RsData!imptotal / RsData!ticktotal, "###,###.00")
        End If
        
        sprGral(i).SetText 7, 4, Format(RsData!descuentos / RsData!imptotal * 100, "###,###.00")
        'sprGral(i).SetText 4, sprGral(i).MaxRows, str(RsData!descuentos)
    

        
    RsData.MoveNext
Loop
L_CalcularTotales
txtRep(0).Refresh
End Sub

