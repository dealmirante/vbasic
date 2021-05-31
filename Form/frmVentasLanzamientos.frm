VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVentasLanzamientos 
   Caption         =   "Puntos por Nacionalidad"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   9930
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
      Picture         =   "frmVentasLanzamientos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Picture         =   "frmVentasLanzamientos.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Picture         =   "frmVentasLanzamientos.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   2
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
      TabPicture(0)   =   "frmVentasLanzamientos.frx":0A26
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frNum(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "sprLanz"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "EZEIZA"
      TabPicture(1)   =   "frmVentasLanzamientos.frx":0A42
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frNum(0)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "AEROPARQUE"
      TabPicture(2)   =   "frmVentasLanzamientos.frx":0A5E
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frNum(1)"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "INTERIOR"
      TabPicture(3)   =   "frmVentasLanzamientos.frx":0A7A
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frNum(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Begin FPSpread.vaSpread sprLanz 
         Height          =   2775
         Left            =   225
         OleObjectBlob   =   "frmVentasLanzamientos.frx":0A96
         TabIndex        =   9
         Top             =   570
         Width           =   8670
      End
      Begin VB.Frame frNum 
         Height          =   3105
         Index           =   3
         Left            =   135
         TabIndex        =   8
         Top             =   390
         Width           =   8865
      End
      Begin VB.Frame frNum 
         Height          =   3150
         Index           =   1
         Left            =   -74900
         TabIndex        =   6
         Top             =   405
         Width           =   8865
      End
      Begin VB.Frame frNum 
         Height          =   3150
         Index           =   2
         Left            =   -74900
         TabIndex        =   7
         Top             =   375
         Width           =   8880
      End
      Begin VB.Frame frNum 
         Height          =   3150
         Index           =   0
         Left            =   -74895
         TabIndex        =   5
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
      Begin VB.ComboBox cboAnio 
         Height          =   315
         ItemData        =   "frmVentasLanzamientos.frx":0EF7
         Left            =   4140
         List            =   "frmVentasLanzamientos.frx":0F0D
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   375
         Width           =   900
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmVentasLanzamientos.frx":0F35
         Left            =   1590
         List            =   "frmVentasLanzamientos.frx":0F5D
         TabIndex        =   11
         Text            =   "cboMes"
         Top             =   360
         Width           =   930
      End
      Begin VB.CommandButton BotExcelResumen 
         Caption         =   "Excel"
         Height          =   480
         Left            =   7755
         Picture         =   "frmVentasLanzamientos.frx":0F91
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   570
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
         Left            =   2790
         TabIndex        =   14
         Top             =   360
         Width           =   1185
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
         Left            =   315
         TabIndex        =   13
         Top             =   375
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmVentasLanzamientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset



Private Sub L_LimpiarGrillas()
Dim i

sprLanz.MaxRows = 0

End Sub

Private Function L_ObtenerDatosProv(pProv As String, pFDesde As String, pFHasta As String, pTipo As String, ByRef NOMBRE As String) As Single
Dim sql As String
Dim rs As Recordset

sql = " select descrip,estadis.venta_local_marca_imp(null,'" & pProv & "',null,"
sql = sql & func_ToDate(pFDesde) & "," & func_ToDate(pFHasta) & ",'" & pTipo & "') imp "
sql = sql & " From Proveedor  "
sql = sql & " Where cod_prov = '" & pProv & "' "

If Aplicacion.ObtenerRsDAO(sql, rs) Then
    If Aplicacion.CantReg(rs) > 0 Then
        NOMBRE = rs!Descrip
        L_ObtenerDatosProv = rs!imp
    Else
        L_ObtenerDatosProv = 0
    End If
    rs.Close
End If
End Function

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

frmVentasLanzamientos.caption = Aplicacion.SeteoProceso(frmVentasLanzamientos.caption)

sql = " select v.cod_prov,v.cod_prod,p.descrip,sum(cantidad) cant,sum(importe) imp "
sql = sql & " ,decode(to_char(fch_ticket,'YYYY'),to_char(sysdate,'YYYY'),'A','H') Tabla "
sql = sql & " ,to_char(fch_desde,'dd/mm/yyyy') fch_desde, to_char(fch_hasta,'dd/mm/yyyy') fch_hasta "
sql = sql & " from baires.producto p, estadis.venta_plg v, estadis.codigos_lanzamientos c "
sql = sql & " Where v.Cod_prod = c.Cod_prod "
sql = sql & " and p.cod_prod = v.cod_prod"
sql = sql & " and fch_ticket between fch_desde and fch_hasta"
sql = sql & " and aniomes = " & cboAnio.Text & cboMes.Text
sql = sql & " group by v.cod_prov,v.cod_prod,p.descrip "
sql = sql & " ,decode(to_char(fch_ticket,'YYYY'),to_char(sysdate,'YYYY'),'A','H') "
sql = sql & " ,fch_desde, fch_hasta "

If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        L_Llenargrilla
    End If

    Aplicacion.CerrarDAO RsData
    
End If

ErrEA:
    frmVentasLanzamientos.caption = Aplicacion.SeteoFin
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




Private Sub L_TratarExcel(titulo As String, subTit As String)
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

frmVentasLanzamientos.caption = Aplicacion.SeteoProceso(frmVentasLanzamientos.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.Application.Visible = True
    
    ReDim titCol(sprLanz.MaxCols)
    Col = 1
    fila = 4
    
    For i = 1 To sprLanz.MaxCols
        sprLanz.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 2, 1, titulo
    rango = Exl_rangos(2, 2, 1, sprLanz.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, sprLanz.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    'Exl_PonerValor AppExcel, fila, 1, ""
    rango = Exl_rangos(fila, fila, 1, 1)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    fila = fila + 1
    Exl_BajarGrillaExel sprLanz, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprLanz.MaxRows, Col + 1, sprLanz.MaxCols)
    Exl_Format AppExcel, rango
    
    Exl.Exl_AnchoCol AppExcel, 1, 1, 8
    Exl.Exl_AnchoCol AppExcel, 2, 2, 20
    Exl.Exl_AnchoCol AppExcel, 5, 5, 20
    Exl.Exl_AnchoCol AppExcel, 6, 6, 8
    
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

    frmVentasLanzamientos.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub BotExcelResumen_Click()
    L_TratarExcel " LANZAMIENTOS ", " Datos total compañía " & cboAnio.Text & "-" & cboMes.Text
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

cboAnio.ListIndex = Year(Now) - 2002
cboMes.ListIndex = Month(Now) - 1

L_LimpiarGrillas

tabEspigon.TabVisible(1) = False
tabEspigon.TabVisible(2) = False
tabEspigon.TabVisible(3) = False

End Sub




Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmGLC"
End Sub


Private Sub L_Llenargrilla()
Dim fecha As String
Dim i As Integer
Dim prov As Single
Dim provDescrip As String

Do While Not RsData.EOF
    sprLanz.MaxRows = sprLanz.MaxRows + 1
    
    sprLanz.SetText 1, sprLanz.MaxRows, Trim(RsData!Cod_prod)
    sprLanz.SetText 2, sprLanz.MaxRows, Trim(RsData!Descrip)
    sprLanz.SetText 3, sprLanz.MaxRows, str(RsData!cant)
    sprLanz.SetText 4, sprLanz.MaxRows, str(RsData!imp)
           
    prov = L_ObtenerDatosProv(RsData!cod_prov, RsData!fch_desde, RsData!fch_hasta, RsData!Tabla, provDescrip)
    sprLanz.SetText 5, sprLanz.MaxRows, Trim(provDescrip)
    If prov <> 0 Then
    sprLanz.SetText 6, sprLanz.MaxRows, str(RsData!imp / prov * 100)
    End If
    RsData.MoveNext
Loop

'Spread_TotalesGrillas sprLanz, sprLanz.MaxCols - 3, 2

'L_Participacion sprLanz

End Sub

