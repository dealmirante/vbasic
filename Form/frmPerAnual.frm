VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form frmPerAnual 
   Caption         =   "Comparación con períodos anteriores"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   7485
   Begin FPSpread.vaSpread sprPer 
      Height          =   780
      Left            =   120
      OleObjectBlob   =   "frmPerAnual.frx":0000
      TabIndex        =   13
      Top             =   4770
      Width           =   6555
   End
   Begin GraphLib.Graph GR 
      Height          =   3150
      Left            =   105
      TabIndex        =   15
      Top             =   1530
      Width           =   6570
      _Version        =   65536
      _ExtentX        =   11589
      _ExtentY        =   5556
      _StockProps     =   96
      BorderStyle     =   1
      Background      =   8
      GraphType       =   4
      GridStyle       =   1
      IndexStyle      =   1
      LegendStyle     =   1
      NumPoints       =   10
      Palette         =   1
      PatternedLines  =   1
      PrintStyle      =   3
      RandomData      =   0
      ThickLines      =   0
      ColorData       =   10
      ColorData[0]    =   12
      ColorData[1]    =   9
      ColorData[2]    =   10
      ColorData[3]    =   14
      ColorData[4]    =   15
      ColorData[5]    =   11
      ColorData[6]    =   13
      ColorData[7]    =   7
      ColorData[8]    =   3
      ColorData[9]    =   2
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   200
      FontSize[1]     =   150
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      GraphData       =   1
      GraphData[]     =   10
      GraphData[0,0]  =   0
      GraphData[0,1]  =   0
      GraphData[0,2]  =   0
      GraphData[0,3]  =   0
      GraphData[0,4]  =   0
      GraphData[0,5]  =   0
      GraphData[0,6]  =   0
      GraphData[0,7]  =   0
      GraphData[0,8]  =   0
      GraphData[0,9]  =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
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
      Left            =   6825
      Picture         =   "frmPerAnual.frx":0207
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1110
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
      Height          =   465
      Index           =   1
      Left            =   6810
      Picture         =   "frmPerAnual.frx":0A29
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   615
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
      Height          =   465
      Index           =   0
      Left            =   6825
      Picture         =   "frmPerAnual.frx":0B2B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   615
      Width           =   615
   End
   Begin VB.Frame FrCab 
      Height          =   1515
      Left            =   90
      TabIndex        =   0
      Top             =   -45
      Width           =   6615
      Begin VB.ComboBox cboGrupo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4860
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1080
         Width           =   1425
      End
      Begin VB.ComboBox cboLocal 
         Height          =   315
         Left            =   4860
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   645
         Width           =   1425
      End
      Begin VB.ListBox LstEspigon 
         Height          =   255
         Left            =   6360
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox CboCodAeropuerto 
         Height          =   315
         ItemData        =   "frmPerAnual.frx":0C2D
         Left            =   1245
         List            =   "frmPerAnual.frx":0C2F
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   645
         Width           =   2070
      End
      Begin VB.ComboBox CboEspigon 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   2070
      End
      Begin MSMask.MaskEdBox mskPer 
         Height          =   285
         Left            =   4860
         TabIndex        =   5
         Top             =   285
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2640
         Picture         =   "frmPerAnual.frx":0C31
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   285
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1485
         TabIndex        =   2
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.Label LblEspigon 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3540
         TabIndex        =   18
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label LblEspigon 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3540
         TabIndex        =   16
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label LblCodAeropuerto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aeropuerto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   645
         Width           =   1080
      End
      Begin VB.Label LblEspigon 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Espigón :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cant. Períodos"
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
         TabIndex        =   4
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Tope"
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
         Left            =   165
         TabIndex        =   3
         Top             =   300
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmPerAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As Recordset
Dim TipoCons As String

Private Sub L_SetearCol(Col As Integer)

    'Select a block of cells

    'Select column(s)
    sprPer.Col = Col
    sprPer.Row = -1
    sprPer.Col2 = Col
    sprPer.Row2 = -1
    sprPer.BlockMode = True

    'Define cells as type FLOAT
    sprPer.CellType = SS_CELL_TYPE_FLOAT
    sprPer.TypeHAlign = SS_CELL_H_ALIGN_RIGHT
    sprPer.TypeFloatMin = "-9999999999.00"
    sprPer.TypeFloatMax = "9999999999.00"
    sprPer.TypeFloatMoney = False
    sprPer.TypeFloatSeparator = True
    sprPer.TypeFloatDecimalPlaces = 0
'    sprPer.TypeFloatCurrencyChar = Asc("$")
    sprPer.TypeFloatSepChar = Asc(",")
    sprPer.TypeFloatDecimalChar = Asc(".")

    'Set the width of a selected column
    sprPer.ColWidth(1) = 10#

End Sub


Private Function L_ArmarAgrupamiento() As String
Dim Cond

Cond = ""

If CboCodAeropuerto.Text <> "" Then
    Cond = Cond & " ,cod_depn "
End If

If CboEspigon.Text <> "" Then
    Cond = Cond & " ,cod_sdep "
End If

If cboLocal.Text <> "" Then
    Cond = Cond & " ,cod_local "
End If

If cboGrupo.Text <> "" Then
    Cond = Cond & " ,grupo_venta "
End If

If Cond <> "" Then
    Cond = " Group By " & Right(Cond, Len(Cond) - 2)
End If

L_ArmarAgrupamiento = Cond

End Function

Private Function L_ArmarSelect() As String
Dim Cond

Cond = ""

TipoCons = "Cia"

If CboCodAeropuerto.Text <> "" Then
    Cond = Cond & " ,cod_depn "
    TipoCons = "Dep"
End If

If CboEspigon.Text <> "" Then
    Cond = Cond & " ,cod_sdep "
    TipoCons = "SDep"
End If

If Cond <> "" Then
    Cond = Cond
End If

L_ArmarSelect = Cond
End Function


Private Sub L_Llenargrilla(anio As Integer)

sprPer.MaxCols = sprPer.MaxCols + 1
sprPer.SetText sprPer.MaxCols, 0, str(anio)

sprPer.SetText sprPer.MaxCols, 1, str(rs!imp)
End Sub

Private Sub L_NombrePrimerCol()

Select Case TipoCons
    Case "Cia"
        sprPer.SetText 1, 0, "Compañía"
    Case "Dep"
        sprPer.SetText 1, 0, "Aeropuerto"
    Case "SDep"
        sprPer.SetText 1, 0, "Espigón"
        
End Select

End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim anio As Integer
Dim i As Integer

On Error GoTo ErrPer:

frmPerAnual.caption = Aplicacion.SeteoProceso(frmPerAnual.caption)

anio = Year(mskFDesde.FormattedText)

GR.NumPoints = mskPer.Text
GR.AutoInc = 0
'GR.YAxisMax = (L_MaximoValor(spr_local.MaxRows - 1, Cant, spr_local) / 100) * 100
GR.YAxisTicks = 10
GR.NumSets = 1

sprPer.MaxRows = 1
For i = 1 To mskPer.Text
    sql = " SELECT "
    sql = sql & " decode(Sum (cant_tickets),null,0,Sum (cant_tickets)) cant_t, "
    sql = sql & " decode (sum(importe),null,0,sum(importe)) imp, "
    sql = sql & " decode(sum(cant_pax),null,0,sum(cant_pax)) cant_p "
'    sql = sql & L_ArmarSelect
    If cboLocal.Text = "" And cboGrupo.Text = "" Then
        sql = sql & "FROM " & funcLocal_Vista("pax_espigon", anio)
    Else
        sql = sql & "FROM " & funcLocal_Vista("pax_grupo", anio)
    End If
    sql = sql & L_Armarcondicion(anio)
    sql = sql & L_ArmarAgrupamiento
    
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    
    If Aplicacion.CantReg(rs) > 0 Then
        sprPer.MaxCols = sprPer.MaxCols + 1
        
        L_SetearCol i
        
        sprPer.SetText sprPer.MaxCols, 0, str(anio)
        sprPer.SetText sprPer.MaxCols, 1, str(rs!imp)
        L_TratarGrafico i, anio
        
        FrCab.Enabled = False
        botEjecutar(0).Enabled = False
    Else
        sprPer.MaxCols = sprPer.MaxCols + 1
        sprPer.SetText sprPer.MaxCols, 0, str(anio)
        sprPer.SetText sprPer.MaxCols, 1, "-"
    End If
    
    Aplicacion.CerrarDAO rs
    
End If
anio = anio - 1
Next
ErrPer:
    frmPerAnual.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub
Public Sub L_TratarGrafico(Col As Integer, anio As Integer)
Dim i, J
Dim cant As Integer, Puntos As Integer
Dim valor As Variant

'Graph1.Visible = False

On Error GoTo ErrGr:

GR.ThisPoint = Col

GR.LegendText = anio
    
GR.LabelText = " "
 
GR.GraphData = rs!imp

GR.DrawMode = 2
ErrGr:
    Exit Sub
    
End Sub


Private Function L_Armarcondicion(anio As Integer)
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String

fechaDesde = "01-01-" & Trim(str(anio))

fechaHasta = Format$(mskFDesde.FormattedText, FTOFECHA)

Cond = " WHERE fch_vta between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

If CboCodAeropuerto.Text <> "" Then
    Cond = Cond & " and cod_depn = '" & CboCodAeropuerto.Text & "'"
End If

If CboEspigon.Text <> "" Then
    Cond = Cond & " and cod_sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
End If

If cboGrupo.Text <> "" Then
    Cond = Cond & " and grupo_venta = '" & cboGrupo.Text & "'"
End If

If cboLocal.Text <> "" Then
    Cond = Cond & " and cod_local = '" & cboLocal.Text & "'"
End If

'If cboRubro.Text <> "" Then
'    Cond = Cond & " and cod_rubr = '" & cboRubro.Text & "'"
'End If

'If optComi(0).Value Then
'    Cond = Cond & " and Comitente NOT IN ('IBAIR','IOSC') "
'ElseIf optComi(1).Value Then
'    Cond = Cond & " and Comitente  IN ('IBAIR','IOSC') "
'End If

L_Armarcondicion = Cond

End Function

Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        L_Refrescar
    Case 1
        FrCab.Enabled = True
        botEjecutar(0).Enabled = True
'        tabEst.Enabled = False
        L_LimpiarGrillas
    Case 2
        Unload Me
End Select
End Sub

Private Sub L_LimpiarGrillas()
Dim i

 sprPer.MaxRows = 0
 sprPer.MaxCols = 0
 
 GR.DrawMode = 1
  
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

Private Sub CboCodAeropuerto_Click()
Dim sql As String

sql = " SELECT cod_sdep,descrip FROM baires.subdependencia "
sql = sql & " WHERE cod_depn = '" & CboCodAeropuerto.Text & "'"
sql = sql & " ORDER BY cod_sdep"
 
FuncCbos_LlenarCboLst CboEspigon, LstEspigon, sql

End Sub


Private Sub CboCodAeropuerto_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    CboCodAeropuerto.ListIndex = -1
    CboEspigon.Clear
    cboLocal.Clear
    cboGrupo.ListIndex = -1
    cboGrupo.Enabled = False
End If
End Sub


Private Sub CboEspigon_Click()
Dim sql As String

sql = " SELECT cod_loc FROM baires.local "
sql = sql & " WHERE cod_depn = '" & CboCodAeropuerto.Text & "'"
sql = sql & " AND cod_sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
sql = sql & " ORDER BY cod_loc"
 
FuncCbos_LlenarCbo cboLocal, sql

cboGrupo.Enabled = True
End Sub

Private Sub CboEspigon_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    CboEspigon.ListIndex = -1
    cboLocal.Clear
    cboGrupo.ListIndex = -1
    cboGrupo.Enabled = False
End If

End Sub


Private Sub cboGrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    cboGrupo.ListIndex = -1
End If
End Sub


Private Sub cboLocal_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    cboLocal.ListIndex = -1
End If
End Sub


Private Sub Form_Load()
Dim sql As String

Me.Left = 50
Me.Top = 100
Me.Height = 6000
Me.Width = 8100

mskFDesde.Text = Format$(Date, FTOFECHA)

mskPer.Text = 2

sql = " SELECT cod_depn,descrip FROM baires.dependencia "
sql = sql & " ORDER BY cod_depn"

FuncCbos_LlenarCbo CboCodAeropuerto, sql

GR.DrawMode = 1

cboGrupo.AddItem "A"
cboGrupo.AddItem "B"
cboGrupo.AddItem "C"

frmPrincipal.lstForms.AddItem "frmPerAnual"
End Sub



Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmPerAnual"
End Sub


Private Sub mskFDesde_LostFocus()

If Not IsDate(mskFDesde.FormattedText) Then
    mskFDesde.Text = Date
End If
mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
    
End Sub


Private Sub mskPer_LostFocus()
If mskPer.Text = "" Then
    mskPer.Text = 2
ElseIf Val(mskPer.Text) < 2 Then
    mskPer.Text = 2
End If

End Sub


