VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmVersus 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "Comparativo de pasajeros volados vs atendidos"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar 
      Height          =   390
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "N"
            Object.ToolTipText     =   "Nueva Selección"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "B"
            Object.ToolTipText     =   "Buscar"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "M"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "P"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "S"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Height          =   525
      Left            =   6765
      Picture         =   "FrmVersus.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4680
      Width           =   540
   End
   Begin VB.Frame FramCabecera 
      Height          =   1785
      Left            =   120
      TabIndex        =   7
      Top             =   315
      Width           =   7395
      Begin VB.Frame Frame1 
         Height          =   660
         Left            =   3855
         TabIndex        =   16
         Top             =   1035
         Width           =   3465
         Begin VB.ComboBox cboCia 
            Height          =   315
            Left            =   855
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   225
            Width           =   2355
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CIA.:"
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
            Left            =   135
            TabIndex        =   17
            Top             =   225
            Width           =   690
         End
      End
      Begin VB.Frame FramAerop 
         Height          =   1515
         Left            =   90
         TabIndex        =   18
         Top             =   180
         Width           =   5475
         Begin VB.CommandButton Command2 
            Height          =   390
            Left            =   4965
            Picture         =   "FrmVersus.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   180
            Width           =   465
         End
         Begin VB.ComboBox CboEspigon 
            Height          =   315
            Left            =   1455
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1035
            Width           =   2130
         End
         Begin VB.ComboBox CboCodAeropuerto 
            Height          =   315
            ItemData        =   "FrmVersus.frx":05CC
            Left            =   1470
            List            =   "FrmVersus.frx":05CE
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   630
            Width           =   2130
         End
         Begin VB.CommandButton Command1 
            Height          =   390
            Left            =   2250
            Picture         =   "FrmVersus.frx":05D0
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   180
            Width           =   465
         End
         Begin VB.ListBox LstEspigon 
            Height          =   255
            Left            =   3285
            TabIndex        =   19
            Top             =   1035
            Visible         =   0   'False
            Width           =   135
         End
         Begin MSMask.MaskEdBox MskFecha 
            Height          =   315
            Left            =   1110
            TabIndex        =   23
            Top             =   210
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   327680
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd-mm-yyyy"
            Mask            =   "##-##-####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFechaH 
            Height          =   315
            Left            =   3840
            TabIndex        =   28
            Top             =   210
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   327680
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd-mm-yyyy"
            Mask            =   "##-##-####"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha.H :"
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
            Left            =   2910
            TabIndex        =   27
            Top             =   225
            Width           =   930
         End
         Begin VB.Label Label2 
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
            Left            =   150
            TabIndex        =   26
            Top             =   1050
            Width           =   1245
         End
         Begin VB.Label LblCodAeropuerto 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cod. Aerop.:"
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
            Left            =   135
            TabIndex        =   25
            Top             =   630
            Width           =   1260
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha.D :"
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
            Left            =   135
            TabIndex        =   24
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Frame FraTipoVuelo 
         Caption         =   "Tipo de vuelo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   930
         Left            =   5625
         TabIndex        =   8
         Top             =   135
         Width           =   1695
         Begin VB.OptionButton OptTipoVuelo 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   375
            TabIndex        =   2
            Top             =   660
            Width           =   855
         End
         Begin VB.OptionButton OptTipoVuelo 
            Caption         =   "Llegada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   375
            TabIndex        =   0
            Top             =   195
            Width           =   1005
         End
         Begin VB.OptionButton OptTipoVuelo 
            Caption         =   "Salida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   375
            TabIndex        =   1
            Top             =   405
            Width           =   855
         End
      End
   End
   Begin VB.Frame FramGrilla 
      Height          =   3555
      Left            =   90
      TabIndex        =   5
      Top             =   2040
      Width           =   7455
      Begin FPSpread.vaSpread Grilla 
         Height          =   2175
         Left            =   90
         OleObjectBlob   =   "FrmVersus.frx":075A
         TabIndex        =   4
         Top             =   165
         Width           =   7230
      End
      Begin VB.Frame FraDescripcion 
         Height          =   1005
         Left            =   90
         TabIndex        =   9
         Top             =   2340
         Width           =   7230
         Begin VB.TextBox txtFecha 
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   570
            Width           =   1230
         End
         Begin VB.TextBox TxtTipoVuelo 
            Enabled         =   0   'False
            Height          =   330
            Left            =   4905
            TabIndex        =   15
            Top             =   150
            Width           =   1185
         End
         Begin VB.TextBox TxtEspigon 
            Enabled         =   0   'False
            Height          =   345
            Left            =   1530
            TabIndex        =   14
            Top             =   570
            Width           =   1530
         End
         Begin VB.TextBox TxtAeropuerto 
            Enabled         =   0   'False
            Height          =   345
            Left            =   1530
            TabIndex        =   13
            Top             =   165
            Width           =   1140
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha de carga de Pax mas retrasada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   405
            Left            =   3435
            TabIndex        =   30
            Top             =   525
            Width           =   1680
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Aeropuerto :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   105
            TabIndex        =   12
            Top             =   165
            Width           =   1380
         End
         Begin VB.Label LblEspigon 
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
            Height          =   330
            Left            =   90
            TabIndex        =   11
            Top             =   585
            Width           =   1380
         End
         Begin VB.Label LblTipoVuelo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipo de vuelo : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3465
            TabIndex        =   10
            Top             =   150
            Width           =   1395
         End
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmVersus.frx":0B64
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmVersus.frx":0E7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmVersus.frx":1198
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmVersus.frx":14B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmVersus.frx":17CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmVersus.frx":18DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmVersus.frx":19F0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmVersus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset

Dim cl_versus As CLVersus

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$

Dim CantReg As Integer


Private Sub MeSetearBotonesToolBar()
Dim i%
Dim but As Button

If CantReg = 0 Then
    Toolbar.Buttons(1).Enabled = False
    Toolbar.Buttons(2).Enabled = False
    
    Toolbar.Buttons(4).Enabled = False
    
    Toolbar.Buttons(6).Enabled = False
    Toolbar.Buttons(7).Enabled = True
    
Else
    Toolbar.Buttons(1).Enabled = True
    Toolbar.Buttons(2).Enabled = False
    
    Toolbar.Buttons(4).Enabled = True
    
    Toolbar.Buttons(6).Enabled = True
    Toolbar.Buttons(7).Enabled = True
End If
  
End Sub

Public Sub LimpiarGrilla()
Dim i As Integer
i = 1
'While i < Grilla.MaxRows
'  Grilla.Row = i
'  Grilla.col = 1
'  Grilla.Text = ""
'  Grilla.col = 2
'  Grilla.Text = ""
'  Grilla.col = 3
'  Grilla.Text = ""
'  Grilla.col = 4
'  Grilla.Text = ""
'  Grilla.col = 5
'  Grilla.Text = ""
'  Grilla.col = 6
'  Grilla.Text = ""
'  Grilla.col = 7
'  Grilla.Text = ""
'  Grilla.col = 8
'  Grilla.Text = ""
'  i = i + 1
'Wend
TxtAeropuerto.Text = ""
TxtEspigon.Text = ""
TxtTipoVuelo.Text = ""
Grilla.MaxRows = 0
End Sub



Public Sub NuevaSeleccion()
Dim i%

FramCabecera.Enabled = True

SetBotonesGeneral True

'Limpiar campos de pantallas

'CboCodAeropuerto.ListIndex = -1
OptTipoVuelo.item(2).Visible = True
OptTipoVuelo(2) = True
LimpiarGrilla
Set cl_versus = New CLVersus

MskFecha.SetFocus

End Sub

Private Sub SeteoBotonesMod(valor As Boolean)
    
    
    Toolbar.Buttons(1).Enabled = valor
    Toolbar.Buttons(2).Enabled = Not valor
    
    Toolbar.Buttons(4).Enabled = valor
    
    Toolbar.Buttons(6).Enabled = valor
    Toolbar.Buttons(7).Enabled = valor
    
'habilitar o des frames y/o campos
FramCabecera.Enabled = Not valor


End Sub
Private Sub SetBotonesGeneral(valor As Boolean)
    
    Toolbar.Buttons(1).Enabled = Not valor
    Toolbar.Buttons(2).Enabled = valor
    
    Toolbar.Buttons(4).Enabled = Not valor
    
    Toolbar.Buttons(6).Enabled = Not valor
    Toolbar.Buttons(7).Enabled = valor
    
'habilitar frames

FramAerop.Enabled = valor
OptTipoVuelo(2).Enabled = valor

End Sub
Public Sub MeLlenargrilla()
Dim i As Integer
Dim TipoVuelo As Variant
Dim Aeropuerto As Variant
Dim Espigon As Variant
Dim valorAt As Variant
Dim valorTr As Variant
Dim desc As String

On Error GoTo E_LlenarGrilla:

i = 0
Grilla.MaxRows = 0

While Not rs.EOF()

Grilla.MaxRows = Grilla.MaxRows + 1
i = i + 1
Grilla.Row = i
If OptTipoVuelo(2) = True Then
   FraDescripcion.Visible = True
 Else
End If

Grilla.Row = i
Grilla.col = 1
Grilla.Text = rs!Descrip
Grilla.col = 2
Grilla.Text = rs!cod_vuelo
Grilla.col = 3
Grilla.Text = rs!pax_vol
Grilla.col = 4
Grilla.Text = rs!pax_at
Grilla.col = 5

If rs!pax_vol = 0 Or rs!pax_at = 0 Then
  Grilla.Text = 0
  Else
   Grilla.Text = (rs!pax_at * 100) / rs!pax_vol
   Grilla.Text = Grilla.Text & "%"
End If

Grilla.col = 6
Grilla.Text = rs!Tipo
Grilla.col = 7
Grilla.Text = rs!cod_depn
Grilla.col = 8
Grilla.Text = rs!cod_sdep
Grilla.col = 9
Grilla.Text = rs!cod_cia_aerea

rs.MoveNext
CantReg = Grilla.MaxRows
Wend
Spread.Spread_TotalesGrillas Grilla, 3, 3

Grilla.GetText 3, Grilla.MaxRows, valorTr
Grilla.GetText 4, Grilla.MaxRows, valorAt

Grilla.SetText 5, Grilla.MaxRows, Val(valorAt) / Val(valorTr) * 100

'refrescar primera fila de la grilla

Grilla.GetText 6, 1, TipoVuelo

Select Case TipoVuelo
  Case Is = "S"
     TxtTipoVuelo.Text = "Salida"
  Case Is = "L"
     TxtTipoVuelo.Text = "LLegada"
  Case Is = "I"
     TxtTipoVuelo.Text = "Incierto"
End Select

Grilla.GetText 7, 1, Aeropuerto

TxtAeropuerto.Text = Aeropuerto

Grilla.GetText 8, 1, Espigon

TxtEspigon.Text = Espigon

'Deshabilito frame cabecera
FramCabecera.Enabled = False
FraDescripcion.Enabled = False


Exit Sub

E_LlenarGrilla:
MsgBox "Se ha producido un error inesperado (LLENANDO GRILLA)", vbOKOnly + vbCritical, "Error"

End Sub

Private Sub MeCargarGrilla()

Aplicacion.SeteoProceso (FrmVersus.caption)
        
    CondConsulta = ArmarCondicion

    sqlGral = ""
    sqlGral = " SELECT cod_vuelo,cod_depn,cod_sdep,tipo,pax_at,pax_vol,descrip,cod_cia_aerea"
    sqlGral = sqlGral & " FROM estadis.pax_at_vs_vol, ventas.companias "
    sqlGral = sqlGral & "WHERE cod_cia_aerea = cod_compania"
    sqlGral = sqlGral & " AND (pax_at <> 0 "
    sqlGral = sqlGral & " OR pax_vol <> 0) "
    
    sqlGral = sqlGral & CondConsulta
    
If Aplicacion.ObtenerRsDAO(sqlGral, rs) Then
 If Aplicacion.CantReg(rs) > 0 Then
     MeLlenargrilla
     MeSetearBotonesToolBar
     'SeteoBotonesMod False
   Else
     NuevaSeleccion
 End If
End If

FrmVersus.caption = Aplicacion.SeteoFin

'SetBotonesGeneral False

End Sub

Private Sub MePrepararMod()
Dim col As Long
Dim Row As Long

SeteoBotonesMod False

col = Grilla.ActiveCol
Row = Grilla.ActiveRow

Grilla_DblClick col, Row
    
    
End Sub

Public Sub Modificacion()
Modo = "MOD"
FrmVersus.caption = FrmVersus.caption & " - Modificación y Bajas -"
Me.Show 1
End Sub

Public Function ArmarCondicion()

Dim Con As String

Con = ""
If MskFecha.Text <> "" Then
   Con = Con & " and fch_vuelo between " & func_ToDate(MskFecha.FormattedText) & " And " & func_ToDate(mskFechaH.FormattedText)
End If

If CboCodAeropuerto.Text <> "" Then
   Con = Con & " and cod_depn = '" & CboCodAeropuerto.Text & "'"
End If

If CboEspigon.Text <> "" Then
   Con = Con & " and cod_sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
End If

If OptTipoVuelo(0).Value = True Then
   Con = Con & " and (tipo = 'L' or tipo = 'I' )"
  Else
    If OptTipoVuelo(1).Value = True Then
       Con = Con & " and (tipo = 'S' or tipo = 'I' )"
     Else
    End If
End If

If cboCia.Text <> "" Then
    Con = Con & " and cod_cia_aerea = " & cboCia.ItemData(cboCia.ListIndex)
End If

ArmarCondicion = Con

End Function


Private Sub cboCia_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    cboCia.ListIndex = -1
End If
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
End If
End Sub


Private Sub CboEspigon_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    CboEspigon.ListIndex = -1
End If

End Sub


Private Sub Command1_Click()
Dim fecha As Date

If MskFecha.Text <> "" Then
    fecha = MskFecha.FormattedText
Else
    fecha = Date
End If

frmFecha.MuestroFormFecha fecha

MskFecha.Text = Format$(fecha, FTOFECHA)

End Sub








Private Sub Command2_Click()
Dim fecha As Date

If mskFechaH.Text <> "" Then
    fecha = mskFechaH.FormattedText
Else
    fecha = Date
End If

frmFecha.MuestroFormFecha fecha

mskFechaH.Text = Format$(fecha, FTOFECHA)


End Sub


Private Sub Command3_Click()
frmAbaut.Muestra
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If Shift = 4 Then
    Select Case KeyCode
        Case 67
            Call Toolbar_ButtonClick(Toolbar.Buttons(2))
        Case 78
            Call Toolbar_ButtonClick(Toolbar.Buttons(1))
    End Select
End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Dim sql As String
Dim desc As String

Top = 500
Left = 500

Set cl_versus = New CLVersus

MskFecha.Text = Format(Date - 1, FTOFECHA)
mskFechaH.Text = Format(Date - 1, FTOFECHA)

OptTipoVuelo(2) = True

'arma el SQL para llenar el combo de Aeropuerto
sql = " SELECT cod_depn,descrip FROM baires.dependencia "
sql = sql & " ORDER BY cod_depn"

FuncCbos_LlenarCbo CboCodAeropuerto, sql

sql = "SELECT cod_compania,descrip FROM ventas.companias ORDER BY cod_compania "
FuncCbos_LlenarCboiTEM cboCia, sql

SetBotonesGeneral True

Call Func.Func_ObtenerDesc(" Select to_char(min(fch_actualizado),'dd/mm/yyyy') descrip From estadis.control_carga ", desc)
txtFecha.Text = desc

End Sub

Private Sub Grilla_DblClick(ByVal col As Long, ByVal Row As Long)
Dim FilActual As Long

Dim Compania As Variant
Dim Vuelo As Variant
Dim Atendidos As Variant
Dim Volados As Variant
Dim fecha As Variant
Dim Aeropuerto As Variant
Dim Espigon As Variant
Dim Desc_cia As Variant
Dim Tipo As Variant

FilActual = Row


Grilla.GetText 1, FilActual, Desc_cia

Grilla.GetText 2, FilActual, Vuelo

Grilla.GetText 3, FilActual, Volados

Grilla.GetText 4, FilActual, Atendidos

Grilla.GetText 6, FilActual, Tipo

Grilla.GetText 7, FilActual, Aeropuerto

Grilla.GetText 8, FilActual, Espigon

'Tomo la columna 9 porque tengo el código de compañia
Grilla.GetText 9, FilActual, Compania

'Cargo la clase

cl_versus.fch_vuelo = MskFecha.FormattedText
cl_versus.cod_tipo = Tipo
cl_versus.cod_depn = Aeropuerto
cl_versus.cod_sdep = Espigon
cl_versus.cod_cia_aerea = Compania
cl_versus.cod_vuelo = Vuelo
cl_versus.pax_at = Atendidos
cl_versus.pax_vol = Volados
cl_versus.des_cia = Desc_cia

 
FrmModifica.cargar cl_versus
MeCargarGrilla
    MeSetearBotonesToolBar

End Sub


Private Sub Grilla_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

Dim TipoVuelo As Variant
Dim Aeropuerto As Variant
Dim Espigon As Variant

Grilla.col = 1
Grilla.Row = 1

Grilla.GetText 6, NewRow, TipoVuelo

Select Case TipoVuelo
  Case Is = "S"
     TxtTipoVuelo.Text = "Salida"
  Case Is = "L"
     TxtTipoVuelo.Text = "LLegada"
  Case Is = "I"
     TxtTipoVuelo.Text = "Incierto"
End Select

Grilla.GetText 7, NewRow, Aeropuerto
TxtAeropuerto.Text = Aeropuerto

Grilla.GetText 8, NewRow, Espigon
TxtEspigon.Text = Espigon

End Sub


Private Sub MskFecha_LostFocus()

If MskFecha.Text <> "" Then
If Not IsDate(MskFecha.FormattedText) Then
    MskFecha.Text = Format(Date, FTOFECHA)
End If
    
MskFecha.Text = MskFecha.FormattedText
End If
End Sub


Private Sub mskFechaH_LostFocus()
If mskFechaH.Text <> "" Then
If Not IsDate(mskFechaH.FormattedText) Then
    mskFechaH.Text = Format(Date, FTOFECHA)
End If
    
mskFechaH.Text = mskFechaH.FormattedText
End If

End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
Dim cia As Integer, Tipo As String
Dim pos As String


Select Case Button.Key
    Case "M"
         MePrepararMod
    Case "S"
        Unload Me
    Case "N"
        NuevaSeleccion
    Case "B"
        MeCargarGrilla
    Case "P"
        If cboCia.Text = "" Then
            cia = -1
        Else
            cia = cboCia.ItemData(cboCia.ListIndex)
        End If
        
        If OptTipoVuelo(0).Value Then
            Tipo = "L"
        ElseIf OptTipoVuelo(1).Value Then
            Tipo = "S"
        Else
            Tipo = ""
        End If
        
        FrmImprFrom.TratarImpresionCia MskFecha.FormattedText, LstEspigon.List(CboEspigon.ListIndex), cia, Tipo
        
End Select


End Sub


