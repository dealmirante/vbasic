VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmInformes 
   Caption         =   "Listados"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   435
      Left            =   105
      TabIndex        =   21
      Top             =   0
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "A"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "B"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame frCia 
      Caption         =   "Datos para Inf. por Cia"
      Enabled         =   0   'False
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
      Height          =   2310
      Left            =   75
      TabIndex        =   15
      Top             =   2430
      Width           =   4500
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
         Height          =   690
         Left            =   240
         TabIndex        =   20
         Top             =   1500
         Width           =   3885
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
            Left            =   1395
            TabIndex        =   8
            Top             =   330
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
            Left            =   75
            TabIndex        =   7
            Top             =   330
            Width           =   1005
         End
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
            Left            =   2715
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.ComboBox cboCia 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1170
         Width           =   2565
      End
      Begin VB.ComboBox CboEspigon 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   750
         Width           =   2565
      End
      Begin VB.ListBox LstEspigon 
         Height          =   255
         Left            =   3285
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox CboCodAeropuerto 
         Height          =   315
         ItemData        =   "frmInformes.frx":0000
         Left            =   1575
         List            =   "frmInformes.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2565
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
         Left            =   255
         TabIndex        =   19
         Top             =   1170
         Width           =   1245
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
         Left            =   255
         TabIndex        =   18
         Top             =   765
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
         Left            =   255
         TabIndex        =   16
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de informe"
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
      Height          =   1980
      Left            =   75
      TabIndex        =   10
      Top             =   465
      Width           =   4485
      Begin VB.OptionButton optTipo 
         Caption         =   "Totales por Cia Aerea"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   22
         Top             =   690
         Width           =   2190
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   390
         Left            =   2865
         Picture         =   "frmInformes.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         Width           =   465
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   390
         Left            =   2865
         Picture         =   "frmInformes.frx":018E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1005
         Width           =   465
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Detallado por Cia Aerea"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   2190
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Totales por Espigón"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1860
      End
      Begin MSMask.MaskEdBox MskFecha 
         Height          =   315
         Left            =   1605
         TabIndex        =   2
         Top             =   1035
         Width           =   1230
         _ExtentX        =   2170
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
         Left            =   1605
         TabIndex        =   3
         Top             =   1455
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "dd-mm-yyyy"
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Hasta :"
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
         Left            =   255
         TabIndex        =   13
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha :"
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
         Left            =   255
         TabIndex        =   12
         Top             =   1065
         Width           =   1275
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4110
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInformes.frx":0318
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInformes.frx":042A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub


Private Sub botHelpFD_Click()
Dim fecha As Date

If MskFecha.Text <> "" Then
    fecha = MskFecha.FormattedText
Else
    fecha = Date
End If

frmFecha.MuestroFormFecha fecha

MskFecha.Text = Format$(fecha, FTOFECHA)

MskFecha.SetFocus

End Sub


Private Sub botHelpFH_Click()
Dim fecha As Date

If mskFechaH.Text <> "" Then
    fecha = mskFechaH.FormattedText
Else
    fecha = Date
End If

frmFecha.MuestroFormFecha fecha

mskFechaH.Text = Format$(fecha, FTOFECHA)

mskFechaH.SetFocus

End Sub


Private Sub CboCodAeropuerto_Click()
 Dim sql As String

sql = " SELECT cod_sdep,descrip FROM baires.subdependencia "
sql = sql & " WHERE cod_depn = '" & CboCodAeropuerto.Text & "'"
sql = sql & " ORDER BY cod_sdep"
 
FuncCbos_LlenarCboLst CboEspigon, LstEspigon, sql

End Sub


Private Sub Form_Activate()
MskFecha.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 4 Then
    Select Case KeyCode
        Case 83 'Salir
            Call Toolbar1_ButtonClick(Toolbar1.Buttons(16))
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

Top = 800
Left = 2400
Height = 5200
Width = 4800

MskFecha.Text = Format$(Date - 1, FTOFECHA)
mskFechaH.Text = Format$(Date - 1, FTOFECHA)

sql = " SELECT cod_depn,descrip FROM baires.dependencia ORDER BY cod_depn"

FuncCbos_LlenarCbo CboCodAeropuerto, sql

sql = "SELECT cod_compania,descrip FROM ventas.companias ORDER BY cod_compania "
FuncCbos_LlenarCboiTEM cboCia, sql

End Sub

Private Sub MskFecha_LostFocus()
If Not IsDate(MskFecha.FormattedText) Then
    MskFecha.Text = Date
End If

MskFecha.Text = MskFecha.FormattedText

If CDate(mskFechaH.FormattedText) < CDate(MskFecha.FormattedText) Then
    mskFechaH.Text = MskFecha.FormattedText
End If

End Sub


Private Sub mskFechaH_LostFocus()

If Not IsDate(mskFechaH.FormattedText) Then
    mskFechaH.Text = MskFecha.FormattedText
ElseIf CDate(mskFechaH.FormattedText) < CDate(MskFecha.FormattedText) Then
    mskFechaH.Text = MskFecha.FormattedText
End If

mskFechaH.Text = mskFechaH.FormattedText

End Sub


Private Sub optTipo_Click(Index As Integer)

Select Case Index
    Case 0, 2
        mskFechaH.Enabled = True
        frCia.Enabled = False
    Case 1
        mskFechaH.Enabled = False
        frCia.Enabled = True
        
End Select

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim cia As Integer
Dim tipo As String

Select Case Button.Key
    Case "A"
        If optTipo(0).Value Then
            FrmImprFrom.TratarImpresionTot MskFecha.FormattedText, mskFechaH.FormattedText
        ElseIf optTipo(2).Value Then
            FrmImprFrom.TratarImpresionTotCia MskFecha.FormattedText, mskFechaH.FormattedText
        Else
            If cboCia.Text = "" Then
                cia = -1
            Else
                cia = cboCia.ItemData(cboCia.ListIndex)
            End If
            
            If OptTipoVuelo(0).Value Then
                tipo = "L"
            ElseIf OptTipoVuelo(1).Value Then
                tipo = "S"
            Else
                tipo = ""
            End If
            
            FrmImprFrom.TratarImpresionCia MskFecha.FormattedText, LstEspigon.List(CboEspigon.ListIndex), cia, tipo
        End If
    Case "B"
        Unload Me
End Select
End Sub


