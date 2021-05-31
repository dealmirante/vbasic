VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEstimado 
   Caption         =   "Modelización de Estimaciones"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar 
      Height          =   420
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   741
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "A"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "B"
            Object.ToolTipText     =   "Llenar %"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "C"
            Object.ToolTipText     =   "Guardar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "S"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame frRub 
      Caption         =   "Porcentajes por Rubros"
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
      Height          =   2400
      Left            =   75
      TabIndex        =   14
      Top             =   2640
      Width           =   2700
      Begin FPSpread.vaSpread sprRubro 
         Height          =   2040
         Left            =   150
         OleObjectBlob   =   "frmEstimado.frx":0000
         TabIndex        =   15
         Top             =   270
         Width           =   2445
      End
   End
   Begin VB.Frame frDia 
      Caption         =   "Porcentajes Grupos por días"
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
      Height          =   3375
      Left            =   2850
      TabIndex        =   12
      Top             =   1665
      Width           =   5700
      Begin FPSpread.vaSpread sprDia 
         Height          =   3045
         Left            =   135
         OleObjectBlob   =   "frmEstimado.frx":02D7
         TabIndex        =   13
         Top             =   240
         Width           =   5460
      End
   End
   Begin VB.Frame frComi 
      Caption         =   "Comitente"
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
      Height          =   945
      Left            =   60
      TabIndex        =   10
      Top             =   1665
      Width           =   2745
      Begin MSMask.MaskEdBox mskIosc 
         Height          =   285
         Left            =   1605
         TabIndex        =   17
         Top             =   195
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskNoIosc 
         Height          =   285
         Left            =   1605
         TabIndex        =   18
         Top             =   555
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No IOSC"
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
         Index           =   5
         Left            =   735
         TabIndex        =   16
         Top             =   555
         Width           =   780
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IOSC"
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
         Index           =   4
         Left            =   750
         TabIndex        =   11
         Top             =   195
         Width           =   765
      End
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
      Height          =   1215
      Left            =   45
      TabIndex        =   1
      Top             =   420
      Width           =   8475
      Begin VB.ComboBox CboEspigon 
         Height          =   315
         Left            =   2910
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   660
         Width           =   1950
      End
      Begin VB.ComboBox CboCodAeropuerto 
         Height          =   315
         ItemData        =   "frmEstimado.frx":0835
         Left            =   2895
         List            =   "frmEstimado.frx":0837
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   270
         Width           =   1950
      End
      Begin VB.ListBox LstEspigon 
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   915
         Visible         =   0   'False
         Width           =   135
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   300
         Left            =   930
         TabIndex        =   8
         Top             =   270
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtNom 
         Height          =   285
         Left            =   6195
         MaxLength       =   30
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtCod 
         Height          =   285
         Left            =   6195
         MaxLength       =   10
         TabIndex        =   6
         Top             =   285
         Width           =   1170
      End
      Begin MSMask.MaskEdBox mskMes 
         Height          =   300
         Left            =   930
         TabIndex        =   9
         Top             =   675
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
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
         Left            =   1815
         TabIndex        =   23
         Top             =   660
         Width           =   1065
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
         Left            =   1815
         TabIndex        =   22
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mes"
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
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   675
         Width           =   780
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Año"
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
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
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
         Left            =   4995
         TabIndex        =   3
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Identificación"
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
         Left            =   4995
         TabIndex        =   2
         Top             =   285
         Width           =   1185
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6195
      Top             =   255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEstimado.frx":0839
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEstimado.frx":094B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEstimado.frx":117D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEstimado.frx":128F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEstimado.frx":13A1
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEstimado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub L_SetearGrillas()
Dim sql As String
Dim RS As Recordset
Dim FD As String, FH As String
Dim fdAux As Date, i As Integer

sql = "select cod_rubr from baires.rubro order by cod_rubr "
If Aplicacion.ObtenerRsDAO(sql, RS) Then
    Spread_CargarGrilla RS, sprRubro
    Aplicacion.CerrarDAO RS
End If

FD = func_Dia1SegunMes_Anio(mskMes.Text, mskAnio.Text)
FH = func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text)

fdAux = FD
i = 1
sprDia.MaxRows = 0
Do While fdAux <= CDate(FH)
        sprDia.MaxRows = sprDia.MaxRows + 1
        sprDia.SetText 1, i, Format$(fdAux, "dd-mm-yy")
        
        i = i + 1
        fdAux = fdAux + 1

Loop

Spread_PintarfinSemana sprDia

End Sub

Private Sub CboCodAeropuerto_Click()
Dim sql As String

sql = " SELECT cod_sdep,descrip FROM baires.subdependencia "
sql = sql & " WHERE cod_depn = '" & CboCodAeropuerto.Text & "'"
sql = sql & " ORDER BY cod_sdep"
 
FuncCbos_LlenarCboLst CboEspigon, LstEspigon, sql

End Sub


Private Sub Form_Load()
Dim sql As String

Top = 450
Left = 400
Height = 5900
Width = 8800

mskAnio.Text = Year(Date)
mskMes.Text = Month(Date)

mskIosc.Text = 0
mskNoIosc.Text = 0

sql = " SELECT cod_depn,descrip FROM baires.dependencia "
sql = sql & " ORDER BY cod_depn"

FuncCbos_LlenarCbo CboCodAeropuerto, sql

End Sub

Private Sub mskAnio_LostFocus()

If mskAnio.Text < 1996 Or mskAnio > 2050 Then
    mskAnio.Text = Year(Date)
End If
End Sub


Private Sub mskMes_LostFocus()

If mskMes.Text < 1 Or mskMes.Text > 12 Then
    mskMes.Text = Month(Date)
End If

End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)

Select Case Button.Key
    Case "A"
        sprDia.MaxRows = 0
        sprRubro.MaxRows = 0
        mskIosc.Text = 0
        mskNoIosc.Text = 0
        frCab.Enabled = True
        frDia.Enabled = False
        frComi.Enabled = False
        frRub.Enabled = False
        Toolbar.Buttons(1).Enabled = False
        Toolbar.Buttons(2).Enabled = True
        Toolbar.Buttons(4).Enabled = False
    Case "B"
        frCab.Enabled = False
        frDia.Enabled = True
        frComi.Enabled = True
        frRub.Enabled = True
        Toolbar.Buttons(1).Enabled = True
        Toolbar.Buttons(2).Enabled = False
        Toolbar.Buttons(4).Enabled = True
        L_SetearGrillas
        
    Case "C"
    Case "S"
        Unload Me
        
End Select

End Sub

