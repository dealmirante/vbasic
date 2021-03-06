VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmEstimN 
   Caption         =   "Modelos de Estimados"
   ClientHeight    =   6090
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   9180
   Icon            =   "FrmEstimN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6090
   ScaleWidth      =   9180
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   105
      TabIndex        =   11
      Top             =   0
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   17
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "k"
            Object.ToolTipText     =   "Nueva Seleci?n"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "l"
            Object.ToolTipText     =   "Buscar"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "a"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "b"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "c"
            Object.ToolTipText     =   "Registro Siguiente"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "e"
            Object.ToolTipText     =   "Ultimo Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "f"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "g"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "h"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "i"
            Object.ToolTipText     =   "Abortar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "o"
            Object.ToolTipText     =   "Copiar Modelo"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "m"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "j"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame frLoc 
      Caption         =   "Porcentajes por Local - Rubro"
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
      Height          =   1740
      Left            =   3075
      TabIndex        =   33
      Top             =   4260
      Width           =   6090
      Begin FPSpread.vaSpread sprLoc 
         Height          =   1350
         Left            =   105
         OleObjectBlob   =   "FrmEstimN.frx":0442
         TabIndex        =   34
         Top             =   240
         Width           =   5880
      End
   End
   Begin VB.Frame frDia 
      Caption         =   "Porcentajes Grupos por d?as"
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
      Height          =   2415
      Left            =   3090
      TabIndex        =   25
      Top             =   1845
      Width           =   6045
      Begin FPSpread.vaSpread sprDia 
         Height          =   2070
         Left            =   180
         OleObjectBlob   =   "FrmEstimN.frx":08CD
         TabIndex        =   8
         Top             =   210
         Width           =   5745
      End
   End
   Begin VB.Frame frComi 
      Caption         =   "Comitente Tot. Cia."
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
      TabIndex        =   22
      Top             =   1845
      Width           =   2955
      Begin MSMask.MaskEdBox mskIosc 
         Height          =   285
         Left            =   1365
         TabIndex        =   5
         Top             =   195
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   8
         Format          =   "00.00000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskNoIosc 
         Height          =   285
         Left            =   1350
         TabIndex        =   6
         Top             =   555
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   8
         Format          =   "00.00000"
         PromptChar      =   " "
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
         Left            =   495
         TabIndex        =   24
         Top             =   195
         Width           =   765
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
         Left            =   480
         TabIndex        =   23
         Top             =   555
         Width           =   780
      End
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
      Height          =   3180
      Left            =   75
      TabIndex        =   21
      Top             =   2820
      Width           =   2940
      Begin FPSpread.vaSpread sprRubro 
         Height          =   2775
         Left            =   105
         OleObjectBlob   =   "FrmEstimN.frx":0E33
         TabIndex        =   7
         Top             =   270
         Width           =   2745
      End
   End
   Begin VB.TextBox txtReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7245
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   435
      Width           =   465
   End
   Begin VB.TextBox txtCantReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7995
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   450
      Width           =   480
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
      Height          =   1410
      Left            =   75
      TabIndex        =   13
      Top             =   435
      Width           =   9060
      Begin VB.OptionButton optTipo 
         Caption         =   "para Pasajeros"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   6705
         TabIndex        =   29
         Top             =   780
         Width           =   1575
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "para Ticket"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   5055
         TabIndex        =   28
         Top             =   1065
         Width           =   1380
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "para Importes"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   5055
         TabIndex        =   27
         Top             =   750
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.TextBox txtSec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6330
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   285
         Width           =   780
      End
      Begin VB.TextBox txtNom 
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1035
         Width           =   3525
      End
      Begin VB.ListBox LstEspigon 
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   915
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox CboCodAeropuerto 
         Height          =   315
         ItemData        =   "FrmEstimN.frx":11DE
         Left            =   2910
         List            =   "FrmEstimN.frx":11E0
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   1950
      End
      Begin VB.ComboBox CboEspigon 
         Height          =   315
         Left            =   2910
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   660
         Width           =   1950
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   300
         Left            =   945
         TabIndex        =   0
         Top             =   255
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskMes 
         Height          =   300
         Left            =   930
         TabIndex        =   1
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
      Begin VB.Label de 
         Caption         =   "de"
         Height          =   255
         Left            =   7695
         TabIndex        =   32
         Top             =   135
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Secuencia"
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
         Left            =   5070
         TabIndex        =   20
         Top             =   285
         Width           =   1185
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
         Left            =   120
         TabIndex        =   19
         Top             =   1035
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A?o"
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
         TabIndex        =   18
         Top             =   270
         Width           =   780
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
         TabIndex        =   17
         Top             =   675
         Width           =   780
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
         Left            =   1815
         TabIndex        =   16
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label LblEspigon 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Espig?n :"
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
         TabIndex        =   15
         Top             =   660
         Width           =   1065
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "chk"
      Height          =   195
      Left            =   7275
      TabIndex        =   12
      Top             =   90
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label LblCodAeropuerto 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estimado"
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
      Index           =   1
      Left            =   5865
      TabIndex        =   31
      Top             =   30
      Width           =   915
   End
   Begin VB.Label labValor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   6795
      TabIndex        =   30
      Top             =   30
      Width           =   2070
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":11E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":14FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":1816
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":1B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":1E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":2164
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":247E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":2798
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":2AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":2DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":2EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmEstimN.frx":2FF0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmEstimN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As Recordset
Dim rsValores As Recordset
Dim rsPorcEsp As Recordset

Dim cl_Estim As CLEstimado

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$

Dim ultCol As Integer

Public Sub Altas()
    
    Modo = "ALTA"
    SetearBotonesAltas
    Me.Show 1

End Sub
Private Sub L_AltasDatos()

If L_TodoCargado Then
    FrmEstim.caption = Aplicacion.SeteoProceso(FrmEstim.caption)

    Aplicacion.ComienzoTrans

    MeLlenarObjeto

    If cl_Estim.Insert_Porciento() Then
        L_SecuenciaAumentarUna L_SegunOptTipo
        txtSec.Text = FuncLocal_Secuencia(mskAnio.Text, mskMes.Text, L_SegunOptTipo)
        txtNom.Text = L_NombreSujerido
        chk.Value = 0
        Aplicacion.TerminarConExitoTrans
    Else
        Aplicacion.TerminarConErrorTrans
    End If


    FrmEstim.caption = Aplicacion.SeteoFin

'Else
'    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub L_CargoComi(anio As Integer, Mes As Integer, sec As Integer, tipo As String)
Dim sql As String
Dim rsPorc As Recordset

sql = "Select comi,porcentaje "
sql = sql & " From Estadis.porciento_comi "
sql = sql & " Where anio = " & anio
sql = sql & " And mes =" & Mes
sql = sql & " And tipo_porc = '" & tipo & "'"
sql = sql & " And secuencia = " & sec

If Aplicacion.ObtenerRsDAO(sql, rsPorc) Then
    If Aplicacion.CantReg(rsPorc) > 0 Then
        Do While Not rsPorc.EOF
            If rsPorc!comi = "I" Then
                mskIosc.Text = rsPorc!porcentaje
            Else
                mskNoIosc.Text = rsPorc!porcentaje
            End If
            rsPorc.MoveNext
        Loop
    Else
        mskIosc.Text = 0
        mskNoIosc.Text = 0
    End If
    Aplicacion.CerrarDAO rsPorc
End If

End Sub

Private Sub L_CargoRubro(anio As Integer, Mes As Integer, sec As Integer, tipo As String)
Dim sql As String
Dim rsPorc As Recordset
Dim rsRub As Recordset
Dim rub As String

sql = "Select cod_rubro,porcentaje,comi "
sql = sql & " From Estadis.porciento_rubro "
sql = sql & " Where anio = " & anio
sql = sql & " And mes =" & Mes
sql = sql & " And tipo_porc = '" & tipo & "'"
sql = sql & " And secuencia = " & sec
sql = sql & " ORDER BY cod_rubro,comi "

If Aplicacion.ObtenerRsDAO(sql, rsPorc) Then
    If Aplicacion.CantReg(rsPorc) > 0 Then
        Do While Not rsPorc.EOF
            rub = rsPorc!cod_rubro
            sprRubro.MaxRows = sprRubro.MaxRows + 1
            sprRubro.SetText 1, sprRubro.MaxRows, Trim(rsPorc!cod_rubro)
            Do While rub = rsPorc!cod_rubro
                Select Case rsPorc!comi
                    Case "T"
                        sprRubro.SetText 2, sprRubro.MaxRows, str(rsPorc!porcentaje)
                    Case "I"
                        sprRubro.SetText 3, sprRubro.MaxRows, str(rsPorc!porcentaje)
                    Case "N"
                        sprRubro.SetText 4, sprRubro.MaxRows, str(rsPorc!porcentaje)
                End Select
            rsPorc.MoveNext
            If rsPorc.EOF Then
                Exit Do
            End If
            Loop
        Loop
        'Spread_CargarGrilla rsPorc, sprRubro
    Else
        sql = "select cod_rubr,0 cant from baires.rubro WHERE COD_RUBR <> 'REG' order by cod_rubr "
        If Aplicacion.ObtenerRsDAO(sql, rsRub) Then
            Spread_CargarGrilla rsRub, sprRubro
            Aplicacion.CerrarDAO rsRub
        End If
    End If
    Spread_TotalesGrillas sprRubro, 1, 2
    Aplicacion.CerrarDAO rsPorc
End If

End Sub
Private Sub L_CargoLocal(anio As Integer, Mes As Integer, sec As Integer, tipo As String)
Dim sql As String, cod_rub As String
Dim rsPorc As Recordset
Dim rsRub As Recordset
Dim i As Integer, c As Integer

sql = "Select cod_rubro,cod_local,porcentaje "
sql = sql & " From Estadis.porciento_local "
sql = sql & " Where anio = " & anio
sql = sql & " And mes =" & Mes
sql = sql & " And tipo_porc = '" & tipo & "'"
sql = sql & " And secuencia = " & sec
sql = sql & " Order By cod_rubro,cod_local "

If Aplicacion.ObtenerRsDAO(sql, rsPorc) Then
'    L_SeteoGrillaLoc
        sql = "select cod_rubr cant from baires.rubro order by cod_rubr "
        If Aplicacion.ObtenerRsDAO(sql, rsRub) Then
            Spread_CargarGrilla rsRub, sprLoc
            Aplicacion.CerrarDAO rsRub
        End If
    
    If Aplicacion.CantReg(rsPorc) > 0 Then

        i = 1
        Do While Not rsPorc.EOF
            cod_rub = rsPorc!cod_rubro
            c = 2
            Do While cod_rub = rsPorc!cod_rubro
                sprLoc.SetText c, i, str(rsPorc!porcentaje)
                c = c + 1
                rsPorc.MoveNext
                If rsPorc.EOF Then
                    Exit Do
                End If
            Loop
            i = i + 1
        Loop
    Else
        For c = 2 To ultCol - 1
            For i = 1 To sprLoc.MaxRows
                sprLoc.SetText c, i, "0"
            Next
        Next
    
    End If
    L_SeteoGrillaLoc
    Aplicacion.CerrarDAO rsPorc
End If

End Sub

Private Sub L_CargoDia(anio As Integer, Mes As Integer, sec As Integer, tipo As String)
Dim sql As String
Dim rsPorc As Recordset
Dim FD As String, FH As String
Dim fdAux As Date, i As Integer

sql = "Select fch_dia,Grupo,porcentaje "
sql = sql & " From Estadis.porciento_diario "
sql = sql & " Where anio = " & anio
sql = sql & " And mes =" & Mes
sql = sql & " And tipo_porc = '" & tipo & "'"
sql = sql & " And secuencia = " & sec
sql = sql & " Order By fch_dia,grupo "

If Aplicacion.ObtenerRsDAO(sql, rsPorc) Then
    If Aplicacion.CantReg(rsPorc) > 0 Then
        Do While Not rsPorc.EOF
            FD = rsPorc!fch_dia
            sprDia.MaxRows = sprDia.MaxRows + 1
            sprDia.SetText 1, sprDia.MaxRows, Format$(rsPorc!fch_dia, "dd-mm-yy")
            Do While FD = CDate(rsPorc!fch_dia)
                Select Case rsPorc!Grupo
                    Case "T"
                        sprDia.SetText 2, sprDia.MaxRows, str(rsPorc!porcentaje)
                    Case "A"
                        sprDia.SetText 3, sprDia.MaxRows, str(rsPorc!porcentaje)
                    Case "B"
                        sprDia.SetText 4, sprDia.MaxRows, str(rsPorc!porcentaje)
                    Case "C"
                        sprDia.SetText 5, sprDia.MaxRows, str(rsPorc!porcentaje)
                End Select
                rsPorc.MoveNext
                If rsPorc.EOF Then
                    Exit Do
                End If
            Loop
        Loop
    Else
        FD = func_Dia1SegunMes_Anio(mskMes.Text, mskAnio.Text)
        FH = func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text)
        
        fdAux = FD
        i = 1
        sprDia.MaxRows = 0
        Do While fdAux <= CDate(FH)
                sprDia.MaxRows = sprDia.MaxRows + 1
                sprDia.SetText 1, i, Format$(fdAux, "dd-mm-yy")
                sprDia.SetText 2, i, "0"
                i = i + 1
                fdAux = fdAux + 1
        Loop
    End If
    Spread_PintarfinSemana sprDia
    Spread_TotalesGrillas sprDia, 1, 2
    Aplicacion.CerrarDAO rsPorc
End If

End Sub


Private Sub L_CargoValor()
Dim sql$
Dim rs As Recordset
Dim valor As Double


    sql$ = ""
    sql$ = sql$ & "Select * From estadis.porciento_espigon "
    sql$ = sql$ & " Where anio = " & mskAnio.Text
    sql$ = sql$ & " And mes =" & mskMes.Text
    sql$ = sql$ & " And cod_depn ='" & CboCodAeropuerto.Text & "'"
    sql$ = sql$ & " And cod_sdep ='" & LstEspigon.List(CboEspigon.ListIndex) & "'"
    sql$ = sql$ & " And Tipo_porc = '" & L_SegunOptTipo & "'"
    
    Select Case L_SegunOptTipo
        Case "I"
            valor = rsValores!Importe
        Case "T"
            valor = rsValores!ticket
        Case "P"
            valor = rsValores!Pax
    End Select
    
    If Aplicacion.ObtenerRsDAO(sql$, rs) Then
        labValor.caption = Format$(valor * rs!porcentaje / 100, "#,#.00")
    End If
End Sub

Private Function L_ControlRubIN() As Integer
Dim i
Dim val1 As Variant, val2 As Variant
Dim Primero As Single, numero As Integer

sprRubro.GetText 3, 1, val1
sprRubro.GetText 4, 1, val2
Primero = Val(val1) + Val(val2)

numero = Primero
If Primero = 100 Or Primero = 0 Then
    For i = 1 To sprRubro.MaxRows - 1
        sprRubro.GetText 3, i, val1
        sprRubro.GetText 4, i, val2
        If Val(val1) + Val(val2) <> Primero Then
            numero = -1
            Exit For
        End If
    Next
End If
L_ControlRubIN = numero

End Function
Private Sub L_CopiarModelo()

FrmEstim.caption = Aplicacion.SeteoProceso(FrmEstim.caption)
    
txtSec.Text = FuncLocal_Secuencia(mskAnio.Text, mskMes.Text, L_SegunOptTipo)

Aplicacion.ComienzoTrans

MeLlenarObjeto
cl_Estim.Descrip = L_NombreSujerido
'cl_Estim.mes = Month(Date)
cl_Estim.Mes = mskMes.Text
cl_Estim.anio = mskAnio.Text

If cl_Estim.Insert_Porciento() Then
   chk.Value = 0
   L_SecuenciaAumentarUna L_SegunOptTipo
   Aplicacion.TerminarConExitoTrans
   MeReconsultar
   MeSetearBotonesToolBar
Else
   Aplicacion.TerminarConErrorTrans
End If

FrmEstim.caption = Aplicacion.SeteoFin

End Sub

Private Sub L_ImprimirCR()

On Error GoTo ErrRp:

    FrmEstimN.caption = Aplicacion.SeteoProceso(FrmEstimN.caption)

    Aplicacion.ComienzoTrans

    MeLlenarObjetoCR

    If cl_Estim.Insert_PrintModelo Then
        If L_SegunOptTipo = "I" Then
            MeLlenarObjetoCRRub
            If cl_Estim.Insert_PrintRubros Then
                Aplicacion.TerminarConExitoTrans
            Else
                Aplicacion.TerminarConErrorTrans
            End If
        Else
            Aplicacion.TerminarConExitoTrans
        End If
    Else
        Aplicacion.TerminarConErrorTrans
    End If
        
    FrmEstimN.caption = Aplicacion.SeteoFin
    FrmImprFrom.TratarImpresionModelos labValor.caption, L_SegunOptTipo
    Exit Sub
ErrRp:
    FrmEstimN.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub

Private Function L_NombreSujerido() As String

L_NombreSujerido = mskMes.Text & "-" & mskAnio.Text & "/" & txtSec.Text _
                    & "(" & L_SegunOptTipo & " " & LstEspigon.List(CboEspigon.ListIndex) & ") " & Left(Aplicacion.username, 10)
End Function

Private Function L_ObtenerPorcRubro(cod As Variant) As Single
Dim i
Dim valor As Variant, rub As Variant

valor = 0
For i = 1 To sprRubro.MaxRows
    sprRubro.GetText 1, i, rub
    If rub = cod Then
        sprRubro.GetText 2, i, valor
        Exit For
    End If
Next
L_ObtenerPorcRubro = Val(valor)
End Function

Private Sub L_PonerValores()
Dim tipoMod
Dim valor As Double

On Error GoTo ErrVal:

rsPorcEsp.MoveFirst
tipoMod = L_SegunOptTipo
Select Case tipoMod
    Case "I"
        valor = rsValores!Importe
    Case "T"
        valor = rsValores!ticket
    Case "P"
        valor = rsValores!Pax
End Select

Do While Not rsPorcEsp.EOF
    If rsPorcEsp!tipo_porc = tipoMod Then
        labValor.caption = Format$(valor * rsPorcEsp!porcentaje / 100, "#,#.00")
    Exit Do
    End If
    rsPorcEsp.MoveNext
Loop

ErrVal:
    Exit Sub
End Sub

Private Function L_SecuenciaAumentarUna(tipo As String) As Boolean
Dim sql As String

sql = "BEGIN Estadis.procesos_estadisticos.act_Sec_Modelo ("
sql = sql & mskAnio.Text & ","
sql = sql & mskMes.Text & ","
sql = sql & "'" & tipo & "'); "
sql = sql & " END ; "

L_SecuenciaAumentarUna = Aplicacion.EjecutarDAO(sql)

End Function

Private Function L_SecuenciaControl() As Boolean
Dim sql As String
Dim rs As Recordset

    If FuncLocal_Secuencia(mskAnio.Text, mskMes.Text, L_SegunOptTipo) >= 0 Then
        L_SecuenciaControl = True
        
        If L_SecuenciasAumentar Then
            txtSec.Text = FuncLocal_Secuencia(mskAnio.Text, mskMes.Text, L_SegunOptTipo)
            Aplicacion.TerminarConExitoTrans
        Else
            Aplicacion.TerminarConErrorTrans
        End If

        L_ValoresInic
        Call CboEspigon_Click
    Else
        txtSec.Text = ""
        L_SecuenciaControl = False
    End If

End Function

Private Function L_SecuenciasAumentar()
Dim tipoMod As String
Dim i
Dim salir As Boolean

Aplicacion.ComienzoTrans

For i = 1 To 3
    Select Case i
        Case 1
            tipoMod = "I"
        Case 2
            tipoMod = "T"
        Case 3
            tipoMod = "P"
    End Select

    salir = L_SecuenciaAumentarUna(tipoMod)
    
    If Not salir Then
        Exit For
    End If
Next

L_SecuenciasAumentar = salir
'If salir Then
'    Aplicacion.TerminarConExitoTrans
'Else
'    Aplicacion.TerminarConErrorTrans
'End If

End Function

Private Sub L_SegunDatoTipo(tipo As String)

Select Case tipo
    Case "I"
        optTipo(0).Value = True
    Case "T"
        optTipo(1).Value = True
    Case "P"
        optTipo(2).Value = True
End Select

End Sub

Private Function L_SegunOptTipo() As String
If optTipo(0).Value Then
    L_SegunOptTipo = "I"
ElseIf optTipo(1).Value Then
    L_SegunOptTipo = "T"
ElseIf optTipo(2).Value Then
    L_SegunOptTipo = "P"
Else
    L_SegunOptTipo = "V"
End If
End Function

Private Sub L_SeteoGrillaLoc()
Dim sql As String
Dim rs As Recordset
Dim i As Integer, j As Integer

sprLoc.col = ultCol
If ultCol > 2 Then
    For j = 1 To sprLoc.MaxRows
        sprLoc.Row = j
        sprLoc.Formula = ""
        sprLoc.Lock = False
    Next
    
    For i = 2 To ultCol
        sprLoc.col = i
        sprLoc.ColHidden = True
    Next
End If

sql = "SELECT cod_loc from baires.local L,ventas.Apertura_sdep A "
sql = sql & " WHERE L.cod_depn = '" & CboCodAeropuerto.Text & "'"
sql = sql & " AND L.cod_loc = A.cod_local "
sql = sql & " AND cod_ssdep ='" & LstEspigon.List(CboEspigon.ListIndex) & "'"
sql = sql & " AND tipo_loc <> 'B' "

If Aplicacion.ObtenerRsDAO(sql, rs) Then
    i = 2
    Do While Not rs.EOF
        sprLoc.col = i
        sprLoc.ColHidden = False
        sprLoc.SetText i, 0, Trim(rs!cod_loc)
        
        rs.MoveNext
        i = i + 1
    Loop
    sprLoc.col = i
    sprLoc.ColHidden = False
    sprLoc.SetText i, 0, "Total"
    ultCol = i
    
    For j = 1 To sprLoc.MaxRows
        Spread_TotalesGrillasCol sprLoc, 2, i - 1, j, i
        sprLoc.Row = j
        sprLoc.Lock = True
    Next
    
    Aplicacion.CerrarDAO rs
End If
End Sub

Private Sub L_ValoresInic()
Dim sql As String

sql$ = ""
sql$ = sql$ & "Select importe,ticket,pax From estadis.modelo_estim "
sql$ = sql$ & " Where anio = " & mskAnio.Text
sql$ = sql$ & " And mes =" & mskMes.Text

Call Aplicacion.ObtenerRsDAO(sql$, rsValores)

End Sub

Private Sub MeImpDatos()
Dim nom As String, NombreArchivo As String


'On Error GoTo ErrFoto:

'FrmEstim.caption = Aplicacion.SeteoProceso(FrmEstim.caption)
'
'    NombreArchivo = RutaFotos & "P" & txtLegajo.Text & ".bmp"
'    Nom = txtApe.Text
'
'    If Dir(NombreArchivo) <> "" Then
'        Image1.Picture = LoadPicture(NombreArchivo)
'        Printer.PaintPicture Image1, 8000, 2000, 2800, 2200
'    End If
'
'
'Printer.FontBold = True
'
'Printer.CurrentX = 10
'Printer.CurrentY = 10
'Printer.FontSize = 10
'
'Printer.Print "  "
'
'Printer.CurrentX = 1000
'Printer.CurrentY = 1000
'Printer.FontSize = 18
'
'Printer.Print txtApe.Text & ", " & txtNom.Text
'
'Printer.CurrentX = 1000
'Printer.CurrentY = 2000
'Printer.FontSize = 10
''Printer.FontBold = False
'Printer.Print "Legajo  : "
'
'Printer.FontBold = False

'Printer.CurrentX = 10
'Printer.CurrentY = 10
'Printer.FontSize = 10
'
'Printer.Print "  "
'
'Printer.CurrentX = 2000
'Printer.CurrentY = 2000
'Printer.FontSize = 10
'Printer.Print txtLegajo.Text
'
'Printer.EndDoc
'
'ErrFoto:
'    Aplicacion.SeteoFin
'    Exit Sub
        
End Sub

Private Sub MeLlenarObjetoCRRub()
Dim cl_Gen As CLGeneric
Dim i, valor As Variant, rub As Variant, ValorGral As Variant
Dim PorcRub As Single

Set cl_Estim.col_PrnPorcModG = New Collection

For i = 1 To sprRubro.MaxRows - 1
    Set cl_Gen = New CLGeneric
            
    cl_Gen.tipo = "B"
    
    sprRubro.GetText 1, i, valor
    cl_Gen.Identif = valor

    sprRubro.GetText 2, i, ValorGral
    cl_Gen.Porc = ValorGral
    
    sprRubro.GetText 3, i, valor
    cl_Gen.PorcGA = ValorGral * valor / 100
    
    sprRubro.GetText 4, i, valor
    cl_Gen.PorcGB = ValorGral * valor / 100
    
    cl_Estim.col_PrnPorcModG.Add cl_Gen
Next

Set cl_Gen = New CLGeneric
cl_Gen.tipo = "A"
cl_Gen.Identif = "IOS"
cl_Gen.Porc = mskIosc.Text

cl_Estim.col_PrnPorcModG.Add cl_Gen

Set cl_Gen = New CLGeneric
cl_Gen.tipo = "A"
cl_Gen.Identif = "NOI"
cl_Gen.Porc = mskNoIosc.Text

cl_Estim.col_PrnPorcModG.Add cl_Gen

'PORC DE PARTICIPACION POR LOCAL
For i = 1 To sprLoc.MaxRows
    Set cl_Gen = New CLGeneric
    
    cl_Gen.tipo = "L"
    
    sprLoc.GetText 1, i, rub
    cl_Gen.Identif = rub
    Select Case LstEspigon.List(CboEspigon.ListIndex)
        Case "AEP", "INTAL", "INTAS"
            sprLoc.GetText 2, i, valor
            cl_Gen.Porc = valor
            sprLoc.GetText 3, i, valor
            cl_Gen.PorcGA = valor
        Case "INTA"
            sprLoc.GetText 2, i, valor
            cl_Gen.Porc = valor
            sprLoc.GetText 3, i, valor
            cl_Gen.PorcGA = valor
            sprLoc.GetText 4, i, valor
            cl_Gen.PorcGB = valor
            sprLoc.GetText 5, i, valor
            cl_Gen.PorcGC = valor
         Case "INTB"
            sprLoc.GetText 2, i, valor
            cl_Gen.Porc = valor
            sprLoc.GetText 3, i, valor
            cl_Gen.PorcGA = valor
            sprLoc.GetText 4, i, valor
            cl_Gen.PorcGB = valor
        
    End Select
    
cl_Estim.col_PrnPorcModG.Add cl_Gen

Next

'PORC DE PARTICIPACION POR LOCAL PONDERADO POR RUBRO
For i = 1 To sprLoc.MaxRows
    Set cl_Gen = New CLGeneric
    
    cl_Gen.tipo = "R"
    
    sprLoc.GetText 1, i, rub
    cl_Gen.Identif = rub
    PorcRub = L_ObtenerPorcRubro(rub) / 100
    
    Select Case LstEspigon.List(CboEspigon.ListIndex)
        Case "AEP", "INTAL", "INTAS"
            sprLoc.GetText 2, i, valor
            cl_Gen.Porc = valor * PorcRub
            sprLoc.GetText 3, i, valor
            cl_Gen.PorcGA = valor * PorcRub
        Case "INTA"
            sprLoc.GetText 2, i, valor
            cl_Gen.Porc = valor * PorcRub
            sprLoc.GetText 3, i, valor
            cl_Gen.PorcGA = valor * PorcRub
            sprLoc.GetText 4, i, valor
            cl_Gen.PorcGB = valor * PorcRub
            sprLoc.GetText 5, i, valor
            cl_Gen.PorcGC = valor * PorcRub
         Case "INTB"
            sprLoc.GetText 2, i, valor
            cl_Gen.Porc = valor * PorcRub
            sprLoc.GetText 3, i, valor
            cl_Gen.PorcGA = valor * PorcRub
            sprLoc.GetText 4, i, valor
            cl_Gen.PorcGB = valor * PorcRub
        
    End Select
    
cl_Estim.col_PrnPorcModG.Add cl_Gen

Next

End Sub

Public Sub Modificacion()
Dim i
Modo = "MOD"
For i = 0 To 2
    optTipo(i).Value = False
Next

frComi.Enabled = False
frRub.Enabled = False
'frDia.Enabled = False
'frLoc.Enabled = False
Me.Show 1
End Sub
Private Sub NuevaSeleccion()
Dim i%

If Modo = "MOD" Then
    SetBotonesGeneral False
    mskAnio.Text = ""
    mskMes.Text = ""
    CboCodAeropuerto.ListIndex = -1
Else
    If chk.Value Then
        If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
            L_AltasDatos
        End If
    End If
End If
'Limpiar campos de pantallas
Set cl_Estim = New CLEstimado

txtNom.Text = ""
txtSec.Text = ""

mskIosc.Text = 0
mskNoIosc.Text = 0

sprRubro.MaxRows = 0
sprDia.MaxRows = 0
sprLoc.MaxRows = 0

mskMes.SetFocus

chk.Value = 0

End Sub

Private Sub MeAbortarMod()
    
    SeteoBotonesMod True
    
    Tollbar.Buttons(2).Enabled = False
    
    MeSetearBotonesToolBar
    
    MellenarPantalla
    
End Sub

Private Sub MeActualizar()

If L_TodoCargado Then

    FrmEstim.caption = Aplicacion.SeteoProceso(FrmEstim.caption)

    Aplicacion.ComienzoTrans

    MeLlenarObjeto


    If cl_Estim.Update_Porciento Then '
        Aplicacion.TerminarConExitoTrans
        SeteoBotonesMod True
    
        If MeReconsultar > 0 Then
    
        Tollbar.Buttons(2).Enabled = False
    
        MeSetearBotonesToolBar
        Else
                NuevaSeleccion
        End If
    
    Else
        Aplicacion.TerminarConErrorTrans
    End If


    FrmEstim.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub MeCargarDatos()
Dim sql$

FrmEstim.caption = Aplicacion.SeteoProceso(FrmEstim.caption)
        
    CondConsulta = ArmarCondicion

    sqlGral$ = " "
    sqlGral$ = sqlGral$ & " SELECT anio,"
    sqlGral$ = sqlGral$ & " mes, "
    sqlGral$ = sqlGral$ & " tipo_porc,"
    sqlGral$ = sqlGral$ & " secuencia,"
    sqlGral$ = sqlGral$ & " cod_depn,"
    sqlGral$ = sqlGral$ & " cod_sdep,"
    sqlGral$ = sqlGral$ & " descrip,"
    sqlGral$ = sqlGral$ & " usuario "
    sqlGral$ = sqlGral$ & " from Estadis.porciento_Cabezera "
    sqlGral$ = sqlGral$ & CondConsulta
    sqlGral$ = sqlGral$ & " Order By anio,mes,secuencia "
    
If Aplicacion.ObtenerRsDAO(sqlGral$, rs) Then
    txtCantReg.Text = Aplicacion.CantReg(rs)
    If txtCantReg.Text > 0 Then
        txtReg.Text = 1
        SetBotonesGeneral True
        L_SeteoGrillaLoc
        MellenarPantalla
        MeSetearBotonesToolBar
    Else
        txtReg.Text = 0
    End If
End If

FrmEstim.caption = Aplicacion.SeteoFin

End Sub
Private Sub MeEliminar()

If Aplicacion.username = rs!usuario Then
    MeLlenarObjeto
    
    FrmEstim.caption = Aplicacion.SeteoProceso(FrmEstim.caption)
    
    Aplicacion.ComienzoTrans
    
    If cl_Estim.Delete_Porciento Then
        Aplicacion.TerminarConExitoTrans
        If MeReconsultar > 0 Then
            'Tollbar.Buttons(2).Enabled = False
            'MeSetearBotonesToolBar
        Else
            NuevaSeleccion
        End If
    
    Else
        Aplicacion.TerminarConErrorTrans
    End If
    
    FrmEstim.caption = Aplicacion.SeteoFin
Else
    MsgBox "No puede Eliminar el modelo de " & rs!usuario, vbCritical + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub MeLlenarObjetoCR()
Dim cl_Gen As CLGeneric
Dim i, j, valor As Variant
Dim hay As Boolean

cl_Estim.anio = mskAnio.Text
cl_Estim.Mes = mskMes.Text
cl_Estim.depn = CboCodAeropuerto.Text
cl_Estim.Sdep = LstEspigon.List(CboEspigon.ListIndex)
cl_Estim.sec = txtSec.Text
cl_Estim.tipo = L_SegunOptTipo
cl_Estim.Descrip = txtNom.Text

Set cl_Estim.col_PrnPorcModG = New Collection

'Porcentajes para dias
sprDia.GetText 2, sprDia.MaxRows, valor
If Val(valor) = 100 Then
    For i = 1 To sprDia.MaxRows - 1
        Set cl_Gen = New CLGeneric
            
'        cl_Gen.Grupo = "T"
        sprDia.GetText 1, i, valor
        cl_Gen.Identif = Format$(valor, FTOFECHA)
        
        sprDia.GetText 2, i, valor
        cl_Gen.Porc = valor
        
        sprDia.GetText 3, i, valor
        cl_Gen.PorcGA = valor

        sprDia.GetText 4, i, valor
        cl_Gen.PorcGB = valor
        
        sprDia.GetText 5, i, valor
        cl_Gen.PorcGC = valor

        cl_Estim.col_PrnPorcModG.Add cl_Gen
    Next
End If

End Sub
Private Sub MeLlenarObjeto()
Dim cl_Gen As CLGeneric
Dim i, j, valor As Variant, rub As Variant
Dim hay As Boolean

cl_Estim.anio = mskAnio.Text
cl_Estim.Mes = mskMes.Text
cl_Estim.depn = CboCodAeropuerto.Text
cl_Estim.Sdep = LstEspigon.List(CboEspigon.ListIndex)
cl_Estim.sec = txtSec.Text
cl_Estim.tipo = L_SegunOptTipo
cl_Estim.Descrip = txtNom.Text

Set cl_Estim.col_PorcComi = New Collection
Set cl_Estim.col_PorcRub = New Collection
Set cl_Estim.col_PorcDia = New Collection
Set cl_Estim.col_PorcLoc = New Collection

'Porcentajes para comitentes
If Val(mskIosc.Text) + Val(mskNoIosc.Text) = 100 Then
    Set cl_Gen = New CLGeneric
    cl_Gen.Identif = "I"
    cl_Gen.Porc = mskIosc.Text
    
    cl_Estim.col_PorcComi.Add cl_Gen
    
    Set cl_Gen = New CLGeneric
    cl_Gen.Identif = "N"
    cl_Gen.Porc = mskNoIosc.Text
    
    cl_Estim.col_PorcComi.Add cl_Gen
End If

'Porcentajes para rubros
sprRubro.GetText 2, sprRubro.MaxRows, valor
If Val(valor) = 100 Then
    For i = 1 To sprRubro.MaxRows - 1
        sprRubro.GetText 1, i, rub
        For j = 2 To sprRubro.MaxCols
            Set cl_Gen = New CLGeneric
                
            cl_Gen.Identif = rub
        
            sprRubro.GetText j, i, valor
            cl_Gen.Porc = valor
            Select Case j
                Case 2
                    cl_Gen.tipo = "T"
                Case 3
                    cl_Gen.tipo = "I"
                Case 4
                    cl_Gen.tipo = "N"
            End Select
            cl_Estim.col_PorcRub.Add cl_Gen
        Next
    Next
End If

'Porcentajes para locales
sprLoc.GetText ultCol, 1, valor
 If Val(valor) = 100 Then
 For i = 1 To sprLoc.MaxRows
        
        For j = 2 To ultCol - 1
            Set cl_Gen = New CLGeneric
            
            sprLoc.GetText j, 0, valor
            cl_Gen.Locale = valor
    
            sprLoc.GetText 1, i, rub
            cl_Gen.Identif = rub
    
            sprLoc.GetText j, i, valor
            cl_Gen.Porc = Val(valor)
    
            cl_Estim.col_PorcLoc.Add cl_Gen
        Next
Next
End If


'Porcentajes para dias
sprDia.GetText 2, sprDia.MaxRows, valor
If Val(valor) = 100 Then
    For i = 1 To sprDia.MaxRows - 1
        Set cl_Gen = New CLGeneric
            
        cl_Gen.Grupo = "T"
        
        sprDia.GetText 1, i, valor
        cl_Gen.Identif = Format$(valor, FTOFECHA)
        
        sprDia.GetText 2, i, valor
        cl_Gen.Porc = valor
        
        cl_Estim.col_PorcDia.Add cl_Gen
    Next
End If

'Porcentajes para dias - grupos
hay = True
For i = 1 To sprDia.MaxRows - 1
     sprDia.GetText 6, i, valor
     If Val(valor) <> 100 Then
        hay = False
        Exit For
     End If
Next
If hay Then
    For i = 1 To sprDia.MaxRows - 1
        'Grupo A
        Set cl_Gen = New CLGeneric
            
        cl_Gen.Grupo = "A"
        
        sprDia.GetText 1, i, valor
        cl_Gen.Identif = Format$(valor, FTOFECHA)
        
        sprDia.GetText 3, i, valor
        cl_Gen.Porc = valor
        
        cl_Estim.col_PorcDia.Add cl_Gen
        
        'Grupo B
        Set cl_Gen = New CLGeneric
            
        cl_Gen.Grupo = "B"
        
        sprDia.GetText 1, i, valor
        cl_Gen.Identif = Format$(valor, FTOFECHA)
        
        sprDia.GetText 4, i, valor
        cl_Gen.Porc = valor
        
        cl_Estim.col_PorcDia.Add cl_Gen
        
        'Grupo C
        Set cl_Gen = New CLGeneric
            
        cl_Gen.Grupo = "C"
        
        sprDia.GetText 1, i, valor
        cl_Gen.Identif = Format$(valor, FTOFECHA)
        
        sprDia.GetText 5, i, valor
        cl_Gen.Porc = valor
        
        cl_Estim.col_PorcDia.Add cl_Gen
        
    Next
End If

End Sub



Private Function L_TodoCargado() As Boolean
Dim valor As Variant
Dim i
Dim salida As Boolean
Dim Primero As Integer

salida = True
If CboCodAeropuerto.Text <> "" And CboEspigon.Text <> "" And txtSec.Text <> "" Then
    If Val(mskIosc.Text) + Val(mskNoIosc.Text) = 100 Or _
    Val(mskIosc.Text) + Val(mskNoIosc.Text) = 0 Then
    
        sprRubro.GetText 2, sprRubro.MaxRows, valor
        If (Val(valor) = 100 And L_ControlRubIN = 100) _
           Or (Val(valor) = 0 And L_ControlRubIN = 0) Then
            
            
            sprLoc.GetText ultCol, 1, valor
            Primero = valor
            For i = 1 To sprLoc.MaxRows - 1
                sprLoc.GetText ultCol, i, valor
                If Val(valor) <> Primero Then
                    salida = False
                    MsgBox "Alguno de los porcentajes de Local no suman 100% ", vbOKOnly + vbInformation, "ATENCION"
                    Exit For
                End If
            Next
            If salida Then
                sprDia.GetText 2, sprDia.MaxRows, valor
                If Val(Format(valor, "#.0000")) = 100 Or Val(valor) = 0 Then
                        sprDia.GetText 6, 1, valor
                        Primero = valor
                       For i = 1 To sprDia.MaxRows - 1
                            sprDia.GetText 6, i, valor
                            
                            If Val(valor) <> Primero Then
                                salida = False
                                MsgBox "Alguno de los porcentajes de Grupo no suman 100% ", vbOKOnly + vbInformation, "ATENCION"
                                Exit For
                            End If
                       Next
                Else
                    salida = False
                    MsgBox "Los porcentajes de Dias no suman 100% ", vbOKOnly + vbInformation, "ATENCION"
                End If
            End If
        Else
            salida = False
            MsgBox "Los porcentajes de Rubros no suman 100% ", vbOKOnly + vbInformation, "ATENCION"
        End If
            
    Else
        salida = False
        MsgBox "Los porcentajes de Comitentes no son correctos", vbOKOnly + vbInformation, "ATENCION"
    End If
Else
    salida = False
    MsgBox "Los datos de Cabecera no estan completos ", vbOKOnly + vbInformation, "ATENCION"

End If
L_TodoCargado = salida
End Function

Private Sub MellenarPantalla()

mskAnio.Text = rs!anio
mskMes.Text = rs!Mes
txtSec.Text = rs!secuencia

Func_SetearCboSTR CboCodAeropuerto, rs!cod_depn
Func_SetearCboConLst CboEspigon, LstEspigon, rs!cod_sdep

L_SegunDatoTipo rs!tipo_porc

L_CargoComi rs!anio, rs!Mes, rs!secuencia, rs!tipo_porc
L_CargoDia rs!anio, rs!Mes, rs!secuencia, rs!tipo_porc
L_CargoRubro rs!anio, rs!Mes, rs!secuencia, rs!tipo_porc
'L_SeteoGrillaLoc
L_CargoLocal rs!anio, rs!Mes, rs!secuencia, rs!tipo_porc

L_ValoresInic
L_CargoValor

txtNom.Text = IIf(IsNull(rs!Descrip), "", rs!Descrip)

End Sub

Private Sub SetBotonesGeneral(valor As Boolean)
    
    Tollbar.Buttons(1).Enabled = valor
    Tollbar.Buttons(2).Enabled = Not valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = valor
    Tollbar.Buttons(6).Enabled = valor
    Tollbar.Buttons(7).Enabled = valor

    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor
'
'    TollBar.Buttons(12).Enabled = Not valor
    Tollbar.Buttons(16).Enabled = valor
'

'habilitar frames
    frCab.Enabled = Not valor
    Spread.spread_LockGrilla sprDia, valor, 2, sprDia.MaxCols - 1
    Spread.spread_LockGrilla sprLoc, valor, 2, sprLoc.MaxCols - 1
    If Not valor Then
        txtReg.Text = 0
        txtCantReg.Text = 0
    End If

End Sub

Private Function ArmarCondicion()
Dim Con$

Con$ = ""

If mskAnio.Text <> "" Then
    Con$ = Con$ & " and Anio = " & mskAnio.Text
End If

If mskMes.Text <> "" Then
    Con$ = Con$ & " and Mes = " & mskMes.Text
End If

If CboCodAeropuerto.Text <> "" Then
    Con$ = Con$ & " and cod_depn = '" & CboCodAeropuerto.Text & "'"
End If

If CboEspigon.Text <> "" Then
    Con$ = Con$ & " and cod_sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
End If

If txtNom.Text <> "" Then
    Con$ = Con$ & " and descrip like '" & txtNom.Text & "%'"
End If

If Con$ <> "" Then
    Con$ = " WHERE " & Mid(Con$, 5, Len(Con$))
End If

ArmarCondicion = Con$

End Function



Private Sub MePrepararMod()
If Aplicacion.username = rs!usuario Then
    SeteoBotonesMod False
Else
    MsgBox "No puede Modificar el modelo de " & rs!usuario, vbCritical + vbOKOnly, "ATENCION"
End If
End Sub

Private Function MeReconsultar() As Integer
Dim sql$
Dim i%
    
'frm_.caption = Aplicacion.SeteoProceso (frm_.caption)
    

If Aplicacion.ObtenerRsDAO(sqlGral$, rs) Then
        txtCantReg.Text = Aplicacion.CantReg(rs)
        If Val(txtReg.Text) > Val(txtCantReg.Text) Then
            txtReg.Text = txtCantReg.Text
        End If
        
        For i% = 1 To txtReg.Text - 1
            rs.MoveNext
        Next
        If txtCantReg.Text > 0 Then
            MellenarPantalla
        End If
        'MeSetearBotonesToolBar
        MeReconsultar = txtCantReg.Text
End If

'frm_.caption = Aplicacion.SeteoProceso (frm_.caption)

End Function
Private Sub MeSetearBotonesToolBar()
Dim i%
Dim but As Button

If txtCantReg.Text = 0 Then
'    TollBar.Buttons(1).Enabled = False
'    TollBar.Buttons(2).Enabled = False
'    TollBar.Buttons(3).Enabled = False
'    TollBar.Buttons(4).Enabled = False
'    TollBar.Buttons(6).Enabled = False
'    TollBar.Buttons(7).Enabled = False
ElseIf txtCantReg.Text = 1 Then
    Tollbar.Buttons(4).Enabled = False
    Tollbar.Buttons(5).Enabled = False
    Tollbar.Buttons(6).Enabled = False
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
ElseIf txtReg.Text = txtCantReg.Text Then
    Tollbar.Buttons(4).Enabled = True
    Tollbar.Buttons(5).Enabled = True
    Tollbar.Buttons(6).Enabled = False
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
ElseIf txtReg.Text = 1 Then
    Tollbar.Buttons(4).Enabled = False
    Tollbar.Buttons(5).Enabled = False
    Tollbar.Buttons(6).Enabled = True
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
Else
    Tollbar.Buttons(4).Enabled = True
    Tollbar.Buttons(5).Enabled = True
    Tollbar.Buttons(6).Enabled = True
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
    
End If
    


End Sub



Private Sub SetearBotonesAltas()
        
    Tollbar.Buttons(1).Enabled = True
    Tollbar.Buttons(11).Enabled = True
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Visible = False
    Tollbar.Buttons(14).Visible = False
    
    Tollbar.Buttons(2).Visible = False
    
    Tollbar.Buttons(4).Visible = False
    Tollbar.Buttons(5).Visible = False
    Tollbar.Buttons(6).Visible = False
    Tollbar.Buttons(7).Visible = False
    
    Tollbar.Buttons(9).Visible = False
    Tollbar.Buttons(10).Visible = False
    Tollbar.Buttons(16).Visible = False
        
    txtCantReg.Visible = False
    txtReg.Visible = False
    de.Visible = False

    frComi.Enabled = True
    frRub.Enabled = True
    'frDia.Enabled = True
    'frLoc.Enabled = True
    
Spread.spread_LockGrilla sprDia, False, 2, sprDia.MaxCols - 1
Spread.spread_LockGrilla sprLoc, False, 2, sprLoc.MaxCols - 1

End Sub

Private Sub SeteoBotonesMod(valor As Boolean)
    
    
    Tollbar.Buttons(1).Enabled = valor
    Tollbar.Buttons(2).Enabled = valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = valor
    Tollbar.Buttons(6).Enabled = valor
    Tollbar.Buttons(7).Enabled = valor

    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor

    Tollbar.Buttons(12).Enabled = Not valor
    Tollbar.Buttons(13).Enabled = Not valor
    Tollbar.Buttons(14).Enabled = valor

'habilitar o des frames y/o campos
    'frDia.Enabled = Not valor
    Spread.spread_LockGrilla sprDia, valor, 2, sprDia.MaxCols - 1
    If optTipo(0).Value Then
        frRub.Enabled = Not valor
        frComi.Enabled = Not valor
        'frLoc.Enabled = Not valor
        Spread.spread_LockGrilla sprLoc, valor, 2, sprLoc.MaxCols - 1
    End If
End Sub





Private Sub CboCodAeropuerto_Click()
Dim sql As String
Dim i As Integer
If Modo = "ALTA" Then
    chk.Value = 1
End If
sql = " SELECT cod_sdep,descrip FROM baires.subdependencia "
sql = sql & " WHERE cod_depn = '" & CboCodAeropuerto.Text & "'"
sql = sql & " ORDER BY cod_sdep"
 
FuncCbos_LlenarCboLst CboEspigon, LstEspigon, sql

CboEspigon.AddItem "Inter Llegadas"
LstEspigon.AddItem "INTAL"

CboEspigon.AddItem "Inter Salidas"
LstEspigon.AddItem "INTAS"

labValor.caption = ""
txtNom.Text = ""

For i = 2 To 6
    sprLoc.col = i
    sprLoc.ColHidden = True
Next

End Sub


Private Sub CboEspigon_Click()
Dim sql As String

If Modo = "ALTA" Then
    chk.Value = 1
    sql$ = ""
    sql$ = sql$ & "Select * From estadis.porciento_espigon "
    sql$ = sql$ & " Where anio = " & mskAnio.Text
    sql$ = sql$ & " And mes =" & mskMes.Text
    sql$ = sql$ & " And cod_depn ='" & CboCodAeropuerto.Text & "'"
    sql$ = sql$ & " And cod_sdep ='" & LstEspigon.List(CboEspigon.ListIndex) & "'"
            
    If CboEspigon.Text <> "" Then
        Call Aplicacion.ObtenerRsDAO(sql$, rsPorcEsp)
        L_PonerValores
        L_SeteoGrillaLoc
        txtNom.Text = L_NombreSujerido
    Else
        labValor.caption = ""
        txtNom.Text = ""
    End If
End If

End Sub

Private Sub Form_Activate()
If mskMes.Enabled Then
mskMes.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim sql As String

Top = 450
Left = 350
Height = 6600
Width = 9400

sql = " SELECT cod_depn,descrip FROM baires.dependencia "
sql = sql & " ORDER BY cod_depn"

FuncCbos_LlenarCbo CboCodAeropuerto, sql

Set cl_Estim = New CLEstimado

If Modo = "ALTA" Then
    mskAnio.Text = Year(Date)
    mskMes.Text = Month(Date)
    
    mskIosc.Text = 0
    mskNoIosc.Text = 0
End If

chk.Value = 0

End Sub

Private Sub L_SetearGrillas()
Dim sql As String
Dim rs As Recordset
Dim FD As String, FH As String
Dim fdAux As Date, i As Integer

sql = "select cod_rubr cant from baires.rubro order by cod_rubr "
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Spread_CargarGrilla rs, sprRubro
    Spread_CargarGrilla rs, sprLoc
    Aplicacion.CerrarDAO rs
End If
FD = func_Dia1SegunMes_Anio(mskMes.Text, mskAnio.Text)
FH = func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text)

fdAux = FD
i = 1
sprDia.MaxRows = 0
Do While fdAux <= CDate(FH)
        sprDia.MaxRows = sprDia.MaxRows + 1
        sprDia.SetText 1, i, Format$(fdAux, "dd-mm-yy")
        sprDia.SetText 2, i, "0"
        i = i + 1
        fdAux = fdAux + 1

Loop

Spread_PintarfinSemana sprDia

Spread_TotalesGrillas sprRubro, 1, 2
Spread_TotalesGrillas sprDia, 1, 2

End Sub






Private Sub mskAnio_Change()
Dim i

For i = 2 To 6
    sprLoc.col = i
    sprLoc.ColHidden = True
Next

sprDia.MaxRows = 0
sprRubro.MaxRows = 0
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub mskAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub mskAnio_LostFocus()

If Modo = "ALTA" Then
        If Val(mskAnio.Text) < 1996 Or Val(mskAnio) > 2050 Then
            mskAnio.Text = Year(Date)
        End If
    'Control de modelos iniciales cargados
    If L_SecuenciaControl Then
        If sprDia.MaxRows = 0 Then
            L_SetearGrillas
        End If
    Else
        MsgBox "No hay datos Iniciales cargados para este A?o-Mes", vbExclamation + vbOKOnly, "ATENCION"
        CboCodAeropuerto.ListIndex = -1
    End If
End If
End Sub

Private Sub mskIosc_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub mskIosc_LostFocus()
If Not IsNumeric(mskIosc.Text) Then
    mskIosc.Text = 0
End If

End Sub


Private Sub mskMes_Change()
Dim i

For i = 2 To 6
    sprLoc.col = i
    sprLoc.ColHidden = True
Next

sprDia.MaxRows = 0
sprRubro.MaxRows = 0
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub


Private Sub mskMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub mskMes_LostFocus()
If Modo = "ALTA" Then
    If Val(mskMes.Text) < 1 Or Val(mskMes.Text) > 12 Then
        mskMes.Text = Month(Date)
    End If
    If L_SecuenciaControl Then
        If sprDia.MaxRows = 0 Then
            L_SetearGrillas
        End If
    Else
        MsgBox "No hay datos Iniciales cargados para este A?o-Mes", vbExclamation + vbOKOnly, "ATENCION"
        CboCodAeropuerto.ListIndex = -1
    End If
End If
End Sub


Private Sub mskNoIosc_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub mskNoIosc_LostFocus()
If Not IsNumeric(mskNoIosc.Text) Then
    mskNoIosc.Text = 0
End If
End Sub


Private Sub OptTipo_Click(Index As Integer)
Dim sql As String
Dim rs As Recordset
Dim i, j

On Error GoTo ErrT:

If Modo = "ALTA" Then
'Aplicacion.SeteoProceso ("")

sql = "SELECT Estadis.Sec_Modelo ("
sql = sql & mskAnio.Text & ","
sql = sql & mskMes.Text & ","
sql = sql & "'" & L_SegunOptTipo & "') sec from dual "

If Aplicacion.ObtenerRsDAO(sql, rs) Then
    If rs!sec > 0 Then
        txtSec.Text = rs!sec
        L_PonerValores
        txtNom.Text = L_NombreSujerido
    Else
        txtSec.Text = ""
        labValor.caption = ""
    End If
    Aplicacion.CerrarDAO rs
End If
Select Case Index
    Case 0
        frComi.Enabled = True
        frRub.Enabled = True
        frLoc.Enabled = True
    Case 1, 2, 3
        frComi.Enabled = False
        frRub.Enabled = False
        frLoc.Enabled = False
        mskIosc.Text = 0
        mskNoIosc = 0
        For i = 1 To sprRubro.MaxRows - 1
            sprRubro.SetText 2, i, "0"
        Next
        For j = 2 To ultCol - 1
            For i = 1 To sprLoc.MaxRows
                sprLoc.SetText j, i, "0"
            Next
        Next
End Select
End If
ErrT:
 '   Aplicacion.SeteoFin
    Exit Sub
End Sub

Private Sub sprDia_Change(ByVal col As Long, ByVal Row As Long)
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub


Private Sub sprRubro_Change(ByVal col As Long, ByVal Row As Long)
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub sprRubro_Click(ByVal col As Long, ByVal Row As Long)
Dim a

a = 1
End Sub

Private Sub TollBar_ButtonClick(ByVal Button As ComctlLib.Button)
Dim a%
Dim pos As String
Dim saltear As Boolean

saltear = True

pos = txtReg.Text

Select Case Button.Key
    Case "a"
        saltear = False
        Func_MoverPrimero rs, pos
    Case "b"
        saltear = False
        Func_MoverAnterior rs, pos
    Case "c"
        saltear = False
        Func_MoverSiguiente rs, pos
    Case "e"
        saltear = False
        Func_MoverUltimo rs, pos
    Case "f"
         MePrepararMod
    Case "g"
         MeEliminar
    Case "h"
        If Modo = "MOD" Then
            MeActualizar
        Else
            L_AltasDatos
        End If
    Case "i"
        MeAbortarMod
    Case "j"
        If chk.Value Then
            If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
                If Modo = "MOD" Then
                    MeActualizar
                Else
                    L_AltasDatos
                End If
            End If
        End If
        Unload Me
    Case "k"
        NuevaSeleccion
    Case "l"
        MeCargarDatos
    Case "m"
        L_ImprimirCR
    Case "o"
        L_CopiarModelo
    
End Select

If Not saltear Then
    txtReg.Text = pos
    MellenarPantalla
    MeSetearBotonesToolBar
End If

End Sub
Private Sub txtNom_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub


