VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.0#0"; "crystl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmAdmPxC 
   Caption         =   "Administración de Proveedores a Concursos"
   ClientHeight    =   5730
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   6015
   Icon            =   "FrmAdmPxC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5730
   ScaleWidth      =   6015
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   20
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "o"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "p"
            Object.ToolTipText     =   "Estado Mod / Consulta"
            Object.Tag             =   ""
            ImageIndex      =   14
            Value           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "k"
            Object.ToolTipText     =   "Nueva Seleción"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "l"
            Object.ToolTipText     =   "Buscar"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "a"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "b"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "c"
            Object.ToolTipText     =   "Registro Siguiente"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "e"
            Object.ToolTipText     =   "Ultimo Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "f"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "g"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "h"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "i"
            Object.ToolTipText     =   "Abortar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "m"
            Object.ToolTipText     =   "imprimir"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "Grilla"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "j"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin Crystal.CrystalReport rptControl 
      Left            =   5625
      Top             =   4065
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327682
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame frConf 
      Height          =   1770
      Left            =   735
      TabIndex        =   26
      Top             =   3645
      Visible         =   0   'False
      Width           =   4680
      Begin VB.TextBox cod 
         Height          =   285
         Left            =   150
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   915
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton botSI 
         Caption         =   "SI"
         Height          =   315
         Left            =   735
         TabIndex        =   30
         Top             =   1335
         Width           =   1140
      End
      Begin VB.CommandButton botNO 
         Caption         =   "NO"
         Height          =   315
         Left            =   2700
         TabIndex        =   28
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   255
         TabIndex        =   29
         Top             =   270
         Width           =   4095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Confirma carga general de productos ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   315
         TabIndex        =   27
         Top             =   630
         Width           =   4095
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "chk"
      Height          =   195
      Left            =   1380
      TabIndex        =   14
      Top             =   450
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Height          =   2940
      Left            =   210
      TabIndex        =   12
      Top             =   2700
      Width           =   5625
      Begin FPSpread.vaSpread spr 
         Height          =   2295
         Left            =   105
         OleObjectBlob   =   "FrmAdmPxC.frx":0442
         TabIndex        =   4
         Tag             =   "NOSALTAR"
         Top             =   570
         Width           =   5430
      End
      Begin VB.Frame frBot 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   435
         Left            =   195
         TabIndex        =   18
         Top             =   120
         Width           =   5055
         Begin ComctlLib.Toolbar Toolbar1 
            Height          =   420
            Left            =   90
            TabIndex        =   19
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   741
            ButtonWidth     =   635
            ButtonHeight    =   582
            Appearance      =   1
            ImageList       =   "ImageList1"
            _Version        =   327682
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   7
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "A"
                  Object.ToolTipText     =   "Agreagar Fila"
                  Object.Tag             =   ""
                  ImageIndex      =   16
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   ""
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "B"
                  Object.ToolTipText     =   "Sacar Fila"
                  Object.Tag             =   ""
                  ImageIndex      =   17
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   ""
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "C"
                  Object.ToolTipText     =   "Limpiar Todo"
                  Object.Tag             =   ""
                  ImageIndex      =   9
               EndProperty
               BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Visible         =   0   'False
                  Key             =   ""
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Visible         =   0   'False
                  Key             =   "D"
                  Object.ToolTipText     =   "Salir"
                  Object.Tag             =   ""
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.CommandButton botHelpProd 
            Height          =   300
            Left            =   4560
            Picture         =   "FrmAdmPxC.frx":0763
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Códigos de Productos"
            Top             =   75
            Width           =   390
         End
      End
   End
   Begin VB.TextBox txtReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4185
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   420
      Width           =   465
   End
   Begin VB.TextBox txtCantReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5100
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   420
      Width           =   480
   End
   Begin VB.Frame frCab 
      Height          =   2145
      Left            =   195
      TabIndex        =   9
      Top             =   525
      Width           =   5640
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   4515
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1185
         Width           =   960
      End
      Begin MSMask.MaskEdBox mskMes 
         Height          =   315
         Left            =   3840
         TabIndex        =   23
         Top             =   315
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.CheckBox chkActivo 
         Caption         =   "ACTIVO"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3825
         TabIndex        =   21
         Top             =   1650
         Width           =   1335
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   3180
         Picture         =   "FrmAdmPxC.frx":0865
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1575
         Width           =   375
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   3165
         Picture         =   "FrmAdmPxC.frx":09D7
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1170
         Width           =   375
      End
      Begin VB.TextBox txtCli 
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   1785
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         Top             =   330
         Width           =   1380
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1785
         TabIndex        =   2
         Top             =   1185
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFHasta 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   1605
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtDescrip 
         Height          =   315
         Left            =   1785
         TabIndex        =   1
         Top             =   735
         Width           =   3555
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   315
         Left            =   4800
         TabIndex        =   33
         Top             =   330
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
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
         Height          =   315
         Index           =   6
         Left            =   4305
         TabIndex        =   32
         Top             =   330
         Width           =   450
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rubro"
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
         Left            =   3795
         TabIndex        =   24
         Top             =   1185
         Width           =   660
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
         Height          =   315
         Index           =   2
         Left            =   3255
         TabIndex        =   22
         Top             =   330
         Width           =   540
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
         Index           =   3
         Left            =   495
         TabIndex        =   17
         Top             =   1620
         Width           =   1260
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
         Index           =   4
         Left            =   495
         TabIndex        =   16
         Top             =   1185
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   510
         TabIndex        =   15
         Top             =   735
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Concurso"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   510
         TabIndex        =   11
         Top             =   315
         Width           =   1230
      End
      Begin VB.Label de 
         Caption         =   "de"
         Height          =   255
         Left            =   4575
         TabIndex        =   10
         Top             =   120
         Width           =   405
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":0B49
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":0E63
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":117D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":1497
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":17B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":1ACB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":1DE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":20FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":2419
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":2733
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":2845
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":2957
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":2EF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":372B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":3EE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":47AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPxC.frx":4AC5
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAdmPxC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As Recordset

Dim cl_PxM As CLConc

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$

Dim ModoEdit As Integer
Dim DatoValido As Boolean

Dim HelpP As Boolean

Dim ModGrilla As Boolean

Public Sub Altas()
    SetearBotonesAltas True
    Modo = "ALTA"
    FrmAdmPxC.caption = "Administración de Concurso/Proveedor - ALTAS -"
    
'    Me.Show 1
End Sub
Private Sub L_AltasDatos()

If L_TodoCargado Then
    FrmAdmPxC.caption = Aplicacion.SeteoProceso(FrmAdmPxC.caption)

    Aplicacion.ComienzoTrans

    MeLlenarObjeto

    If cl_PxM.Insert_PxC() Then
        Aplicacion.TerminarConExitoTrans
        chk.Value = 0
       
        NuevaSeleccion
    
    Else
        Aplicacion.TerminarConErrorTrans
    End If

    FrmAdmPxC.caption = Aplicacion.SeteoFin
Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub L_llenarProveedor()
Dim sql As String
Dim rsP As Recordset

sql = "SELECT P.cod_prov,P.descrip,cp.secuencia " _
& " FROM baires.proveedor P,estadis.concurso_prov CP " _
& " WHERE  P.cod_prov = CP.cod_prov " _
& " and CP.id_concurso = '" & txtCli.Text & "' " _
& " UNION " _
& " SELECT CP.cod_prov, 'DISCONTINUADOS' AS Descrip,cp.secuencia  " _
& " FROM estadis.concurso_prov CP " _
& " WHERE  CP.cod_prov ='N' " _
& " and CP.id_concurso = '" & txtCli.Text & "' "

spr.MaxRows = 0
If Aplicacion.ObtenerRsDAO(sql, rsP) Then

    Do While Not rsP.EOF
        spr.MaxRows = spr.MaxRows + 1
        
        spr.SetText 1, spr.MaxRows, Trim(rsP!cod_prov)
        spr.SetText 2, spr.MaxRows, Trim(rsP!Descrip)
        spr.SetText 3, spr.MaxRows, str(rsP!secuencia)
        rsP.MoveNext
    Loop
    Aplicacion.CerrarDAO rsP
End If


End Sub

Private Sub MeImpDatos()
Dim nom As String, NombreArchivo As String


' On Error GoTo ErrFoto:

Aplicacion.SeteoProceso ("fromadmpxp.caption")

rptControl.Connect = "ODBC;UID=" & Aplicacion.username & ";PWD=" & Aplicacion.password & ";DATABASE=INTER;"

rptControl.PrinterCopies = 1
rptControl.Destination = 0

'rptform.PrinterStopPage = 0
'rptform.PrinterStartPage = 0

rptControl.ReportFileName = RutaReportes
rptControl.ReportFileName = RutaReportes & "controlconcurso.rpt"

rptControl.ParameterFields(0) = "Anio;" & mskAnio.Text & ";true"
rptControl.ParameterFields(1) = "Mes;" & mskMes.Text & ";true"
'rptform.ParameterFields(0) = "PrActFechaDesde;" & " DATE(" & Year(mskFDesde.FormattedText) & "," & Month(mskFDesde.FormattedText) & "," & Day(mskFDesde.FormattedText) & ")" & ";true)"

rptControl.Action = True
DoEvents

ErrFoto:
    Aplicacion.SeteoFin
    Exit Sub
        
End Sub

Public Sub Modificacion()
Dim sql As String

SetearBotonesAltas False
Modo = "MOD"
FrmAdmPxC.caption = "Administració de Concurso/Proveedor - Modificacion y Bajas -"

'Me.Show 1
End Sub
Private Sub NuevaSeleccion()
Dim i%

If Modo = "MOD" Then
    SetBotonesGeneral False
Else
    If chk.Value = 1 Then
        If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
            L_AltasDatos
        End If
        
    End If
End If
'Limpiar campos de pantallas
Set cl_PxM = New CLConc

    spr.MaxRows = 0
    txtDescrip.Text = ""
    txtCli.Text = ""
    mskFDesde.Text = ""
    mskFHasta.Text = ""
    chkActivo.Value = 0
    mskMes.Text = ""
    mskAnio.Text = ""
    cboRubro.ListIndex = -1
    
    chk.Value = 0

End Sub

Private Sub MeAbortarMod()
    
    SeteoBotonesMod True
    
    Tollbar.Buttons(2).Enabled = False
    
    MeSetearBotonesToolBar
    
    MellenarPantalla
    
End Sub

Private Sub MeActualizar()
Dim ViejoOrgan$
Dim Viejocargo%

If L_TodoCargado Then

FrmAdmPxC.caption = Aplicacion.SeteoProceso(FrmAdmPxC.caption)

Aplicacion.ComienzoTrans

MeLlenarObjeto


    If cl_PxM.Update_PxC(ModGrilla) Then
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


FrmAdmPxC.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub MeCargarDatos()
Dim sql$

FrmAdmPxC.caption = Aplicacion.SeteoProceso(FrmAdmPxC.caption)
        
    CondConsulta = ArmarCondicion

    sqlGral$ = ""
    sqlGral$ = sqlGral$ & " SELECT ID_CONCURSO ,"
    sqlGral$ = sqlGral$ & " DESCRIP, "
    sqlGral$ = sqlGral$ & " FCH_VDESDE, "
    sqlGral$ = sqlGral$ & " FCH_VHASTA, "
    sqlGral$ = sqlGral$ & " Activo, "
    sqlGral$ = sqlGral$ & " Mes, "
    sqlGral$ = sqlGral$ & " cod_rubr "
    sqlGral$ = sqlGral$ & " FROM estadis.CONCURSO_H "
    sqlGral$ = sqlGral$ & CondConsulta
    
If Aplicacion.ObtenerRsDAO(sqlGral$, rs) Then
    txtCantReg.Text = Aplicacion.CantReg(rs)
    If txtCantReg.Text > 0 Then
        txtReg.Text = 1
        SetBotonesGeneral True
        MellenarPantalla
        MeSetearBotonesToolBar
    Else
        txtReg.Text = 0
    End If
End If

FrmAdmPxC.caption = Aplicacion.SeteoFin

End Sub
Private Sub MeEliminar()

If MsgBox("Esta seguro de eliminar el registro", vbYesNo + vbExclamation, "ATENCION") = vbYes Then

'MeLlenarObjeto
cl_PxM.codConc = txtCli.Text

FrmAdmPxC.caption = Aplicacion.SeteoProceso(FrmAdmPxC.caption)

Aplicacion.ComienzoTrans

If cl_PxM.Delete_PxC Then
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

FrmAdmPxC.caption = Aplicacion.SeteoFin

End If

End Sub

Private Sub MeLlenarObjeto()
'Dim codigo As String
Dim concprov As CLProdPrec
Dim i As Long, valor As Variant

cl_PxM.Mes = mskMes.Text
cl_PxM.codConc = txtCli.Text
cl_PxM.Descrip = txtDescrip.Text
cl_PxM.fchDesde = mskFDesde.FormattedText
cl_PxM.fchHasta = mskFHasta.FormattedText
cl_PxM.CodRubr = cboRubro.Text

If chkActivo.Value = 1 Then
    cl_PxM.activo = "S"
Else
    cl_PxM.activo = "N"
End If

Set cl_PxM.col_producto = New Collection

For i = 1 To spr.MaxRows

    spr.GetText 1, i, valor
    If valor <> "" And Spread_FilaOcupada(spr, i) Then
        Set concprov = New CLProdPrec
        concprov.CodProv = valor

        spr.GetText 3, i, valor
        concprov.sec = valor
        AdicionarAColeccion concprov
    End If

Next

End Sub


Private Sub AdicionarAColeccion(CP As CLProdPrec)
'(cl As CLProdMon)
Dim concprov As CLProdPrec
Dim Resp As Boolean

Resp = True

For Each concprov In cl_PxM.col_producto
    If concprov.CodProv = CP.CodProv And concprov.sec = CP.sec Then
        Resp = False
        Exit For
    End If
Next

If Resp Then
    cl_PxM.col_producto.Add CP
End If

End Sub


Private Function L_TodoCargado() As Boolean
    L_TodoCargado = False
    
    If txtCli.Text <> "" And txtDescrip.Text <> "" _
    And mskFDesde.Text <> "" And mskFHasta.Text <> "" And mskMes.Text <> "" _
    And cboRubro.Text <> "" Then
        L_TodoCargado = True
    Else
        L_TodoCargado = False
    End If
    
End Function

Private Sub MellenarPantalla()
Dim desc As String


txtCli.Text = rs!id_concurso
txtDescrip.Text = rs!Descrip

If rs!activo = "S" Then
    chkActivo.Value = 1
Else
    chkActivo.Value = 0
End If

mskFDesde.Text = Format$(rs!fch_vdesde, FTOFECHA)
mskFHasta.Text = Format$(rs!fch_vhasta, FTOFECHA)

mskMes.Text = rs!Mes

Func.Func_SetearCboSTR cboRubro, rs!COD_RUBR
L_llenarProveedor

End Sub

Public Sub PonerValores(cod As Variant, desc As String)
    spr.SetText 1, spr.MaxRows, Trim(cod)
    spr.SetText 2, spr.MaxRows, desc
'    If spr.ActiveRow = spr.MaxRows Then
       Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
'    End If
    If spr.MaxRows > 6 Then
        spr.Row = spr.MaxRows - 4
    Else
        spr.Row = spr.MaxRows
    End If
    spr.col = 1
    spr.Position = SS_POSITION_UPPER_LEFT
    spr.Action = SS_ACTION_GOTO_CELL

DatoValido = True
ModGrilla = True
End Sub

Private Sub SetBotonesGeneral(valor As Boolean)
    
    Tollbar.Buttons(1).Enabled = Not valor
    Tollbar.Buttons(2).Enabled = Not valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = Not valor
    
    Tollbar.Buttons(7).Enabled = valor
    Tollbar.Buttons(8).Enabled = valor
    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor

    Tollbar.Buttons(12).Enabled = valor
    Tollbar.Buttons(13).Enabled = valor
'
    Tollbar.Buttons(18).Enabled = valor
'    TollBar.Buttons(13).Enabled = Not valor
'
    Spread.spread_LockGrilla spr, valor, 1, spr.MaxCols
    frCab.Enabled = Not valor
'habilitar frames

    If Not valor Then
        txtReg.Text = 0
        txtCantReg.Text = 0
    End If

End Sub


Private Function ArmarCondicion()
Dim Con$

Con$ = ""

If txtCli.Text <> "" Then
    Con$ = Con$ & " and Id_concurso = '" & txtCli.Text & "' "
End If
If txtDescrip.Text <> "" Then
    Con$ = Con$ & " and descrip like '" & txtDescrip.Text & "%' "
End If
If mskMes.Text <> "" Then
    Con$ = Con$ & " and mes = " & mskMes.Text & " "
End If
If mskAnio.Text <> "" Then
    Con$ = Con$ & " and anio_alta = " & mskAnio.Text & " "
End If
If cboRubro.Text <> "" Then
    Con$ = Con$ & " and cod_rubr = '" & cboRubro.Text & "' "
End If
If Con$ <> "" Then
    Con$ = " WHERE " & Mid(Con$, 5, Len(Con$))
End If

ArmarCondicion = Con$

End Function



Private Sub MePrepararMod()
    
    SeteoBotonesMod False
    DatoValido = True
    ModGrilla = False
    
End Sub

Private Function MeReconsultar() As Integer
Dim sql$
Dim i%
    

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
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(8).Enabled = False
    Tollbar.Buttons(9).Enabled = False
    Tollbar.Buttons(10).Enabled = False
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
ElseIf txtReg.Text = txtCantReg.Text Then
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(8).Enabled = True
    Tollbar.Buttons(9).Enabled = False
    Tollbar.Buttons(10).Enabled = False
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
ElseIf txtReg.Text = 1 Then
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(8).Enabled = False
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
Else
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(8).Enabled = True
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
    
End If
    


End Sub



Private Sub SetearBotonesAltas(valor As Boolean)
'valor = true -> altas
'valor = false -> modif
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(15).Enabled = valor
    Tollbar.Buttons(16).Visible = Not valor
    
    Tollbar.Buttons(17).Visible = Not valor 'False
    Tollbar.Buttons(18).Visible = Not valor 'False
    
    Tollbar.Buttons(5).Visible = Not valor 'False
    
    Tollbar.Buttons(7).Visible = Not valor 'False
    Tollbar.Buttons(8).Visible = Not valor 'False
    Tollbar.Buttons(9).Visible = Not valor 'False
    Tollbar.Buttons(10).Visible = Not valor 'False
    
    Tollbar.Buttons(12).Visible = Not valor 'False
    Tollbar.Buttons(13).Visible = Not valor 'False

    Tollbar.Buttons(18).Visible = Not valor 'False
    Tollbar.Buttons(19).Visible = Not valor 'False
    
    txtCantReg.Visible = Not valor 'False
    txtReg.Visible = Not valor 'False
    de.Visible = Not valor 'False
    
    Spread.spread_LockGrilla spr, Not valor, 1, spr.MaxCols
    txtCli.Locked = Not valor
    frBot.Enabled = valor
    
End Sub


Private Sub SeteoBotonesMod(valor As Boolean)
    
    
    Tollbar.Buttons(4).Enabled = valor
    'Tollbar.Buttons(5).Enabled = valor
    
    Tollbar.Buttons(7).Enabled = valor
    Tollbar.Buttons(8).Enabled = valor
    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor

    Tollbar.Buttons(12).Enabled = valor
    Tollbar.Buttons(13).Enabled = valor

    Tollbar.Buttons(15).Enabled = Not valor
    Tollbar.Buttons(16).Enabled = Not valor

    Tollbar.Buttons(18).Enabled = valor
    Tollbar.Buttons(19).Enabled = valor
'habilitar o des frames y/o campos
    
    Spread.spread_LockGrilla spr, valor, 1, spr.MaxCols
    frBot.Enabled = Not valor
    frCab.Enabled = Not valor
    
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


Private Sub botHelpProd_Click()
Dim cod As String
Dim desc As String
Dim sql As String


If spr.MaxRows > 0 Then
sql = "Select cod_prov,descrip from baires.proveedor "
If frmHelpProd.MuestraHlp(cod, desc, "Proveedor", sql, "") = vbOK Then
       spr.SetText 1, spr.ActiveRow, cod
       spr.SetText 2, spr.ActiveRow, desc
       'If spr.ActiveRow = spr.MaxRows Then
       '   Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
       'End If
       DatoValido = True
End If
End If

spr.SetFocus
End Sub












Private Sub botNO_Click()
frConf.Visible = False
End Sub

Private Sub botSI_Click()
Dim sql As String
Dim sqlD As String

sqlD = sqlD & " delete "
sqlD = sqlD & " FROM estadis.concurso_d c"
sqlD = sqlD & " where id_concurso = '" & txtCli.Text & "' "
sqlD = sqlD & " And c.cod_prov = '" & cod.Text & "' "

sql = "Insert into ESTADIS.CONCURSO_D "
sql = sql & " (ID_CONCURSO,COD_PROV,COD_PROD,PAGA,SECUENCIA)"
sql = sql & " SELECT ID_CONCURSO,p.cod_prov,cod_prod,0,SECUENCIA"
sql = sql & " FROM ESTADIS.concurso_prov c, baires.producto p"
sql = sql & " where id_concurso = '" & txtCli.Text & "' And p.cod_rubr <> 'REG' "
sql = sql & " and c.cod_prov = p.cod_prov And c.cod_prov = '" & cod.Text & "' "

Aplicacion.ComienzoTrans

If Aplicacion.EjecutarDAO(sqlD) Then
   If Aplicacion.EjecutarDAO(sql) Then
      Aplicacion.TerminarConExitoTrans
      MsgBox "La actualización termino correctamente", vbOKOnly, "Atencion"
      frConf.Visible = False
   Else
      Aplicacion.TerminarConErrorTrans
   End If
End If

End Sub

Private Sub cboRubro_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    cboRubro.ListIndex = -1
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If Shift = 4 Then
    Select Case KeyCode
        Case 78 'Nueva sel
            Call TollBar_ButtonClick(Tollbar.Buttons(4))
        Case 71 'Guardar
            If Tollbar.Buttons(12).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(15))
            End If
        Case 66 'Buscar
            If Modo = "MOD" Then
            Call TollBar_ButtonClick(Tollbar.Buttons(5))
            End If
        Case 83 'Salir
            Call TollBar_ButtonClick(Tollbar.Buttons(20))
        Case 107
            Call Toolbar1_ButtonClick(Toolbar1.Buttons(4))
        Case 109
            Call Toolbar1_ButtonClick(Toolbar1.Buttons(6))
    
    End Select
    If Modo = "MOD" And Val(txtCantReg.Text) > 0 Then
    Select Case KeyCode
        Case 37 'Izq
            If Tollbar.Buttons(8).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(5))
            End If
        Case 38 'Arriba
            Call TollBar_ButtonClick(Tollbar.Buttons(10))
        Case 40 'Abajo
            Call TollBar_ButtonClick(Tollbar.Buttons(7))
        Case 39 'Der
            If Tollbar.Buttons(6).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(9))
            End If
    End Select
    End If
Else
'    Select Case KeyCode
'        Case 107
'            Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
'        Case 109
'            Call Toolbar1_ButtonClick(Toolbar1.Buttons(3))
 '   End Select

End If
'Debug.Print KeyCode
End Sub

Private Sub MePrepararAlterar()

    Tollbar.Buttons(2).Value = tbrPressed
    Tollbar.Buttons(1).Value = tbrUnpressed
    
    Modificacion
    
    NuevaSeleccion
    
End Sub

Private Sub MePrepararAgregar()

    Tollbar.Buttons(1).Value = tbrPressed
    Tollbar.Buttons(2).Value = tbrUnpressed
    
    Altas
    
    NuevaSeleccion
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Tag = "" Then
        SendKeys "{TAB}"
    End If
End If
End Sub


Private Sub Form_Load()
Dim sql As String

Top = 800
Left = 1000
Width = 6200
Height = 6100

Set cl_PxM = New CLConc

HelpP = False
Modo = "MOD"
    
If Aplicacion.Nivel <> 0 Then
    chkActivo.Enabled = False
End If

sql = "SELECT cod_rubr FROM baires.rubro Union select 'OTR' from dual "

FuncCbos.FuncCbos_LlenarCbo cboRubro, sql


cboRubro.AddItem "DIS"

End Sub


Private Sub mskFDesde_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub mskFDesde_LostFocus()

If mskFDesde.Text <> "" Then
    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    End If
    
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
Else
    mskFHasta.Text = ""
End If
End Sub


Private Sub mskFHasta_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub mskFHasta_LostFocus()

If mskFDesde.Text <> "" Then
    If mskFHasta.Text <> "" Then
        If Not IsDate(mskFHasta.FormattedText) Then
                mskFHasta.Text = mskFDesde.Text
        End If
        
        mskFHasta.Text = Format$(mskFHasta.FormattedText, FTOFECHA)
        If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
            mskFHasta.Text = mskFDesde.FormattedText
        End If
    End If
Else
    mskFHasta.Text = ""
End If
End Sub


Private Sub mskMes_LostFocus()
If mskMes.Text <> "" Then
If Val(mskMes.Text) > 12 Or Val(mskMes.Text) < 1 Then
    mskMes.Text = Month(Now)
End If
End If
End Sub


Private Sub spr_DblClick(ByVal col As Long, ByVal Row As Long)
Dim desc As Variant
Dim vCod As Variant

spr.GetText 2, Row, desc
spr.GetText 1, Row, vCod


frConf.Visible = True
Label3.caption = desc
cod.Text = vCod

End Sub

Private Sub spr_EditMode(ByVal col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If Mode = 1 Then
    ModoEdit = Mode
End If
End Sub


Private Sub spr_GotFocus()
FrmAdmPxC.Tag = "NOSALTAR"
End Sub


Private Sub spr_KeyPress(KeyAscii As Integer)
Dim sql As String
Dim desc As String
Dim cod As Variant
    
If KeyAscii = 13 And ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    spr.GetText 1, spr.ActiveRow, cod
    
    sql = "SELECT descrIP FROM BAIRES.Proveedor where cod_proV = '" & cod & "'"
    DatoValido = Func_ObtenerDesc(sql, desc)
    Select Case spr.ActiveCol
        Case 1
            spr.SetText 2, spr.ActiveRow, desc
            If DatoValido Then
                spr.col = 3
                spr.Row = spr.ActiveRow
                spr.Action = 0
                spr.Action = 1
            End If
        Case 3
            If DatoValido And spr.ActiveRow = spr.MaxRows Then
               Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
            End If
            spr.col = 1
            'If spr.MaxRows > 6 Then
            '    spr.Row = spr.MaxRows - 4
            'Else
                spr.Row = spr.MaxRows
            'End If
            spr.Action = 0
            spr.Action = 1
            
            
     End Select
     spr.TopRow = spr.TopRow - 6
    ModGrilla = True
End If

End Sub


Private Sub spr_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim sql As String
Dim desc As String
Dim cod As Variant


If ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    spr.GetText 1, Row, cod
    
    sql = "SELECT descrip FROM BAIRES.Proveedor where cod_proV = '" & cod & "'"
        
    DatoValido = Func_ObtenerDesc(sql, desc)
    
    spr.SetText 2, Row, desc
    ModGrilla = True
End If

Cancel = Not DatoValido


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
    Case "d"
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
        If chk.Value = 1 Then
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
    Case "n"
'        L_DatosGrilla
    
    Case "m"
        MeImpDatos
    Case "o"
        MePrepararAgregar
    Case "p"
        MePrepararAlterar
End Select

If Not saltear Then
    txtReg.Text = pos
    MellenarPantalla
    MeSetearBotonesToolBar
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim valor As Variant

Select Case Button.Key
    Case "A"
        If Spread_FilaOcupada(spr, spr.MaxRows) Then
           Spread_AddRow spr
        End If
    Case "B"
        Spread_DelOneRow spr, spr.ActiveRow
        DatoValido = True
        ModGrilla = True
    Case "C"
        spr.MaxRows = 0
        ModGrilla = True
    Case "D"
'        LlenarColeccion
'        Unload Me
        
End Select

End Sub


Private Sub txtCli_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub txtCli_GotFocus()
'spr.SetFocus
End Sub



Private Sub txtCli_LostFocus()
txtCli.Text = UCase(txtCli.Text)
End Sub

Private Sub txtDescrip_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub


