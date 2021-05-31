VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.0#0"; "crystl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmConsultaTicket 
   Caption         =   "Administración de Consulta de Tickets de Venta de DutyFree"
   ClientHeight    =   6555
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   9495
   Icon            =   "FrmConsultaTicket.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6555
   ScaleWidth      =   9495
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   165
      TabIndex        =   0
      Top             =   30
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "k"
            Object.ToolTipText     =   "Nueva Seleción"
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
            Key             =   "d"
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
            Key             =   "t"
            Object.ToolTipText     =   "Buscar Código"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "m"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "n"
            Object.ToolTipText     =   "Grilla"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "j"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin Crystal.CrystalReport RptTicket 
      Left            =   6195
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327682
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   5325
      Left            =   150
      TabIndex        =   2
      Top             =   570
      Width           =   8460
      Begin VB.Frame Frame3 
         Height          =   510
         Left            =   210
         TabIndex        =   55
         Top             =   120
         Width           =   3255
         Begin VB.OptionButton OptOnline 
            Caption         =   "Del día"
            Height          =   195
            Index           =   1
            Left            =   2010
            TabIndex        =   57
            Top             =   240
            Width           =   930
         End
         Begin VB.OptionButton OptOnline 
            Caption         =   "Estadístico"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   56
            Top             =   225
            Width           =   1605
         End
      End
      Begin VB.CommandButton CmdDetalle 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle del Ticket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   6015
         Picture         =   "FrmConsultaTicket.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Frame FraFecha 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   720
         Left            =   180
         TabIndex        =   9
         Top             =   660
         Width           =   8100
         Begin VB.CommandButton CmdFechaHasta 
            Height          =   390
            Left            =   7500
            Picture         =   "FrmConsultaTicket.frx":1564
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   225
            Width           =   465
         End
         Begin VB.CommandButton CmdFechaDesde 
            Height          =   390
            Left            =   3435
            Picture         =   "FrmConsultaTicket.frx":16EE
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   210
            Width           =   465
         End
         Begin MSMask.MaskEdBox MskFechaDesde 
            Height          =   315
            Left            =   1845
            TabIndex        =   11
            Top             =   255
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd-mm-yyyy"
            Mask            =   "##-##-####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskFchHasta 
            Height          =   315
            Left            =   5895
            TabIndex        =   14
            Top             =   270
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd-mm-yyyy"
            Mask            =   "##-##-####"
            PromptChar      =   " "
         End
         Begin VB.Label LblFchHasta 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha hasta :"
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
            Left            =   4170
            TabIndex        =   15
            Top             =   270
            Width           =   1605
         End
         Begin VB.Label LblFechaDesde 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha desde :"
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
            Left            =   105
            TabIndex        =   12
            Top             =   255
            Width           =   1605
         End
      End
      Begin VB.Frame FramCab 
         Caption         =   "Datos de Cabecera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2850
         Left            =   180
         TabIndex        =   6
         Top             =   1395
         Width           =   8115
         Begin VB.TextBox TxtTurno 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   7680
            TabIndex        =   52
            Top             =   1920
            Width           =   270
         End
         Begin VB.Frame Frame2 
            Height          =   1110
            Left            =   4080
            TabIndex        =   40
            Top             =   240
            Width           =   3885
            Begin MSMask.MaskEdBox txtCredit 
               Height          =   300
               Left            =   1890
               TabIndex        =   54
               Top             =   375
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.CommandButton CmdNext 
               Height          =   285
               Left            =   1440
               Picture         =   "FrmConsultaTicket.frx":1878
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   765
               Width           =   525
            End
            Begin VB.CommandButton CmdPrev 
               Height          =   285
               Left            =   855
               Picture         =   "FrmConsultaTicket.frx":1D7A
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   780
               Width           =   465
            End
            Begin VB.TextBox TxtTarj 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   2355
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   -30
               Width           =   465
            End
            Begin VB.TextBox TxtCantTarj 
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   -30
               Width           =   480
            End
            Begin VB.Label LblTarjeta 
               Caption         =   "Mastercard"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   2340
               TabIndex        =   47
               Top             =   675
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label LblCredit 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Tarjeta de Crédito :"
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
               Left            =   75
               TabIndex        =   44
               Top             =   360
               Width           =   1755
            End
            Begin VB.Label Label2 
               Caption         =   "de"
               Height          =   255
               Left            =   2955
               TabIndex        =   43
               Top             =   -30
               Width           =   285
            End
         End
         Begin VB.TextBox TxtSloc 
            Height          =   240
            Left            =   2385
            TabIndex        =   39
            Top             =   1530
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.TextBox TxtCia 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1335
            TabIndex        =   37
            Top             =   2340
            Width           =   690
         End
         Begin VB.TextBox TxtGrupo 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   6480
            TabIndex        =   36
            Top             =   1920
            Width           =   300
         End
         Begin VB.TextBox TxtDescCajero 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   2115
            TabIndex        =   34
            Top             =   1890
            Width           =   3435
         End
         Begin VB.TextBox TxtCodCajero 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1335
            TabIndex        =   33
            Top             =   1905
            Width           =   690
         End
         Begin VB.TextBox TxtVuelo 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   6525
            TabIndex        =   31
            Top             =   2325
            Width           =   1425
         End
         Begin VB.TextBox TxtDescCia 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   2130
            TabIndex        =   29
            Top             =   2340
            Width           =   3420
         End
         Begin VB.ComboBox CboCaja 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6495
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1470
            Width           =   1485
         End
         Begin VB.ComboBox cboLocal 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4305
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1470
            Width           =   1230
         End
         Begin VB.ListBox LstLocal 
            Height          =   255
            Left            =   7680
            TabIndex        =   23
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.ListBox LstEspigon 
            Height          =   255
            Left            =   5100
            TabIndex        =   21
            Top             =   1515
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.ComboBox CboEspigon 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1335
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1485
            Width           =   1500
         End
         Begin VB.TextBox TxtPasaporte 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1320
            TabIndex        =   8
            Top             =   330
            Width           =   2625
         End
         Begin MSMask.MaskEdBox MskFch 
            Height          =   315
            Left            =   1950
            TabIndex        =   16
            Top             =   720
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd-mm-yyyy"
            Mask            =   "##-##-####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MskTicket 
            Height          =   330
            Left            =   1950
            TabIndex        =   18
            Top             =   1065
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   582
            _Version        =   393216
            ClipMode        =   1
            BackColor       =   16777215
            ForeColor       =   16711680
            PromptInclude   =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label LblTurno 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Turno :"
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
            Left            =   6930
            TabIndex        =   53
            Top             =   1920
            Width           =   705
         End
         Begin VB.Label LblGrupo 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Grupo :"
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
            Left            =   5715
            TabIndex        =   35
            Top             =   1920
            Width           =   705
         End
         Begin VB.Label LblCajero 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cajero :"
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
            TabIndex        =   32
            Top             =   1905
            Width           =   1110
         End
         Begin VB.Label LblVuelo 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vuelo :"
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
            Left            =   5700
            TabIndex        =   30
            Top             =   2325
            Width           =   735
         End
         Begin VB.Label LblCia 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cía Aerea :"
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
            Left            =   120
            TabIndex        =   28
            Top             =   2340
            Width           =   1125
         End
         Begin VB.Label LblCaja 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Caja :"
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
            Index           =   0
            Left            =   5730
            TabIndex        =   27
            Top             =   1485
            Width           =   675
         End
         Begin VB.Label LblLocal 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Local :"
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
            Left            =   2955
            TabIndex        =   25
            Top             =   1485
            Width           =   1275
         End
         Begin VB.Label LblEspigon 
            BackColor       =   &H00C0FFC0&
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
            Left            =   105
            TabIndex        =   22
            Top             =   1500
            Width           =   1155
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nro. de Ticket :"
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
            Left            =   105
            TabIndex        =   19
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label LblFch 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha de Venta :"
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
            TabIndex        =   17
            Top             =   705
            Width           =   1680
         End
         Begin VB.Label LblPasaporte 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pasaporte :"
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
            Left            =   150
            TabIndex        =   7
            Top             =   345
            Width           =   1095
         End
      End
      Begin VB.TextBox txtCantReg 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7665
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   0
         Width           =   480
      End
      Begin VB.TextBox txtReg 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   255
         TabIndex        =   51
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Unid :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2475
         TabIndex        =   50
         Top             =   4470
         Width           =   1005
      End
      Begin VB.Label LblImp 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   630
         TabIndex        =   49
         Top             =   4425
         Width           =   1485
      End
      Begin VB.Label LblUni 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   3645
         TabIndex        =   48
         Top             =   4455
         Width           =   1230
      End
      Begin VB.Label de 
         Caption         =   "de"
         Height          =   255
         Left            =   7380
         TabIndex        =   5
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "chk"
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   615
      Width           =   990
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
         NumListImages   =   15
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":227C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":2596
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":28B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":2BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":2EE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":31FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":3518
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":3832
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":3B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":3E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":3F78
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":408A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":462C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":4E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmConsultaTicket.frx":561C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmConsultaTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As Recordset
Dim Rs_T As Recordset
Dim codigo As String
Dim Reclamo As Boolean
Dim Salida As Boolean
Dim Online As Boolean

Dim cl_Ticket As CLTicket

Dim CondConsulta As String
Dim CondTarj As String



Dim Modo As String

Dim sqlGral$
Dim Sql_T As String
Private Sub MeImpDatos()

Dim nom As String, NombreArchivo As String

Dim i%

On Error GoTo ErrPrint:

FrmConsultaTicket.caption = Aplicacion.SeteoProceso(FrmConsultaTicket.caption)

RptTicket.Connect = "ODBC;UID=" & Aplicacion.username & ";PWD=" & Aplicacion.password & ";DATABASE=INTER;"

RptTicket.ReportFileName = RutaReportes

If OptOnline(0) Then
   RptTicket.ReportFileName = RutaReportes & "Ticket.rpt"
 Else
   RptTicket.ReportFileName = RutaReportes & "Ticket_O.rpt"
End If

RptTicket.ParameterFields(1) = "PrFch;DATE(" & Year(MskFch.FormattedText) & "," & Month(MskFch.FormattedText) & "," & Day(MskFch.FormattedText) & ");true"
RptTicket.ParameterFields(2) = "PrCaja;" & CboCaja.Text & ";true"
RptTicket.ParameterFields(3) = "PrNro;" & MskTicket.Text & ";true"

RptTicket.Destination = 1    ' 0 pantalla
                                            '1 impresora
RptTicket.Action = True

DoEvents

FrmConsultaTicket.caption = Aplicacion.SeteoFin

Exit Sub
ErrPrint:
     FrmConsultaTicket.caption = Aplicacion.SeteoFin
     Exit Sub
        
End Sub
Public Sub MeSetearBotonesTarjetas()

If TxtCantTarj.Text < 2 Then
    CmdNext.Enabled = False
    CmdPrev.Enabled = False
  Else
    Select Case Val(TxtCantTarj.Text)
        Case Is = TxtTarj.Text
              CmdNext.Enabled = False
              CmdPrev.Enabled = True
        Case Is > TxtTarj.Text
            Select Case Val(TxtTarj.Text)
               Case Is > 1
                   CmdNext.Enabled = True
                   CmdPrev.Enabled = True
               Case Is = 1
                    CmdNext.Enabled = True
                    CmdPrev.Enabled = False
            End Select
     End Select
     
End If

End Sub
Private Sub MeTraerTarjeta()

Sql_T = ""
Sql_T = "             SELECT  P.nro_cuenta, T.descrip "
If OptOnline(0).Value Then
  Sql_T = Sql_T & " FROM    ventas.pago_tarjeta P, ventas.tarjetas T "
 Else
  Sql_T = Sql_T & " FROM   tickdia.pago_tarjeta_dia@link_nt_vta P, ventas.tarjetas T  "
End If
Sql_T = Sql_T & " WHERE cod_local = '" & rs!Cod_Local & "'"
Sql_T = Sql_T & "       AND  cod_sloc =  " & rs!Cod_Sloc
Sql_T = Sql_T & "       AND  cod_caja = " & rs!Cod_Caja
Sql_T = Sql_T & "       AND   fch_ticket = " & func_ToDate(Format(rs!Fch_Ticket, FTOFECHA))
Sql_T = Sql_T & "       AND   nro_ticket = " & rs!Nro_Ticket
Sql_T = Sql_T & "   AND   P.Cod_Tarjeta = T.cod_tarjeta "

If Aplicacion.ObtenerRsDAO(Sql_T, Rs_T) Then
    TxtCantTarj.Text = Aplicacion.CantReg(Rs_T)
    If TxtCantTarj.Text > 0 Then
        TxtTarj.Text = 1
        MellenarTarjeta
        MeSetearBotonesTarjetas
      Else
        TxtTarj.Text = 0
        CmdNext.Enabled = False
        CmdPrev.Enabled = False
        LblTarjeta.Visible = False
        txtCredit.Text = ""
    End If

End If

End Sub
Private Sub NuevaSeleccion()
Dim i%
Dim codigo
Dim Sql As String

If Modo = "MOD" Then
    SetBotonesGeneral False
Else
End If

'Limpiar campos de pantallas

MskFch.Text = ""
TxtGrupo.Text = ""
TxtPasaporte.Text = ""
TxtCodCajero.Text = ""
TxtDescCajero.Text = ""
TxtCia.Text = ""
TxtDescCia.Text = ""
TxtVuelo.Text = ""
MskTicket.Text = ""
TxtTurno.Text = ""
LblImp.caption = ""
LblUni.caption = ""
CboEspigon.Clear
cboLocal.Clear
TxtSloc.Text = ""
CboCaja.Clear
txtCredit.Text = ""
codigo = ""
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
TxtCantTarj.Text = ""
TxtTarj.Text = ""
'CmdPasarTicket.Visible = False
LblTarjeta.caption = ""

FramCab.Enabled = True
FraFecha.Enabled = True
MskFechaDesde.SetFocus

'arma el SQL para llenar el combo de Aeropuerto
Sql = " SELECT cod_sdep,descrip FROM baires.subdependencia "
Sql = Sql & " WHERE cod_sdep <> 'BODE'  AND cod_sdep <> 'BODA' "
Sql = Sql & " ORDER BY cod_sdep"
 
Aplicacion.ObtenerRsDAO Sql, rs

Func_LlenarComboLst CboEspigon, LstEspigon, rs

txtCantReg.Text = ""
txtReg.Text = ""

CmdDetalle.Enabled = False

SetBotonesGeneral False

chk.Value = 0

End Sub
Private Sub MeCargarDatos()
Dim Sql$

FrmConsultaTicket.caption = Aplicacion.SeteoProceso(FrmConsultaTicket.caption)
        
    CondConsulta = ArmarCondicion

If CondTarj = "" Then
  If codigo = "" Then
    sqlGral$ = ""
    sqlGral$ = " SELECT  "
    sqlGral$ = sqlGral$ & " H.fch_ticket,H.cod_depn,H.cod_sdep,H.cod_local,H.cod_sloc,H.nro_ticket, "
    sqlGral$ = sqlGral$ & "                 H.cod_caja,H.grupo_venta,H.cod_cia_aerea,"
    sqlGral$ = sqlGral$ & "                 H.cod_vuelo,H.nro_pax,H.cod_cajero,H.sub_grupo,"
    sqlGral$ = sqlGral$ & "                 H.cant_items, H.importe,B.descrip, "
    sqlGral$ = sqlGral$ & "                 H.cod_cia_aerea, H.cod_vuelo "
    If OptOnline(0).Value Then
      sqlGral$ = sqlGral$ & " FROM ventas.ticket_h H"
      sqlGral$ = sqlGral$ & ", ventas.companias B "
     Else
      sqlGral$ = sqlGral$ & " FROM     tickdia.ticket_h_dia@link_nt_vta H, ventas.companias B "
    End If
    sqlGral$ = sqlGral$ & " WHERE  h.fch_ticket  BETWEEN " & func_ToDate(MskFechaDesde.FormattedText) & " And " & func_ToDate(MskFchHasta.FormattedText)
    sqlGral$ = sqlGral$ & CondConsulta
    sqlGral$ = sqlGral$ & "                   and H.cod_cia_aerea = B.cod_compania "
    sqlGral$ = sqlGral$ & " ORDER BY FCH_TICKET "
  Else
    sqlGral$ = ""
    sqlGral$ = " SELECT DISTINCT H.fch_ticket,H.cod_depn,H.cod_sdep,H.cod_local, "
    sqlGral$ = sqlGral$ & "                 H.cod_sloc,H.nro_ticket, "
    sqlGral$ = sqlGral$ & "                 H.cod_caja,H.grupo_venta,H.cod_cia_aerea,"
    sqlGral$ = sqlGral$ & "                 H.cod_vuelo,H.nro_pax,H.cod_cajero,H.sub_grupo,"
    sqlGral$ = sqlGral$ & "                 H.cant_items, H.importe,B.descrip "
    If OptOnline(0).Value Then
      sqlGral$ = sqlGral$ & " FROM     ventas.ticket_h H, ventas.companias B,ventas.ticket_d D "
     Else
      sqlGral$ = sqlGral$ & " FROM     tickdia.ticket_h_dia@link_nt_vta H, ventas.companias B,ventas.ticket_d D "
    End If
    sqlGral$ = sqlGral$ & " WHERE  h.fch_ticket  BETWEEN " & func_ToDate(MskFechaDesde.FormattedText) & " And " & func_ToDate(MskFchHasta.FormattedText)
    sqlGral$ = sqlGral$ & " And H.fch_ticket = D.fch_ticket "
    sqlGral$ = sqlGral$ & " And H.cod_local = D.cod_local "
    sqlGral$ = sqlGral$ & " And H.cod_sloc = D.cod_sloc "
    sqlGral$ = sqlGral$ & " And H.nro_ticket = D.nro_ticket "
    sqlGral$ = sqlGral$ & CondConsulta
    sqlGral$ = sqlGral$ & "                   and H.cod_cia_aerea = B.cod_compania "
    sqlGral$ = sqlGral$ & " ORDER BY FCH_TICKET "
  End If
Else
  If codigo = "" Then
   sqlGral$ = ""
   sqlGral$ = " SELECT                   H.fch_ticket,H.cod_depn,H.cod_sdep,H.cod_local,H.cod_sloc,H.nro_ticket, "
   sqlGral$ = sqlGral$ & "                 H.cod_caja,H.grupo_venta,H.cod_cia_aerea, "
   sqlGral$ = sqlGral$ & "                 H.cod_vuelo,H.nro_pax,H.cod_cajero,H.sub_grupo,  "
   sqlGral$ = sqlGral$ & "                 H.cant_items, H.importe,B.descrip "
   If OptOnline(0).Value Then
     sqlGral$ = sqlGral$ & " FROM      "
     sqlGral$ = sqlGral$ & "                 ventas.ticket_h H, ventas.companias B, "
     sqlGral$ = sqlGral$ & "                 ventas.pago_tarjeta P, ventas.tarjetas T "
    Else
     sqlGral$ = sqlGral$ & " FROM      "
     sqlGral$ = sqlGral$ & "                 tickdia.ticket_h_dia@link_nt_vta H, ventas.companias B, "
     sqlGral$ = sqlGral$ & "                 tickdia.pago_tarjeta_dia@link_nt_vta P, ventas.tarjetas T "
   End If
   sqlGral$ = sqlGral$ & " WHERE   "
   sqlGral$ = sqlGral$ & "                 H.COD_LOCAL = P.COD_LOCAL "
   sqlGral$ = sqlGral$ & "                 AND H.COD_SLOC = P.COD_SLOC "
   sqlGral$ = sqlGral$ & "                 AND H.COD_CAJA = P.COD_CAJA "
   sqlGral$ = sqlGral$ & "                 AND H.FCH_TICKET = P.FCH_TICKET "
   sqlGral$ = sqlGral$ & "                 AND H.NRO_TICKET = P.NRO_TICKET "
   sqlGral$ = sqlGral$ & "                 AND P.COD_TARJETA = T.COD_TARJETA "
   sqlGral$ = sqlGral$ & "                 and H.cod_cia_aerea = B.cod_compania "
   sqlGral$ = sqlGral$ & CondConsulta
   sqlGral$ = sqlGral$ & " AND h.fch_ticket BETWEEN " & func_ToDate(MskFechaDesde.FormattedText) & " And " & func_ToDate(MskFchHasta.FormattedText)
  Else
  sqlGral$ = ""
   sqlGral$ = " SELECT    DISTINCT               H.fch_ticket,H.cod_depn,H.cod_sdep,H.cod_local,H.cod_sloc,H.nro_ticket, "
   sqlGral$ = sqlGral$ & "                 H.cod_caja,H.grupo_venta,H.cod_cia_aerea, "
   sqlGral$ = sqlGral$ & "                 H.cod_vuelo,H.nro_pax,H.cod_cajero,H.sub_grupo,  "
   sqlGral$ = sqlGral$ & "                 H.cant_items, H.importe,B.descrip "
   If OptOnline(0).Value Then
     sqlGral$ = sqlGral$ & " FROM      "
     sqlGral$ = sqlGral$ & "                 ventas.ticket_h H, ventas.companias B, "
     sqlGral$ = sqlGral$ & "                 ventas.pago_tarjeta P, ventas.tarjetas T "
    Else
     sqlGral$ = sqlGral$ & " FROM      "
     sqlGral$ = sqlGral$ & "                 tickdia.ticket_h_dia@link_nt_vta H, ventas.companias B, "
     sqlGral$ = sqlGral$ & "                 tickdia.pago_tarjeta_dia@link_nt_vta P, ventas.tarjetas T "
   End If
   sqlGral$ = sqlGral$ & " WHERE   "
   sqlGral$ = sqlGral$ & "                 H.COD_LOCAL = P.COD_LOCAL "
   sqlGral$ = sqlGral$ & "                 AND H.COD_SLOC = P.COD_SLOC "
   sqlGral$ = sqlGral$ & "                 AND H.COD_CAJA = P.COD_CAJA "
   sqlGral$ = sqlGral$ & "                 AND H.FCH_TICKET = P.FCH_TICKET "
   sqlGral$ = sqlGral$ & "                 AND H.NRO_TICKET = P.NRO_TICKET "
   sqlGral$ = sqlGral$ & "                 AND P.COD_TARJETA = T.COD_TARJETA "
   sqlGral$ = sqlGral$ & "                 AND H.COD_CIA_AEREA = B.COD_COMPANIA "
   sqlGral$ = sqlGral$ & "                 AND H.COD_LOCAL = D.COD_LOCAL "
   sqlGral$ = sqlGral$ & "                 AND H.COD_SLOC = D.COD_SLOC "
   sqlGral$ = sqlGral$ & "                 AND H.NRO_TICKET = D.NRO_TICKET "
   sqlGral$ = sqlGral$ & "                 AND H.FCH_TICKET = D.FCH_TICKET "
   sqlGral$ = sqlGral$ & CondConsulta
   sqlGral$ = sqlGral$ & " AND h.fch_ticket BETWEEN " & func_ToDate(MskFechaDesde.FormattedText) & " And " & func_ToDate(MskFchHasta.FormattedText)

  End If
  
End If

If Aplicacion.ObtenerRsDAO(sqlGral$, rs) Then
    txtCantReg.Text = Aplicacion.CantReg(rs)
    'If Reclamo = True Then
    '   CmdPasarTicket.Visible = True
    '   CmdPasarTicket.Enabled = True
    'End If
    If txtCantReg.Text > 0 Then
       txtReg.Text = 1
       SetBotonesGeneral True
       MellenarPantalla
       MeSetearBotonesToolBar
       CmdDetalle.Enabled = True
       'CmdPasarTicket.Visible = True
       'CmdPasarTicket.Enabled = True
    Else
       txtReg.Text = 0
       CmdDetalle.Enabled = False
       'CmdPasarTicket.Enabled = False
    End If
End If

FrmConsultaTicket.caption = Aplicacion.SeteoFin

End Sub
Private Sub MeLlenarObjeto()

cl_Ticket.Apellido = TxtDescCajero.Text
cl_Ticket.Cod_Caja = CboCaja.Text
cl_Ticket.Cod_Cajero = TxtCodCajero.Text
cl_Ticket.Cod_Cia_Aerea = TxtCia.Text
cl_Ticket.Cod_Sloc = TxtSloc.Text
cl_Ticket.Cod_Local = LstLocal.List(cboLocal.ListIndex)
cl_Ticket.Cod_Sdep = LstEspigon.List(CboEspigon.ListIndex)
cl_Ticket.Cod_Vuelo = TxtVuelo.Text
cl_Ticket.Descrip = TxtDescCia.Text
cl_Ticket.Grupo_Venta = TxtGrupo.Text
cl_Ticket.Nro_Pax = TxtPasaporte.Text
cl_Ticket.Nro_Ticket = MskTicket.Text

End Sub
Private Function L_TodoCargado() As Boolean
    L_TodoCargado = False
End Function
Private Sub MellenarPantalla()

MskFch.Text = IIf(IsNull(rs!Fch_Ticket), "", Format(rs!Fch_Ticket, FTOFECHA))
TxtGrupo.Text = IIf(IsNull(rs!Grupo_Venta), "", rs!Grupo_Venta)
TxtPasaporte.Text = IIf(IsNull(rs!Nro_Pax), "", Trim(rs!Nro_Pax))
TxtCodCajero.Text = IIf(IsNull(rs!Cod_Cajero), "", rs!Cod_Cajero)
TxtCia.Text = IIf(IsNull(rs!Cod_Cia_Aerea), "", rs!Cod_Cia_Aerea)
TxtDescCia.Text = IIf(IsNull(rs!Descrip), "", rs!Descrip)
TxtVuelo.Text = IIf(IsNull(rs!Cod_Vuelo), "", rs!Cod_Vuelo)
MskTicket.Text = IIf(IsNull(rs!Nro_Ticket), "", rs!Nro_Ticket)
TxtTurno.Text = IIf(IsNull(rs!sub_grupo), "", rs!sub_grupo)
LblImp.caption = IIf(IsNull(rs!Importe), "", rs!Importe)
LblUni.caption = IIf(IsNull(rs!cant_items), "", rs!cant_items)
Func_SetearCboConLst CboEspigon, LstEspigon, rs!Cod_Sdep
Func_SetearCboSTR cboLocal, rs!Cod_Local
TxtSloc.Text = IIf(IsNull(rs!Cod_Sloc), "", rs!Cod_Sloc)
Func_SetearCboINT CboCaja, rs!Cod_Caja

MeTraerTarjeta

Frame1.Enabled = True
FraFecha.Enabled = False
FramCab.Enabled = False
Frame2.Enabled = False
Frame3.Enabled = False

CmdDetalle.Enabled = True

End Sub

Private Sub MellenarTarjeta()
txtCredit.Text = Rs_T!nro_cuenta
LblTarjeta.Visible = True
LblTarjeta.caption = Rs_T!Descrip
End Sub
Private Sub SetBotonesGeneral(valor As Boolean)
    
    Tollbar.Buttons(1).Enabled = valor
    Tollbar.Buttons(2).Enabled = Not valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = valor
    Tollbar.Buttons(6).Enabled = valor
    Tollbar.Buttons(7).Enabled = valor

    Tollbar.Buttons(9).Enabled = Not valor
    Tollbar.Buttons(10).Enabled = valor
    Tollbar.Buttons(11).Enabled = valor
    
'habilitar frames

    If Not valor Then
        txtReg.Text = 0
        txtCantReg.Text = 0
    End If

End Sub

Private Function ArmarCondicion()
Dim Con$
Dim ConTarj$
Dim c As Integer
Dim Espacios As Integer

If MskTicket.Text <> "" Then
  Con$ = Con$ & " And H.nro_ticket = " & MskTicket.Text
End If

If LstEspigon.List(CboEspigon.ListIndex) <> "" Then
  Con$ = Con$ & " And H.cod_sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
End If

If LstLocal.List(cboLocal.ListIndex) <> "" Then
  Con$ = Con$ & " And H.cod_local = '" & LstLocal.List(cboLocal.ListIndex) & "'"
End If

If CboCaja.Text <> "" Then
  Con$ = Con$ & " And H.cod_caja = " & CboCaja.Text
End If

If TxtGrupo.Text <> "" Then
  Con$ = Con$ & " And H.grupo_venta = '" & TxtGrupo.Text & "'"
End If

If TxtVuelo.Text <> "" Then
  Con$ = Con$ & " And H.cod_vuelo = " & TxtVuelo.Text
End If

If TxtCia.Text <> "" Then
  Con$ = Con$ & " And H.cod_cia_aerea = " & TxtCia.Text
End If

If TxtCodCajero.Text <> "" Then
  Con$ = Con$ & " And H.cod_cajero = " & TxtCodCajero.Text
End If

If TxtPasaporte.Text <> "" Then
Con$ = Con$ & " And LTRIM(RTRIM(H.nro_pax)) = '" & TxtPasaporte.Text & "'"
End If

If codigo <> "" Then
   Con$ = Con$ & " And D.Cod_Prod = " & codigo
End If

If txtCredit.Text <> "" Then
  ConTarj$ = Con$ & " And LTRIM(RTRIM(P.nro_cuenta)) = '" & txtCredit.Text & "' "
End If

If ConTarj$ = "" Then
      ArmarCondicion = Con$
      CondTarj = ""
   Else
     ArmarCondicion = ConTarj$
     CondTarj = ConTarj$
End If



End Function

Private Function MeReconsultar() As Integer
Dim Sql$
Dim i%
    
FrmConsultaTicket.caption = Aplicacion.SeteoProceso(FrmConsultaTicket.caption)
    

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

FrmConsultaTicket.caption = Aplicacion.SeteoProceso(FrmConsultaTicket.caption)

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
'    Tollbar.Buttons(12).Enabled = True
'    Tollbar.Buttons(13).Enabled = True
ElseIf txtReg.Text = txtCantReg.Text Then
    Tollbar.Buttons(4).Enabled = True
    Tollbar.Buttons(5).Enabled = True
    Tollbar.Buttons(6).Enabled = False
    Tollbar.Buttons(7).Enabled = False
'    Tollbar.Buttons(12).Enabled = True
'    Tollbar.Buttons(13).Enabled = True
ElseIf txtReg.Text = 1 Then
    Tollbar.Buttons(4).Enabled = False
    Tollbar.Buttons(5).Enabled = False
    Tollbar.Buttons(6).Enabled = True
    Tollbar.Buttons(7).Enabled = True
'    Tollbar.Buttons(12).Enabled = True
'    Tollbar.Buttons(13).Enabled = True
Else
    Tollbar.Buttons(4).Enabled = True
    Tollbar.Buttons(5).Enabled = True
    Tollbar.Buttons(6).Enabled = True
    Tollbar.Buttons(7).Enabled = True
   ' Tollbar.Buttons(12).Enabled = True
'    Tollbar.Buttons(13).Enabled = True
    
End If
    
End Sub
Private Sub SeteoBotonesMod(valor As Boolean)
    
    
    Tollbar.Buttons(1).Enabled = Not valor
    Tollbar.Buttons(2).Enabled = valor
    
    Tollbar.Buttons(4).Enabled = Not valor
    Tollbar.Buttons(5).Enabled = Not valor
    Tollbar.Buttons(6).Enabled = Not valor
    Tollbar.Buttons(7).Enabled = Not valor

    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor
    Tollbar.Buttons(11).Enabled = valor
'habilitar o des frames y/o campos

End Sub
Private Sub CboEspigon_Change()
cboLocal.Clear
CboCaja.Clear
End Sub
Private Sub CboEspigon_Click()
Dim Sql As String
Dim rs As Recordset

cboLocal.Clear
CboCaja.Clear

Sql = " SELECT cod_loc FROM baires.local "
Sql = Sql & " WHERE cod_sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
Sql = Sql & " ORDER BY cod_loc"
 
Aplicacion.ObtenerRsDAO Sql, rs

Func_LlenarComboNoItem cboLocal, rs

End Sub
Private Sub CboLocal_Click()
Dim Sql As String
Dim rs As Recordset

Sql = " SELECT cod_caja FROM ventas.local_caja "
'Sql = Sql & " WHERE cod_local = "
'Sql = Sql & " DECODE('" & CboLocal.Text & "','L05','L21','L06','L21','L07','L22','L08','L22','" & CboLocal.Text & "') "
Sql = Sql & " ORDER BY cod_caja"
 
Aplicacion.ObtenerRsDAO Sql, rs

Func_LlenarComboNoItem CboCaja, rs
  
End Sub
Private Sub CmdDetalle_Click()
L_DatosGrilla
End Sub
Private Sub CmdFechaDesde_Click()
Dim fch As Date

If MskFechaDesde.Text <> "" Then
    fch = MskFechaDesde.FormattedText
Else
    fch = Date
End If

frmFecha.MuestroFormFecha fch

MskFechaDesde.Text = Format$(fch, FTOFECHA)

MskFechaDesde.SetFocus

End Sub




Private Sub CmdFechaHasta_Click()
Dim fch As Date

If MskFchHasta.Text <> "" Then
    fch = MskFchHasta.FormattedText
Else
    fch = Date
End If

frmFecha.MuestroFormFecha fch

MskFchHasta.Text = Format$(fch, FTOFECHA)

MskFchHasta.SetFocus

End Sub

Private Sub CmdNext_Click()
Dim pos As String
Dim saltear As Boolean

saltear = True

pos = TxtTarj.Text

saltear = False
Func_MoverSiguiente Rs_T, pos

If Not saltear Then
    TxtTarj.Text = pos
    MellenarTarjeta
    MeSetearBotonesTarjetas
End If

End Sub

Private Sub CmdPrev_Click()
Dim pos As String
Dim saltear As Boolean

saltear = True

pos = TxtTarj.Text

saltear = False
Func_MoverAnterior Rs_T, pos

If Not saltear Then
    TxtTarj.Text = pos
    MellenarTarjeta
    MeSetearBotonesTarjetas
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
            Call TollBar_ButtonClick(Tollbar.Buttons(19))
    End Select
    If Modo = "MOD" And Val(txtCantReg.Text) > 0 Then
    Select Case KeyCode
        Case 37 'Izq
            If Tollbar.Buttons(5).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(8))
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

End If
'Debug.Print KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub Form_Load()

Dim Sql As String

Width = Screen.Width * 0.75
Height = Screen.Height * 0.75
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

Set cl_Ticket = New CLTicket

'CmdPasarTicket.Visible = BotonVisible
'CmdPasarTicket.Visible = True

MskFechaDesde.Text = "01-" & Format$(Month(Date), "0#") & "-" & Format$(Year(Date), "####")
MskFchHasta.Text = Format$(Date - 1, FTOFECHA)

'arma el SQL para llenar el combo de Aeropuerto
Sql = " SELECT cod_sdep,descrip FROM baires.subdependencia "
Sql = Sql & " WHERE cod_sdep <> 'BODE'  AND cod_sdep <> 'BODA' "
Sql = Sql & " ORDER BY cod_sdep"
 
Aplicacion.ObtenerRsDAO Sql, rs

Func_LlenarComboLst CboEspigon, LstEspigon, rs

CmdDetalle.Enabled = False
'CmdPasarTicket.Enabled = False

OptOnline(0).Value = 1

End Sub
Private Sub Form_Unload(Cancel As Integer)
'BotonVisible = False
'Reclamo = False
'NoReclamo = False
End Sub

Private Sub OptOnline_Click(Index As Integer)

Select Case Index
   Case 0
       FraFecha.Enabled = True
       MskFechaDesde.Text = "01-" & Format$(Month(Date), "0#") & "-" & Format$(Year(Date), "####")
       MskFchHasta.Text = Format$(Date - 1, FTOFECHA)
   Case 1
       MskFechaDesde.Text = Format$(Date, FTOFECHA)
       MskFchHasta.Text = Format$(Date, FTOFECHA)
       FraFecha.Enabled = False
End Select

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
'         MePrepararMod
    Case "g"
     '    MeEliminar
    Case "h"
        If Modo = "MOD" Then
         '   MeActualizar
        Else
        '    L_AltasDatos
        End If
    Case "i"
     '   MeAbortarMod
    Case "j"
        If chk.Value = 1 Then
            If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
                If Modo = "MOD" Then
                 '   MeActualizar
                Else
               '     L_AltasDatos
                End If
            End If
        End If
        Salida = False
        Unload Me
    Case "k"
        NuevaSeleccion
    Case "l"
        MeCargarDatos
    Case "t"
        'FrmBuscarCodigo.Show 1
    Case "n"
        L_DatosGrilla
    Case "m"
        MeImpDatos
    Case "o"
    '    MePrepararAgregar
    Case "p"
    '    MePrepararAlterar
End Select

If Not saltear Then
    txtReg.Text = pos
    MellenarPantalla
    MeSetearBotonesToolBar
End If

End Sub
Private Sub L_DatosGrilla()
Dim Sql As String
Dim i
Dim nro As Integer

On Error GoTo DG:

Sql = "          SELECT D.cod_prod, P.descrip,P.cod_prov, P.cod_rubr, D.cantidad, D.importe, E.apellido || ', ' || E.nombre promotor "
If OptOnline(0).Value Then
  Sql = Sql & " FROM     ventas.ticket_d D, baires.producto P, personal.empleado E "
 Else
  Sql = Sql & " FROM     tickdia.ticket_d_dia@link_nt_vta D, baires.producto P, personal.empleado E "
End If
Sql = Sql & " WHERE D.fch_ticket = " & func_ToDate(MskFch.FormattedText)
Sql = Sql & "        AND D.nro_ticket = " & MskTicket.Text
Sql = Sql & "        AND D.cod_caja = " & CboCaja.Text
Sql = Sql & "        AND D.cod_local = '" & Trim(cboLocal.Text) & "'"
Sql = Sql & "        AND P.cod_prod = D.cod_prod "
Sql = Sql & "        AND D.vendedor = E.legajo(+) "

nro = frmGridTicket.DatosGrilla(Sql)
 
If nro > 0 Then
    rs.MoveFirst
    For i = 1 To nro - 1
        rs.MoveNext
    Next
    MellenarPantalla
    txtReg.Text = nro
    MeSetearBotonesToolBar

DG:
    Exit Sub
End If

End Sub
