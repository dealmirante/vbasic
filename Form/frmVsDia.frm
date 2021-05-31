VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVsDia 
   Caption         =   "Estimados vs Ventas-Tickets-Pasajeros"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   8670
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
      Height          =   510
      Index           =   2
      Left            =   7485
      Picture         =   "frmVsDia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1230
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
      Height          =   480
      Index           =   1
      Left            =   7470
      Picture         =   "frmVsDia.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   105
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
      Height          =   480
      Index           =   0
      Left            =   7470
      Picture         =   "frmVsDia.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   675
      Width           =   615
   End
   Begin VB.CommandButton botExcel 
      Caption         =   "Excel"
      Height          =   540
      Left            =   7485
      Picture         =   "frmVsDia.frx":0A26
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1815
      Width           =   615
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
      Height          =   2235
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   7275
      Begin VB.Frame frMod 
         Caption         =   "Modelos"
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
         Height          =   2025
         Left            =   3495
         TabIndex        =   10
         Top             =   120
         Width           =   3600
         Begin VB.ComboBox cboModPax 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1605
            Width           =   2310
         End
         Begin VB.ComboBox cboModTicket 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1260
            Width           =   2310
         End
         Begin VB.TextBox txtMesMod 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   255
            Width           =   735
         End
         Begin VB.ComboBox cboModImp 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   915
            Width           =   2325
         End
         Begin VB.OptionButton optMod 
            Caption         =   "Real"
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
            Left            =   2430
            TabIndex        =   13
            Top             =   210
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.OptionButton optMod 
            Caption         =   "Modelo"
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
            Left            =   150
            TabIndex        =   12
            Top             =   630
            Value           =   -1  'True
            Width           =   930
         End
         Begin VB.ComboBox cboAnio 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2415
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   465
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pasajeros"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   135
            TabIndex        =   19
            Top             =   1605
            Width           =   915
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ticket"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   135
            TabIndex        =   18
            Top             =   1260
            Width           =   915
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Importe"
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
            Left            =   150
            TabIndex        =   17
            Top             =   915
            Width           =   915
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
            Index           =   1
            Left            =   150
            TabIndex        =   16
            Top             =   255
            Width           =   780
         End
      End
      Begin VB.ListBox LstEspigon 
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   915
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox CboCodAeropuerto 
         Height          =   315
         ItemData        =   "frmVsDia.frx":0FB8
         Left            =   1260
         List            =   "frmVsDia.frx":0FBA
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1170
         Width           =   2070
      End
      Begin VB.ComboBox CboEspigon 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1710
         Width           =   2070
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   300
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
         Left            =   960
         TabIndex        =   5
         Top             =   750
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
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
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   315
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
         TabIndex        =   8
         Top             =   750
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
         Left            =   150
         TabIndex        =   7
         Top             =   1170
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
         Left            =   135
         TabIndex        =   6
         Top             =   1710
         Width           =   1065
      End
   End
   Begin TabDlg.SSTab tabEst 
      Height          =   3660
      Left            =   15
      TabIndex        =   26
      Top             =   2250
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   6456
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   7
      TabHeight       =   441
      ForeColor       =   255
      TabCaption(0)   =   "TOTAL Mes"
      TabPicture(0)   =   "frmVsDia.frx":0FBC
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tabTotal"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "T. Mañana"
      TabPicture(1)   =   "frmVsDia.frx":0FD8
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabGA"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "T. Tarde"
      TabPicture(2)   =   "frmVsDia.frx":0FF4
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabGB"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "T. Noche"
      TabPicture(3)   =   "frmVsDia.frx":1010
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tabGC"
      Tab(3).Control(0).Enabled=   0   'False
      Begin TabDlg.SSTab tabTotal 
         Height          =   3150
         Left            =   135
         TabIndex        =   27
         Top             =   315
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5556
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   441
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmVsDia.frx":102C
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprTotal(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmVsDia.frx":1048
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprTotal(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pax Trans."
         TabPicture(2)   =   "frmVsDia.frx":1064
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprTotal(2)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick"
         TabPicture(3)   =   "frmVsDia.frx":1080
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprTotal(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom Imp x Pax"
         TabPicture(4)   =   "frmVsDia.frx":109C
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprTotal(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprTotal 
            Height          =   2670
            Index           =   2
            Left            =   -74775
            OleObjectBlob   =   "frmVsDia.frx":10B8
            TabIndex        =   36
            Top             =   120
            Width           =   7515
         End
         Begin FPSpread.vaSpread sprTotal 
            Height          =   2670
            Index           =   1
            Left            =   -74760
            OleObjectBlob   =   "frmVsDia.frx":18F1
            TabIndex        =   35
            Top             =   105
            Width           =   7515
         End
         Begin FPSpread.vaSpread sprTotal 
            Height          =   2655
            Index           =   0
            Left            =   270
            OleObjectBlob   =   "frmVsDia.frx":212A
            TabIndex        =   28
            Top             =   120
            Width           =   7515
         End
         Begin FPSpread.vaSpread sprTotal 
            Height          =   2670
            Index           =   3
            Left            =   -74775
            OleObjectBlob   =   "frmVsDia.frx":2969
            TabIndex        =   43
            Top             =   105
            Width           =   7515
         End
         Begin FPSpread.vaSpread sprTotal 
            Height          =   2670
            Index           =   4
            Left            =   -74760
            OleObjectBlob   =   "frmVsDia.frx":2F11
            TabIndex        =   44
            Top             =   105
            Width           =   7515
         End
      End
      Begin TabDlg.SSTab tabGA 
         Height          =   3210
         Left            =   -74820
         TabIndex        =   29
         Top             =   315
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5662
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   441
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmVsDia.frx":34B9
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprGA(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmVsDia.frx":34D5
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprGA(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pr. Imp x Ticket"
         TabPicture(2)   =   "frmVsDia.frx":34F1
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprGA(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprGA 
            Height          =   2715
            Index           =   2
            Left            =   -74760
            OleObjectBlob   =   "frmVsDia.frx":350D
            TabIndex        =   38
            Top             =   120
            Width           =   7515
         End
         Begin FPSpread.vaSpread sprGA 
            Height          =   2700
            Index           =   1
            Left            =   -74745
            OleObjectBlob   =   "frmVsDia.frx":3AB5
            TabIndex        =   37
            Top             =   135
            Width           =   7515
         End
         Begin FPSpread.vaSpread sprGA 
            Height          =   2700
            Index           =   0
            Left            =   255
            OleObjectBlob   =   "frmVsDia.frx":42EE
            TabIndex        =   30
            Top             =   150
            Width           =   7515
         End
      End
      Begin TabDlg.SSTab tabGB 
         Height          =   3210
         Left            =   -74790
         TabIndex        =   31
         Top             =   315
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5662
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   441
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmVsDia.frx":4B2D
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprGB(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmVsDia.frx":4B49
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprGB(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pr. Imp x Ticket"
         TabPicture(2)   =   "frmVsDia.frx":4B65
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprGB(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprGB 
            Height          =   2700
            Index           =   2
            Left            =   -74805
            OleObjectBlob   =   "frmVsDia.frx":4B81
            TabIndex        =   40
            Top             =   120
            Width           =   7515
         End
         Begin FPSpread.vaSpread sprGB 
            Height          =   2700
            Index           =   1
            Left            =   -74790
            OleObjectBlob   =   "frmVsDia.frx":5129
            TabIndex        =   39
            Top             =   135
            Width           =   7515
         End
         Begin FPSpread.vaSpread sprGB 
            Height          =   2700
            Index           =   0
            Left            =   195
            OleObjectBlob   =   "frmVsDia.frx":5962
            TabIndex        =   32
            Top             =   135
            Width           =   7515
         End
      End
      Begin TabDlg.SSTab tabGC 
         Height          =   3210
         Left            =   -74775
         TabIndex        =   33
         Top             =   315
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5662
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   441
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmVsDia.frx":61A1
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprGC(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmVsDia.frx":61BD
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprGC(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pr. Imp x Ticket"
         TabPicture(2)   =   "frmVsDia.frx":61D9
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprGC(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprGC 
            Height          =   2715
            Index           =   2
            Left            =   -74745
            OleObjectBlob   =   "frmVsDia.frx":61F5
            TabIndex        =   42
            Top             =   105
            Width           =   7515
         End
         Begin FPSpread.vaSpread sprGC 
            Height          =   2715
            Index           =   1
            Left            =   -74790
            OleObjectBlob   =   "frmVsDia.frx":679D
            TabIndex        =   41
            Top             =   105
            Width           =   7515
         End
         Begin FPSpread.vaSpread sprGC 
            Height          =   2715
            Index           =   0
            Left            =   180
            OleObjectBlob   =   "frmVsDia.frx":6FD6
            TabIndex        =   34
            Top             =   120
            Width           =   7515
         End
      End
   End
End
Attribute VB_Name = "frmVsDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsEstim As Recordset
Dim rsEstimT As Recordset
Dim rs As Recordset
Dim rsPax As Recordset

Dim siImp As Boolean, siTic As Boolean, siPax As Boolean
Private Sub L_LlenarGrillasImp(col As Integer)
Dim FD As Date
Dim SumImp As Double, fila As Integer
Dim fch As Variant

On Error GoTo ErrGLC:

frmVsDia.caption = Aplicacion.SeteoProceso(frmVsDia.caption)

fila = 1
If rs.RecordCount > 0 And sprTotal(0).MaxRows > 0 Then
rs.MoveFirst

'Existen valores de importes
siImp = True

Do While fila <= sprTotal(0).MaxRows
    
    sprTotal(0).GetText 1, fila, fch
    
    SumImp = 0
    If Not rs.EOF Then
    Do While Day(CDate(fch)) = Day(CDate(rs!fch_ticket))
        Select Case rs!Grupo
            Case "A"
                If sprGA(0).MaxRows > 0 Then
                    sprGA(0).SetText col, fila, str(rs!imp)
                    Spread_TotalesGrillaAcum sprGA(0), 3, 7, fila
                End If
            Case "B"
                If sprGB(0).MaxRows > 0 Then
                    sprGB(0).SetText col, fila, str(rs!imp)
                    Spread_TotalesGrillaAcum sprGB(0), 3, 7, fila
                End If
            Case "C"
                If sprGC(0).MaxRows > 0 Then
                    sprGC(0).SetText col, fila, str(rs!imp)
                    Spread_TotalesGrillaAcum sprGC(0), 3, 7, fila
                End If
        End Select
        SumImp = SumImp + rs!imp
        
        rs.MoveNext
        If rs.EOF Then
            Exit Do
        End If
    Loop
    sprTotal(0).SetText col, fila, str(SumImp)
    
    Spread_TotalesGrillaAcum sprTotal(0), 3, 7, fila
        
    spread_ResaltarCelda sprTotal(0), 5, fila
    spread_ResaltarCelda sprGA(0), 5, fila
    spread_ResaltarCelda sprGB(0), 5, fila
    spread_ResaltarCelda sprGC(0), 5, fila
    spread_ResaltarCelda sprTotal(0), 6, fila
    spread_ResaltarCelda sprGA(0), 6, fila
    spread_ResaltarCelda sprGB(0), 6, fila
    spread_ResaltarCelda sprGC(0), 6, fila
    spread_ResaltarCelda sprTotal(0), 9, fila
    spread_ResaltarCelda sprGA(0), 9, fila
    spread_ResaltarCelda sprGB(0), 9, fila
    spread_ResaltarCelda sprGC(0), 9, fila
    spread_ResaltarCelda sprTotal(0), 10, fila
    spread_ResaltarCelda sprGA(0), 10, fila
    spread_ResaltarCelda sprGB(0), 10, fila
    spread_ResaltarCelda sprGC(0), 10, fila
    
    fila = fila + 1
    Else
        Exit Do
    End If
Loop
    
If col = 3 Then
    Spread_TotalesGrillas sprTotal(0), sprTotal(0).MaxCols - 7, 2
    Spread_TotalesGrillas sprGA(0), sprGA(0).MaxCols - 7, 2
    Spread_TotalesGrillas sprGB(0), sprGB(0).MaxCols - 7, 2
    Spread_TotalesGrillas sprGC(0), sprGC(0).MaxCols - 7, 2
    
    spread_ResaltarCelda sprTotal(0), 5, sprTotal(0).MaxRows
    spread_ResaltarCelda sprTotal(0), 6, sprTotal(0).MaxRows
    spread_ResaltarCelda sprGA(0), 5, sprGA(0).MaxRows
    spread_ResaltarCelda sprGA(0), 6, sprGA(0).MaxRows
    spread_ResaltarCelda sprGB(0), 5, sprGB(0).MaxRows
    spread_ResaltarCelda sprGB(0), 6, sprGB(0).MaxRows
    spread_ResaltarCelda sprGC(0), 5, sprGC(0).MaxRows
    spread_ResaltarCelda sprGC(0), 6, sprGC(0).MaxRows

End If

End If

ErrGLC:

    frmVsDia.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub


Private Sub L_LlenarGrillasTicket(col As Integer)
Dim FD As Date
Dim SumImp As Double, fila As Integer
Dim fch As Variant

On Error GoTo ErrGLC:

frmVsDia.caption = Aplicacion.SeteoProceso(frmVsDia.caption)

fila = 1
If rs.RecordCount > 0 And sprTotal(1).MaxRows > 0 Then
rs.MoveFirst
Do While fila <= sprTotal(1).MaxRows
    
    sprTotal(1).GetText 1, fila, fch
    
    SumImp = 0
    If Not rs.EOF Then
    Do While Day(CDate(fch)) = Day(CDate(rs!fch_ticket))
        Select Case rs!Grupo
            Case "A"
                If sprGA(1).MaxRows > 1 Then
                    sprGA(1).SetText col, fila, str(rs!cant)
                    Spread_TotalesGrillaAcum sprGA(1), 3, 7, fila
                End If
            Case "B"
                If sprGB(1).MaxRows > 1 Then
                    sprGB(1).SetText col, fila, str(rs!cant)
                    Spread_TotalesGrillaAcum sprGB(1), 3, 7, fila
                End If
            Case "C"
                If sprGC(1).MaxRows > 1 Then
                    sprGC(1).SetText col, fila, str(rs!cant)
                    Spread_TotalesGrillaAcum sprGC(1), 3, 7, fila
                End If
        End Select
        SumImp = SumImp + rs!cant
        
        rs.MoveNext
        If rs.EOF Then
            Exit Do
        End If
    Loop
    sprTotal(1).SetText col, fila, str(SumImp)
    
    Spread_TotalesGrillaAcum sprTotal(1), 3, 7, fila
        
    spread_ResaltarCelda sprTotal(1), 5, fila
    spread_ResaltarCelda sprGA(1), 5, fila
    spread_ResaltarCelda sprGB(1), 5, fila
    spread_ResaltarCelda sprGC(1), 5, fila
    spread_ResaltarCelda sprTotal(1), 6, fila
    spread_ResaltarCelda sprGA(1), 6, fila
    spread_ResaltarCelda sprGB(1), 6, fila
    spread_ResaltarCelda sprGC(1), 6, fila
    spread_ResaltarCelda sprTotal(1), 9, fila
    spread_ResaltarCelda sprGA(1), 9, fila
    spread_ResaltarCelda sprGB(1), 9, fila
    spread_ResaltarCelda sprGC(1), 9, fila
    spread_ResaltarCelda sprTotal(1), 10, fila
    spread_ResaltarCelda sprGA(1), 10, fila
    spread_ResaltarCelda sprGB(1), 10, fila
    spread_ResaltarCelda sprGC(1), 10, fila
    
    fila = fila + 1
    Else
        Exit Do
    End If
Loop
    
If col = 3 Then
    Spread_TotalesGrillas sprTotal(1), sprTotal(1).MaxCols - 7, 2
    Spread_TotalesGrillas sprGA(1), sprGA(1).MaxCols - 7, 2
    Spread_TotalesGrillas sprGB(1), sprGB(1).MaxCols - 7, 2
    Spread_TotalesGrillas sprGC(1), sprGC(1).MaxCols - 7, 2
    
    spread_ResaltarCelda sprTotal(1), 5, sprTotal(1).MaxRows
    spread_ResaltarCelda sprTotal(1), 6, sprTotal(1).MaxRows
    spread_ResaltarCelda sprGA(1), 5, sprGA(1).MaxRows
    spread_ResaltarCelda sprGA(1), 6, sprGA(1).MaxRows
    spread_ResaltarCelda sprGB(1), 5, sprGB(1).MaxRows
    spread_ResaltarCelda sprGB(1), 6, sprGB(1).MaxRows
    spread_ResaltarCelda sprGC(1), 5, sprGC(1).MaxRows
    spread_ResaltarCelda sprGC(1), 6, sprGC(1).MaxRows
End If

End If

ErrGLC:

    frmVsDia.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub



Private Sub L_LimpiarGrillas()
Dim i

For i = 0 To 2
 sprTotal(i).MaxRows = 0
 sprGB(i).MaxRows = 0
 sprGA(i).MaxRows = 0
 sprGC(i).MaxRows = 0
Next
 sprTotal(3).MaxRows = 0
 sprTotal(4).MaxRows = 0
 
End Sub




Private Function L_NombreDato(i As Integer) As String
Select Case i
    Case 1
        L_NombreDato = "Importes"
    Case 2
        L_NombreDato = "Tickets"
    Case 3
        L_NombreDato = "Pasajeros Transitados"
    Case 4
        L_NombreDato = "Promedios por Ticket"
    Case 5
        L_NombreDato = "Promedios por Pasajeros Transitados"
End Select
End Function

Private Sub L_SetearNinel0()
Dim i

For i = 0 To cboModImp.ListCount - 1

Next

End Sub

Private Sub L_TratarPromedios(sprS As Control, sprD As Control, sprR As Control)
Dim fila As Integer
Dim fch As Variant
Dim i

If sprS.MaxRows = sprD.MaxRows Then
fila = 1
Do While fila <= sprS.MaxRows
    sprR.MaxRows = sprR.MaxRows + 1
    sprS.GetText 1, fila, fch
    sprR.SetText 1, sprR.MaxRows, Format$(fch, "dd-mm-yy")
    fila = fila + 1
Loop
Spread.Func_PromediosCol sprS, sprD, sprR, 3, 1, sprS.MaxRows
Spread.Func_PromediosCol sprS, sprD, sprR, 4, 1, sprS.MaxRows

Spread.Spread_TotalesLinea sprR

For fila = 1 To sprR.MaxRows
    spread_ResaltarCelda sprR, 5, fila
    spread_ResaltarCelda sprR, 6, fila
    'spread_ResaltarCelda sprR, 9, fila
    'spread_ResaltarCelda sprR, 10, fila
Next
End If

End Sub

Private Sub botEjecutar_Click(Index As Integer)

Select Case Index
    Case 0
        siImp = False
        If optMod(1).Value Then
            L_EjecutarReal mskAnio.Text
            If cboModImp.Text <> "" Then
               L_RefrescarEstim "I", cboModImp.Text, 0
               L_LlenarGrillasImp 3
            End If
            If cboModTicket.Text <> "" Then
               L_RefrescarEstim "T", cboModTicket.Text, 1
               L_LlenarGrillasTicket 3
               L_TratarPromedios sprTotal(0), sprTotal(1), sprTotal(3)
               L_TratarPromedios sprGA(0), sprGA(1), sprGA(2)
               L_TratarPromedios sprGB(0), sprGB(1), sprGB(2)
               L_TratarPromedios sprGC(0), sprGC(1), sprGC(2)
            End If
            If cboModPax.Text <> "" Then
                L_RefrescarEstim "P", cboModPax.Text, 2
                L_LlenarGrillasPax
                L_TratarPromedios sprTotal(0), sprTotal(2), sprTotal(4)
            End If
        Else
            If cboAnio.Text <> "" Then
            End If
        End If
    Case 1
        frCab.Enabled = True
        botEjecutar(0).Enabled = True
        tabEst.Enabled = False
        L_LimpiarGrillas
    Case 2
        Unload Me
End Select


End Sub

Private Sub L_EjecutarReal(anio As Integer)
Dim Sql As String
Dim sqlP As String


'On Error GoTo ErrGLC:

frmVsDia.caption = Aplicacion.SeteoProceso(frmVsDia.caption)

Sql = " SELECT "
Sql = Sql & " fch_ticket,"
Sql = Sql & " grupo_venta grupo,"
Sql = Sql & " sum(importe) imp, "
Sql = Sql & " sum(cantidad) cant "
Sql = Sql & "FROM " & funcLocal_Vista("venta_lgi", anio)
Sql = Sql & L_Armarcondicion(anio, "")
Sql = Sql & " group by fch_ticket,grupo_venta "
Sql = Sql & " order by fch_ticket,grupo_venta "

sqlP = " SELECT "
sqlP = sqlP & "fch_vuelo dia,"
sqlP = sqlP & " SUM(pax_vol) CANT_V "
'sqlP = sqlP & " From estadis.pax_volados "
sqlP = sqlP & " From estadis.pax_at_vs_vol "
sqlP = sqlP & L_Armarcondicion(anio, "X")
sqlP = sqlP & " GROUP BY fch_vuelo"

Call Aplicacion.ObtenerRsDAO(Sql, rs)
Call Aplicacion.ObtenerRsDAO(sqlP, rsPax)
    
tabEst.Enabled = True
botEjecutar(0).Enabled = False
frCab.Enabled = False

ErrGLC:
    frmVsDia.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub
Private Function L_Armarcondicion(anio As Integer, Tipo As String)
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String

fechaDesde = func_Dia1SegunMes_Anio(mskMes.Text, anio)
If mskMes.Text = Month(Date) Then
    fechaHasta = Format$(Date, FTOFECHA)
Else
    fechaHasta = func_Dia30SegunMes_Anio(mskMes.Text, anio)
End If
If Tipo = "X" Then
    Cond = " WHERE fch_vuelo between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)
Else
    Cond = " WHERE comitente = 'T' and fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)
End If

If CboCodAeropuerto.Text <> "" Then
    Cond = Cond & " and cod_depn = '" & CboCodAeropuerto.Text & "'"
End If

If CboEspigon.Text <> "" Then
    Cond = Cond & " and cod_sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
End If

L_Armarcondicion = Cond

End Function



Private Sub L_RefrescarEstim(TPorc As String, Modelo As String, indSpr As Integer)
Dim Sql As String, sqlT As String

On Error GoTo ErrEstim:

frmVsDia.caption = Aplicacion.SeteoProceso(frmVsDia.caption)

Sql = "Select dia, "
Sql = Sql & " grupo,"
Sql = Sql & " porc "
Sql = Sql & " From estadis.porcentaje_dg "
Sql = Sql & L_ArmarcondicionEstim(TPorc, Modelo)
Sql = Sql & " order by dia,grupo "

sqlT = "Select dia, "
sqlT = sqlT & " porc "
sqlT = sqlT & " From estadis.porcentaje_d "
sqlT = sqlT & L_ArmarcondicionEstim(TPorc, Modelo)
sqlT = sqlT & " order by dia "

If Aplicacion.ObtenerRsDAO(Sql, rsEstim) And Aplicacion.ObtenerRsDAO(sqlT, rsEstimT) Then
    
    L_LlenarGrillasEstim indSpr
    L_LlenarGrillasEstimT indSpr
    
    Aplicacion.CerrarDAO rsEstim
    Aplicacion.CerrarDAO rsEstimT

End If

ErrEstim:
    frmVsDia.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub


Private Sub L_LlenarGrillasEstim(indexSpr As Integer)
Dim FD As Date
Dim Sql As String


Do While Not rsEstim.EOF
    FD = rsEstim!dia

    Do While FD = CDate(rsEstim!dia)
        Select Case rsEstim!Grupo
            Case "A"
                sprGA(indexSpr).MaxRows = sprGA(indexSpr).MaxRows + 1
                sprGA(indexSpr).SetText 1, sprGA(indexSpr).MaxRows, Format$(FD, "dd-mm-yy")
                sprGA(indexSpr).SetText 4, sprGA(indexSpr).MaxRows, str(rsEstim!Porc)
                
                If rsEstim!Porc > 0 Then
                     Spread_TotalesGrillaAcum sprGA(indexSpr), 4, 8, sprGA(indexSpr).MaxRows
                End If
            Case "B"
                sprGB(indexSpr).MaxRows = sprGB(indexSpr).MaxRows + 1
                sprGB(indexSpr).SetText 1, sprGB(indexSpr).MaxRows, Format$(FD, "dd-mm-yy")
                sprGB(indexSpr).SetText 4, sprGB(indexSpr).MaxRows, str(rsEstim!Porc)
                
                If rsEstim!Porc > 0 Then
                    Spread_TotalesGrillaAcum sprGB(indexSpr), 4, 8, sprGB(indexSpr).MaxRows
                End If
            
            Case "C"
                sprGC(indexSpr).MaxRows = sprGC(indexSpr).MaxRows + 1
                sprGC(indexSpr).SetText 1, sprGC(indexSpr).MaxRows, Format$(FD, "dd-mm-yy")
                sprGC(indexSpr).SetText 4, sprGC(indexSpr).MaxRows, str(rsEstim!Porc)
                
                If rsEstim!Porc > 0 Then
                    Spread_TotalesGrillaAcum sprGC(indexSpr), 4, 8, sprGC(indexSpr).MaxRows
                End If
                
        End Select
        rsEstim.MoveNext
        If rsEstim.EOF Then
            Exit Do
        End If
    Loop
    If FD = Date - 1 Then
        Exit Do
    End If
Loop
   Spread_PintarfinSemana sprGA(indexSpr)
   Spread_PintarfinSemana sprGB(indexSpr)
   Spread_PintarfinSemana sprGC(indexSpr)

End Sub

Private Sub L_LlenarGrillasEstimT(indexSpr As Integer)
Dim FD As Date
Dim rsGr As Recordset
Dim Sql As String
Dim tot As Double

Do While Not rsEstimT.EOF
    FD = rsEstimT!dia
    sprTotal(indexSpr).MaxRows = sprTotal(indexSpr).MaxRows + 1
    sprTotal(indexSpr).SetText 1, sprTotal(indexSpr).MaxRows, Format$(rsEstimT!dia, "dd-mm-yy")
    
    sprTotal(indexSpr).SetText 4, sprTotal(indexSpr).MaxRows, str(rsEstimT!Porc)
    
    If rsEstimT!Porc > 0 Then
        Spread_TotalesGrillaAcum sprTotal(indexSpr), 4, 8, sprTotal(indexSpr).MaxRows
    End If
                
    rsEstimT.MoveNext

    If FD = Date - 1 Then
        Exit Do
    End If

Loop
Spread_PintarfinSemana sprTotal(indexSpr)

End Sub

Private Sub L_LlenarGrillasPax()
Dim FD As Date
Dim Sql As String
Dim fila As Integer

On Error GoTo ErrP:

frmVsDia.caption = Aplicacion.SeteoProceso(frmVsDia.caption)

fila = 1
Do While Not rsPax.EOF And sprTotal(2).MaxRows > 0
    FD = rsPax!dia
    
    sprTotal(2).SetText 3, fila, str(rsPax!cant_v)
    
    If rsPax!cant_v > 0 Then
        Spread_TotalesGrillaAcum sprTotal(2), 3, 7, fila
    End If
                
    rsPax.MoveNext
        
    spread_ResaltarCelda sprTotal(2), 5, fila
    spread_ResaltarCelda sprTotal(2), 6, fila
    spread_ResaltarCelda sprTotal(2), 9, fila
    spread_ResaltarCelda sprTotal(2), 10, fila

    fila = fila + 1
    If FD = Date - 1 Then
        Exit Do
    End If
Loop
    
    Spread_TotalesGrillas sprTotal(2), sprTotal(2).MaxCols - 7, 2
    spread_ResaltarCelda sprTotal(2), 5, sprTotal(2).MaxRows
    spread_ResaltarCelda sprTotal(2), 6, sprTotal(2).MaxRows
    spread_ResaltarCelda sprGA(2), 5, sprGA(2).MaxRows
    spread_ResaltarCelda sprGA(2), 6, sprGA(2).MaxRows
    spread_ResaltarCelda sprGB(2), 5, sprGB(2).MaxRows
    spread_ResaltarCelda sprGB(2), 6, sprGB(2).MaxRows
    spread_ResaltarCelda sprGC(2), 5, sprGC(2).MaxRows
    spread_ResaltarCelda sprGC(2), 6, sprGC(2).MaxRows

ErrP:

    frmVsDia.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub



Private Function L_ArmarcondicionEstim(TPorc As String, Modelo As String)
Dim Cond

Cond = " WHERE anio = " & mskAnio.Text
Cond = Cond & " And Mes = " & mskMes.Text
Cond = Cond & " And Tipo_porc = '" & TPorc & "'"
Cond = Cond & " And Descrip = '" & Modelo & "'"

If CboCodAeropuerto.Text <> "" Then
    Cond = Cond & " and depn = '" & CboCodAeropuerto.Text & "'"
End If

If CboEspigon.Text <> "" Then
    Cond = Cond & " and sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
End If


L_ArmarcondicionEstim = Cond

End Function

Private Sub botExcel_Click()


Select Case tabEst.Tab
    Case 0
        L_TratarExcelTot
    Case 1
    Case 2
    Case 3
    
End Select

End Sub



Private Sub L_TratarExcelTot()
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim fila As Integer
Dim i As Integer
Dim tit As Variant

Dim nombre As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

frmVsDia.caption = Aplicacion.SeteoProceso(frmVsDia.caption)

If nombre <> "" Then

    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.Application.Visible = True
    
    ReDim titCol(sprTotal(0).MaxCols)
    col = 1
    fila = 3
    
    For i = 1 To sprTotal(0).MaxCols
        sprTotal(0).GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 1, 1, "INFORME GENERAL DE COMPARATIVO VENTA-ESTIMADO"
    rango = Exl_rangos(1, 1, 1, sprTotal(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Blanco
    
    Exl_PonerValor AppExcel, fila, col, CboCodAeropuerto.Text & " - " & CboEspigon.Text
    rango = Exl_rangos(fila, fila, 1, sprTotal(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    '---------------------------------------
For i = 0 To 4
    
    If sprTotal(i).MaxRows > 0 Then
    
    fila = fila + 2
    
    Exl_PonerValor AppExcel, fila, col, L_NombreDato(i + 1)
    rango = Exl_rangos(fila, fila, 1, sprTotal(i).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, Exl_Blanco
    AppExcel.Application.Range(rango).Merge
        
    fila = fila + 1
    
    Exl_BajarGrillaExel sprTotal(i), AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + sprTotal(i).MaxRows, col + 1, sprTotal(i).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprTotal(i).MaxRows
    rango = Exl_rangos(fila, fila, col, sprTotal(i).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    fila = fila + 1
'    rango = Exl_rangos(fila, fila, col, sprTotal(i).MaxCols)
'    Exl_ColorInt AppExcel, rango, Exl_
    
    rango = Exl_rangos(fila - sprTotal(i).MaxRows, fila - 1, 6, sprTotal(i).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
'    AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
    End If
    
Next
    Exl.Exl_AnchoCol AppExcel, sprTotal(1).MaxCols, sprTotal(1).MaxCols, 1
    Exl.Exl_AnchoCol AppExcel, 1, 1, 8
    Exl.Exl_AnchoCol AppExcel, 5, 5, 7
    Exl.Exl_AnchoCol AppExcel, 9, 9, 7
    
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    
    AppExcel.SaveAs nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmVsDia.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub

Private Sub L_TratarExcelGA()
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim fila As Integer
Dim i As Integer
Dim tit As Variant

Dim nombre As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

frmVsDia.caption = Aplicacion.SeteoProceso(frmIndic.caption)

If nombre <> "" Then

    Set AppExcel = CreateObject("excel.sheet")
    
    AppExcel.Application.Visible = True
    
    ReDim titCol(sprTotal(0).MaxCols)
    col = 1
    fila = 3
    
    For i = 1 To sprTotal(0).MaxCols
        sprTotal(0).GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 1, 1, "INFORME GENERAL DE COMPARATIVO VENTA-ESTIMADO"
    rango = Exl_rangos(1, 1, 1, sprTotal(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Blanco
    
    Exl_PonerValor AppExcel, fila, col, CboCodAeropuerto.Text & " - " & CboEspigon.Text
    rango = Exl_rangos(fila, fila, 1, sprTotal(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    '---------------------------------------
For i = 0 To 4
    
    If sprTotal(i).MaxRows > 0 Then
    
    fila = fila + 2
    
    Exl_PonerValor AppExcel, fila, col, L_NombreDato(i + 1)
    rango = Exl_rangos(fila, fila, 1, sprTotal(i).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, Exl_Cel
    AppExcel.Application.Range(rango).Merge
        
    fila = fila + 1
    
    Exl_BajarGrillaExel sprTotal(i), AppExcel, fila, col, titCol
    fila = fila + sprTotal(i).MaxRows + 1
    rango = Exl_rangos(fila, fila, col, sprTotal(i).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Cel
    
    rango = Exl_rangos(fila - sprTotal(i).MaxRows, fila - 1, 6, sprTotal(i).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    End If
    
Next
    Exl.Exl_AnchoCol AppExcel, sprTotal(1).MaxCols, sprTotal(1).MaxCols, 1
    Exl.Exl_AnchoCol AppExcel, 1, 1, 8
    Exl.Exl_AnchoCol AppExcel, 5, 5, 7
    Exl.Exl_AnchoCol AppExcel, 9, 9, 7
    
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    
    AppExcel.SaveAs nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmVsDia.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub CboCodAeropuerto_Click()
Dim Sql As String

Sql = " SELECT cod_sdep,descrip FROM baires.subdependencia "
Sql = Sql & " WHERE cod_depn = '" & CboCodAeropuerto.Text & "'"
Sql = Sql & " ORDER BY cod_sdep"
 
FuncCbos_LlenarCboLst CboEspigon, LstEspigon, Sql

cboModImp.Clear
cboModTicket.Clear
cboModPax.Clear

End Sub


Private Sub CboCodAeropuerto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    CboCodAeropuerto.ListIndex = -1
End If
End Sub


Private Sub CboEspigon_Click()
Dim Sql As String
Dim rs As Recordset

Sql = "Select tipo_porc,descrip,nivel from estadis.porciento_cabezera "
Sql = Sql & " Where anio= " & mskAnio.Text
Sql = Sql & " And mes= " & txtMesMod.Text
Sql = Sql & " And cod_depn= '" & CboCodAeropuerto.Text & "'"
Sql = Sql & " And cod_sdep= '" & LstEspigon.List(CboEspigon.ListIndex) & "'"

cboModImp.Clear
cboModTicket.Clear
cboModPax.Clear
If Aplicacion.ObtenerRsDAO(Sql, rs) Then
    Do While Not rs.EOF
        Select Case rs!tipo_porc
            Case "I"
                cboModImp.AddItem rs!Descrip
                If rs!Nivel = 0 Then
                    cboModImp.ListIndex = cboModImp.ListCount - 1
                End If
            Case "T"
                cboModTicket.AddItem rs!Descrip
                If rs!Nivel = 0 Then
                    cboModTicket.ListIndex = cboModTicket.ListCount - 1
                End If
            
            Case "P"
                cboModPax.AddItem rs!Descrip
                If rs!Nivel = 0 Then
                    cboModPax.ListIndex = cboModPax.ListCount - 1
                End If
        
        End Select
        
        rs.MoveNext
    Loop
End If

'L_SetearNinel0

End Sub

Private Sub CboEspigon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    CboEspigon.ListIndex = -1
End If
End Sub


Private Sub cboModImp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    cboModImp.ListIndex = -1
End If

End Sub


Private Sub cboModPax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    cboModPax.ListIndex = -1
End If
End Sub


Private Sub cboModTicket_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    cboModTicket.ListIndex = -1
End If
End Sub


Private Sub Form_Load()
Dim Sql As String

Top = 30
Left = 250
Height = 6300
Width = 9000

Sql = " SELECT cod_depn,descrip FROM baires.dependencia "
Sql = Sql & " ORDER BY cod_depn"

FuncCbos_LlenarCbo CboCodAeropuerto, Sql

mskAnio.Text = Year(Date)
mskMes.Text = Month(Date)

txtMesMod.Text = Month(Date)
L_LimpiarGrillas
frmPrincipal.lstForms.AddItem "frmVsDia"

End Sub


Private Sub Form_Unload(Cancel As Integer)
    FuncLocal_SacarForm "frmVsDia"
End Sub


Private Sub mskAnio_Change()
CboCodAeropuerto.ListIndex = -1
End Sub

Private Sub mskAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub mskAnio_LostFocus()
If Val(mskAnio.Text) < 1996 Or Val(mskAnio) > 2050 Then
    mskAnio.Text = Year(Date)
End If

End Sub


Private Sub mskMes_Change()
CboCodAeropuerto.ListIndex = -1
End Sub

Private Sub mskMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub mskMes_LostFocus()
If Val(mskMes.Text) < 1 Or Val(mskMes.Text) > 12 Then
    mskMes.Text = Month(Date)
End If
txtMesMod.Text = mskMes.Text
End Sub


Private Sub tabEst_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    Select Case tabEst.Tab
        Case 0
            sprTotal(tabTotal.Tab).SetFocus
        Case 1
            sprGA(tabGA.Tab).SetFocus
        Case 2
            sprGB(tabGB.Tab).SetFocus
        Case 3
            sprGC(tabGC.Tab).SetFocus
    End Select
    
    
ErrT:
    Exit Sub

End Sub

Private Sub tabGA_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    sprGA(tabGA.Tab).SetFocus
    
ErrT:
    Exit Sub

End Sub

Private Sub tabGB_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    sprGB(tabGB.Tab).SetFocus
    
ErrT:
    Exit Sub

End Sub

Private Sub tabGC_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    sprGC(tabGC.Tab).SetFocus
    
ErrT:
    Exit Sub

End Sub

Private Sub tabTotal_Click(PreviousTab As Integer)

On Error GoTo ErrT:

    sprTotal(tabTotal.Tab).SetFocus
    
ErrT:
    Exit Sub
End Sub

