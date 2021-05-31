VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMonitoreoProducto 
   Caption         =   "Estimados vs Ventas-Tickets-Pasajeros"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   10875
   Begin FPSpread.vaSpread sprEXEL 
      Height          =   1365
      Left            =   7740
      OleObjectBlob   =   "frmMonitoreoProducto.frx":0000
      TabIndex        =   40
      Top             =   4005
      Visible         =   0   'False
      Width           =   2760
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
      Height          =   510
      Index           =   2
      Left            =   9345
      Picture         =   "frmMonitoreoProducto.frx":06CE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   645
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
      Left            =   9345
      Picture         =   "frmMonitoreoProducto.frx":0EF0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   165
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
      Left            =   10020
      Picture         =   "frmMonitoreoProducto.frx":0FF2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   165
      Width           =   615
   End
   Begin VB.CommandButton botExcel 
      Caption         =   "Excel"
      Height          =   525
      Left            =   10035
      Picture         =   "frmMonitoreoProducto.frx":10F4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   645
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
      Height          =   1395
      Left            =   45
      TabIndex        =   0
      Top             =   -30
      Width           =   10755
      Begin VB.OptionButton optCota 
         Caption         =   "Todos"
         Height          =   210
         Index           =   2
         Left            =   7935
         TabIndex        =   54
         Top             =   1080
         Value           =   -1  'True
         Width           =   1110
      End
      Begin VB.OptionButton optCota 
         Caption         =   "Por Arriba"
         Height          =   210
         Index           =   1
         Left            =   6765
         TabIndex        =   53
         Top             =   1080
         Width           =   1110
      End
      Begin VB.OptionButton optCota 
         Caption         =   "Por Debajo"
         Height          =   210
         Index           =   0
         Left            =   5400
         TabIndex        =   52
         Top             =   1065
         Width           =   1110
      End
      Begin VB.ListBox lstProv 
         Height          =   255
         Left            =   6720
         TabIndex        =   43
         Top             =   660
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox cboProv 
         Height          =   315
         Left            =   6780
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   600
         Width           =   2415
      End
      Begin VB.ListBox lstFamilia 
         Height          =   255
         Left            =   6735
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox cboFamilia 
         Height          =   315
         Left            =   6810
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   210
         Width           =   2370
      End
      Begin VB.ListBox LstEspigon 
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   645
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox CboRubro 
         Height          =   315
         ItemData        =   "frmMonitoreoProducto.frx":1686
         Left            =   3330
         List            =   "frmMonitoreoProducto.frx":1688
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   195
         Width           =   2415
      End
      Begin VB.ComboBox CboSubRubro 
         Height          =   315
         Left            =   3315
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskCota 
         Height          =   300
         Left            =   4290
         TabIndex        =   23
         Top             =   1005
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PROVEEDOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   5790
         TabIndex        =   41
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COTA %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3300
         TabIndex        =   31
         Top             =   990
         Width           =   960
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FAMILIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   5820
         TabIndex        =   20
         Top             =   210
         Width           =   960
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
         Caption         =   "RUBRO"
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
         Left            =   2085
         TabIndex        =   7
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label LblEspigon 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SUB RUBRO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2085
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab tabEst 
      Height          =   5115
      Left            =   15
      TabIndex        =   14
      Top             =   1395
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   9022
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   6
      TabHeight       =   441
      ForeColor       =   255
      TabCaption(0)   =   "TOTAL CIA"
      TabPicture(0)   =   "frmMonitoreoProducto.frx":168A
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tabTotal"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "EZEIZA"
      TabPicture(1)   =   "frmMonitoreoProducto.frx":16A6
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabGA"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "AEROPARQUE"
      TabPicture(2)   =   "frmMonitoreoProducto.frx":16C2
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab1"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "INTERIOR"
      TabPicture(3)   =   "frmMonitoreoProducto.frx":16DE
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTab2"
      Tab(3).Control(0).Enabled=   0   'False
      Begin TabDlg.SSTab tabTotal 
         Height          =   4590
         Left            =   180
         TabIndex        =   15
         Top             =   360
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   8096
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
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
         TabCaption(0)   =   "REAL MES"
         TabPicture(0)   =   "frmMonitoreoProducto.frx":16FA
         Tab(0).ControlCount=   4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblProd(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblSTKBOD(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "sprReal(0)"
         Tab(0).Control(3).Enabled=   0   'False
         TabCaption(1)   =   "PROYECTADO"
         TabPicture(1)   =   "frmMonitoreoProducto.frx":1716
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprProy(0)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lblProdPy(0)"
         Tab(1).Control(1).Enabled=   0   'False
         Begin FPSpread.vaSpread sprReal 
            Height          =   3615
            Index           =   0
            Left            =   105
            OleObjectBlob   =   "frmMonitoreoProducto.frx":1732
            TabIndex        =   16
            Top             =   90
            Width           =   10290
         End
         Begin FPSpread.vaSpread sprProy 
            Height          =   3615
            Index           =   0
            Left            =   -74925
            OleObjectBlob   =   "frmMonitoreoProducto.frx":1DB3
            TabIndex        =   32
            Top             =   75
            Width           =   10290
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Stock Bodega"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            TabIndex        =   45
            Top             =   3795
            Width           =   1395
         End
         Begin VB.Label lblSTKBOD 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   7545
            TabIndex        =   44
            Top             =   3780
            Width           =   2580
         End
         Begin VB.Label lblProdPy 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   -74925
            TabIndex        =   36
            Top             =   3765
            Width           =   5670
         End
         Begin VB.Label lblProd 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   105
            TabIndex        =   27
            Top             =   3780
            Width           =   5670
         End
      End
      Begin TabDlg.SSTab tabGA 
         Height          =   4605
         Left            =   -74775
         TabIndex        =   17
         Top             =   390
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   8123
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
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
         TabCaption(0)   =   "REAL MES"
         TabPicture(0)   =   "frmMonitoreoProducto.frx":2434
         Tab(0).ControlCount=   4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblProd(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblSTKBOD(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "sprReal(1)"
         Tab(0).Control(3).Enabled=   0   'False
         TabCaption(1)   =   "PROYECTADO"
         TabPicture(1)   =   "frmMonitoreoProducto.frx":2450
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprProy(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lblProdPy(1)"
         Tab(1).Control(1).Enabled=   0   'False
         Begin FPSpread.vaSpread sprReal 
            Height          =   3615
            Index           =   1
            Left            =   60
            OleObjectBlob   =   "frmMonitoreoProducto.frx":246C
            TabIndex        =   24
            Top             =   90
            Width           =   10290
         End
         Begin FPSpread.vaSpread sprProy 
            Height          =   3615
            Index           =   1
            Left            =   -74955
            OleObjectBlob   =   "frmMonitoreoProducto.frx":2AED
            TabIndex        =   33
            Top             =   75
            Width           =   10290
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Stock Bodega"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6030
            TabIndex        =   49
            Top             =   3795
            Width           =   1395
         End
         Begin VB.Label lblSTKBOD 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   1
            Left            =   7455
            TabIndex        =   46
            Top             =   3780
            Width           =   2580
         End
         Begin VB.Label lblProdPy 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   1
            Left            =   -74955
            TabIndex        =   37
            Top             =   3765
            Width           =   5670
         End
         Begin VB.Label lblProd 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   1
            Left            =   60
            TabIndex        =   28
            Top             =   3795
            Width           =   5490
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4605
         Left            =   -74805
         TabIndex        =   18
         Top             =   360
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   8123
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
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
         TabCaption(0)   =   "REAL MES"
         TabPicture(0)   =   "frmMonitoreoProducto.frx":316E
         Tab(0).ControlCount=   4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblProd(2)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblSTKBOD(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label5"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "sprReal(2)"
         Tab(0).Control(3).Enabled=   0   'False
         TabCaption(1)   =   "PROYECTADO"
         TabPicture(1)   =   "frmMonitoreoProducto.frx":318A
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprProy(2)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lblProdPy(2)"
         Tab(1).Control(1).Enabled=   0   'False
         Begin FPSpread.vaSpread sprReal 
            Height          =   3615
            Index           =   2
            Left            =   60
            OleObjectBlob   =   "frmMonitoreoProducto.frx":31A6
            TabIndex        =   25
            Top             =   75
            Width           =   10290
         End
         Begin FPSpread.vaSpread sprProy 
            Height          =   3615
            Index           =   2
            Left            =   -74955
            OleObjectBlob   =   "frmMonitoreoProducto.frx":3827
            TabIndex        =   34
            Top             =   75
            Width           =   10290
         End
         Begin VB.Label Label5 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Stock Bodega"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            TabIndex        =   50
            Top             =   3780
            Width           =   1395
         End
         Begin VB.Label lblSTKBOD 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   2
            Left            =   7545
            TabIndex        =   47
            Top             =   3765
            Width           =   2580
         End
         Begin VB.Label lblProdPy 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   2
            Left            =   -74925
            TabIndex        =   38
            Top             =   3765
            Width           =   5670
         End
         Begin VB.Label lblProd 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   2
            Left            =   60
            TabIndex        =   29
            Top             =   3780
            Width           =   5490
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4620
         Left            =   -74850
         TabIndex        =   19
         Top             =   345
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   8149
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
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
         TabCaption(0)   =   "REAL MES"
         TabPicture(0)   =   "frmMonitoreoProducto.frx":3EA8
         Tab(0).ControlCount=   4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblProd(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblSTKBOD(3)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label6"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "sprReal(3)"
         Tab(0).Control(3).Enabled=   0   'False
         TabCaption(1)   =   "PROYECTADO"
         TabPicture(1)   =   "frmMonitoreoProducto.frx":3EC4
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprProy(3)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lblProdPy(3)"
         Tab(1).Control(1).Enabled=   0   'False
         Begin FPSpread.vaSpread sprReal 
            Height          =   3615
            Index           =   3
            Left            =   60
            OleObjectBlob   =   "frmMonitoreoProducto.frx":3EE0
            TabIndex        =   26
            Top             =   60
            Width           =   10290
         End
         Begin FPSpread.vaSpread sprProy 
            Height          =   3615
            Index           =   3
            Left            =   -74940
            OleObjectBlob   =   "frmMonitoreoProducto.frx":4561
            TabIndex        =   35
            Top             =   75
            Width           =   10290
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Stock Bodega"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6135
            TabIndex        =   51
            Top             =   3795
            Width           =   1395
         End
         Begin VB.Label lblSTKBOD 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   3
            Left            =   7560
            TabIndex        =   48
            Top             =   3765
            Width           =   2580
         End
         Begin VB.Label lblProdPy 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   3
            Left            =   -74940
            TabIndex        =   39
            Top             =   3765
            Width           =   5670
         End
         Begin VB.Label lblProd 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
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
            Index           =   3
            Left            =   75
            TabIndex        =   30
            Top             =   3765
            Width           =   5490
         End
      End
   End
End
Attribute VB_Name = "frmMonitoreoProducto"
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

Private Sub L_LimpiarGrillas()
Dim i

For i = 0 To 3
 sprReal(i).MaxRows = 0
 sprProy(i).MaxRows = 0
 lblProd(i).caption = ""
 lblSTKBOD(i).caption = ""
Next
 
End Sub




Private Function L_PasarGrillaExel(spr As control) As String
Dim i, j
Dim valor As Variant
Dim desc As String, sql As String

sprEXEL.MaxRows = 0

For i = 1 To spr.MaxRows
    sprEXEL.MaxRows = sprEXEL.MaxRows + 1
    
    For j = 1 To 9
    
      If j = 2 Then
          spr.GetText j, i, valor
          sql = "Select descrip From baires.producto where cod_prod = " & valor
          Func.Func_ObtenerDesc sql, desc
          sprEXEL.SetText j, i, Trim(valor) & " - " & desc
      ElseIf j > 5 Then
         spr.GetText j, i, valor
         sprEXEL.SetText j + 1, i, Trim(valor)
      Else
         spr.GetText j, i, valor
         sprEXEL.SetText j, i, Trim(valor)
      End If
    Next
Next

End Function



Private Sub botEjecutar_Click(Index As Integer)

Select Case Index
    Case 0
        L_Refrescar
        L_RefrescarProy
    Case 1
        frCab.Enabled = True
        botEjecutar(0).Enabled = True
        tabEst.Enabled = False
        L_LimpiarGrillas
    Case 2
        Unload Me
End Select


End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlP As String


'On Error GoTo ErrGLC:

frmMonitoreoProducto.caption = Aplicacion.SeteoProceso(frmMonitoreoProducto.caption)


sql = "  select depn ,cod_prov "
sql = sql & " ,sum(Estimado_U) Est_U,sum(Estimado_$) Est_P "
sql = sql & " ,sum(Venta_U) Vta_U ,sum(Venta_$) Vta_P "
sql = sql & " From "
sql = sql & " ( select depn,sdep,p.cod_prov,"
sql = sql & "   ROUND(sum(nvl(estm_ult,0)*porc/100)) Estimado_U, "
sql = sql & "   ROUND(sum(nvl(estm_ult,0) * p.precio_ppal * porc /100 ),2) Estimado_$"
sql = sql & "   ,0 Venta_U ,0 Venta_$ "
sql = sql & "   From estimado e, producto p,"
sql = sql & "  (Select depn,sdep, sum(trunc(porcdia*porcesp/100,6)) porc"
sql = sql & "   From estadis.porcentaje_d "
sql = sql & "   WHERE anio = to_number(to_char(sysdate,'YYYY')) And Mes = to_number(to_char(sysdate,'MM')) "
sql = sql & "   And Tipo_porc = 'I'  and sdep <> 'INTA' "
sql = sql & "   And dia between  first_day(sysdate) and sysdate - 1 "
sql = sql & "   Group by depn,sdep ) g "
sql = sql & "   Where P.Cod_prod = e.Cod_prod and aniomes = to_number(to_char(sysdate,'YYYYMM')) "

sql = sql & "   and p.cod_prod in (select cod_prod from monitoreo_productos)"

sql = sql & "   " & L_Armarcondicion
sql = sql & "   Group by depn,sdep,cod_prov "
sql = sql & "   Union"
sql = sql & "   SELECT v.cod_depn,cod_ssdep,cod_prov"
sql = sql & "   ,0 Estimado_U,0 Estimado_$"
sql = sql & "   ,sum(cantidad),sum(importe) "
sql = sql & "   From venta_plg v, ventas.apertura_sdep s"
sql = sql & "   Where fch_ticket between first_day(sysdate) and sysdate-1 "
sql = sql & "   And v.cod_depn = s.cod_depn and v.cod_sdep = s.cod_sdep and v.cod_local = s.cod_local"

sql = sql & "   and cod_prod in (select cod_prod from monitoreo_productos)"

sql = sql & "   " & L_Armarcondicion
sql = sql & "   group by v.cod_depn,cod_ssdep,cod_prov ) "
sql = sql & "   group by depn,cod_prov "
sql = sql & "   Union "
sql = sql & "   Select 'CIA' depn ,cod_prov "
sql = sql & "   ,sum(Estimado_U) Est_U,sum(Estimado_$) Est_$ "
sql = sql & "   ,sum(Venta_U) Vta_U ,sum(Venta_$) Vta_$"
sql = sql & " From ( "
sql = sql & "  select p.cod_prov, ROUND(sum(nvl(estm_ult,0)*porc/100)) Estimado_U, "
sql = sql & "  ROUND(sum(nvl(estm_ult,0) * p.precio_ppal * porc /100 ),2) Estimado_$ "
sql = sql & " ,0 Venta_U ,0 Venta_$ "
sql = sql & " From estimado e, producto p, "
sql = sql & " (Select sum(trunc(porcdia*porcesp/100,6)) porc "
sql = sql & " From estadis.porcentaje_d "
sql = sql & " WHERE anio = to_number(to_char(sysdate,'YYYY')) And Mes = to_number(to_char(sysdate,'MM')) "
sql = sql & " And Tipo_porc = 'I'  and sdep <> 'INTA' "
sql = sql & " And dia between  first_day(sysdate) and sysdate - 1 ) g "
sql = sql & " Where P.Cod_prod = e.Cod_prod "
sql = sql & " And aniomes = to_number(to_char(sysdate,'YYYYMM')) "

sql = sql & "   and p.cod_prod in (select cod_prod from monitoreo_productos)"

sql = sql & " " & L_Armarcondicion
sql = sql & " Group by cod_prov "
sql = sql & " Union "
sql = sql & " SELECT cod_prov ,0 Estimado_U,0 Estimado_$ ,sum(cantidad),sum(importe) "
sql = sql & " From venta_plg v "
sql = sql & " Where fch_ticket between first_day(sysdate) and sysdate-1 "

sql = sql & "   and cod_prod in (select cod_prod from monitoreo_productos)"

sql = sql & " " & L_Armarcondicion
sql = sql & " Group by cod_prov ) "
sql = sql & " Group by cod_prov "
sql = sql & " "


sql = "  select depn ,pv.cod_prov,pv.descrip,p.cod_prod, "
sql = sql & " sum(ESTIMADO_U) EST_U ,"
sql = sql & " sum(ESTIMADO_$) EST_P,"
sql = sql & " sum(VENTA_U) VTA_U,"
sql = sql & " sum(VENTA_$) VTA_P,"
sql = sql & " sum(proyectad_u) PR_U,"
sql = sql & " sum(proyectado_$) PR_P,"
sql = sql & " Sum (stock_u) STK_U,"
sql = sql & " Sum (stock_$) STK_P"
sql = sql & " FROM  baires.proveedor pv, producto p,ESTADIS.MONITOREO_PRODUCTOS m   "
sql = sql & " Where P.Cod_prod = m.Cod_prod And p.cod_prov = pv.cod_prov "
sql = sql & "   " & L_Armarcondicion
sql = sql & " group by     DEPN, pv.cod_prov, pv.descrip,p.cod_prod "
    If optCota(2).Value Then
       sql = sql & " Having Abs((Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$))) >= " & mskCota.Text / 100
    ElseIf optCota(1).Value Then
        sql = sql & " Having (Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$)) >= " & mskCota.Text / 100
    ElseIf optCota(0).Value Then
        sql = sql & " Having (Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$)) <= - " & mskCota.Text / 100
    End If
sql = sql & " Union"
sql = sql & " Select 'CIA' depn ,pv.cod_prov,pv.descrip,p.cod_prod, "
sql = sql & " sum(ESTIMADO_U) EST_U ,"
sql = sql & " sum(ESTIMADO_$) EST_P,"
sql = sql & " sum(VENTA_U) VTA_U,"
sql = sql & " sum(VENTA_$) VTA_P,"
sql = sql & " sum(proyectad_u) PR_U,"
sql = sql & " sum(proyectado_$) PR_P,"
sql = sql & " Sum (stock_u) STK_U,"
sql = sql & " Sum (stock_$) STK_P"
sql = sql & " FROM  baires.proveedor pv , producto p,ESTADIS.MONITOREO_PRODUCTOS m   "
sql = sql & " Where P.Cod_prod = m.Cod_prod  And p.cod_prov = pv.cod_prov "
sql = sql & "   " & L_Armarcondicion
sql = sql & " group by  pv.cod_prov, pv.descrip, p.cod_prod "
    If optCota(2).Value Then
       sql = sql & " Having Abs((Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$))) >= " & mskCota.Text / 100
    ElseIf optCota(1).Value Then
        sql = sql & " Having (Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$)) >= " & mskCota.Text / 100
    ElseIf optCota(0).Value Then
        sql = sql & " Having  (Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$)) <= - " & mskCota.Text / 100
    End If

If Aplicacion.ObtenerRsDAO(sql, rs) Then
      L_LlenarGrillasReal
      tabEst.Enabled = True
      botEjecutar(0).Enabled = False
      frCab.Enabled = False
End If


ErrGLC:
    frmMonitoreoProducto.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub

Private Sub L_RefrescarProy()
Dim sql As String
Dim sqlP As String


'On Error GoTo ErrGLC:

frmMonitoreoProducto.caption = Aplicacion.SeteoProceso(frmMonitoreoProducto.caption)



sql = "  select depn ,pv.cod_prov,pv.descrip,p.cod_prod, "
sql = sql & " sum(ESTIMADO_U+proyectad_u) EST_U ,"
sql = sql & " sum(ESTIMADO_$+proyectado_$) EST_P,"
sql = sql & " sum(VENTA_U+proyectad_u) VTA_U,"
sql = sql & " sum(VENTA_$+proyectado_$) VTA_P,"
sql = sql & " sum(proyectad_u) PR_U,"
sql = sql & " sum(proyectado_$) PR_P,"
sql = sql & " Sum (stock_u) STK_U,"
sql = sql & " Sum (stock_$) STK_P"
sql = sql & " FROM  baires.proveedor pv, producto p,ESTADIS.MONITOREO_PRODUCTOS m   "
sql = sql & " Where P.Cod_prod = m.Cod_prod And p.cod_prov = pv.cod_prov "
sql = sql & "   " & L_Armarcondicion
sql = sql & " group by     DEPN, pv.cod_prov, pv.descrip,p.cod_prod "
sql = sql & " Union"
sql = sql & " Select 'CIA' depn ,pv.cod_prov,pv.descrip,p.cod_prod, "
sql = sql & " sum(ESTIMADO_U+proyectad_u) EST_U ,"
sql = sql & " sum(ESTIMADO_$+proyectado_$) EST_P,"
sql = sql & " sum(VENTA_U+proyectad_u) VTA_U,"
sql = sql & " sum(VENTA_$+proyectado_$) VTA_P,"
sql = sql & " sum(proyectad_u) PR_U,"
sql = sql & " sum(proyectado_$) PR_P,"
sql = sql & " Sum (stock_u) STK_U,"
sql = sql & " Sum (stock_$) STK_P"
sql = sql & " FROM  baires.proveedor pv , producto p,ESTADIS.MONITOREO_PRODUCTOS m   "
sql = sql & " Where P.Cod_prod = m.Cod_prod  And p.cod_prov = pv.cod_prov "
sql = sql & "   " & L_Armarcondicion
sql = sql & " group by  pv.cod_prov, pv.descrip,p.cod_prod "

If Aplicacion.ObtenerRsDAO(sql, rs) Then
      L_LlenarGrillasProy
End If


ErrGLC:
    frmMonitoreoProducto.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub

Private Function L_Armarcondicion()
Dim Cond

If cboRubro.Text <> "" Then
   Cond = Cond & " And cod_rubr = '" & cboRubro.Text & "'"
End If

If cboSubRubro.Text <> "" Then
   Cond = Cond & " And cod_srub = '" & LstEspigon.List(cboSubRubro.ListIndex) & "'"
End If

If cboFamilia.Text <> "" Then
   Cond = Cond & " And cod_familia = '" & lstFamilia.List(cboFamilia.ListIndex) & "'"
End If

If cboProv.Text <> "" Then
   Cond = Cond & " And pv.cod_prov = '" & lstProv.List(cboProv.ListIndex) & "'"
End If

L_Armarcondicion = Cond

End Function



Private Sub L_LlenarGrillasReal()

Do While Not rs.EOF
        Select Case rs!depn
            Case "AEP"
                sprReal(2).MaxRows = sprReal(2).MaxRows + 1
                sprReal(2).SetText 1, sprReal(2).MaxRows, Trim(rs!Descrip)
                sprReal(2).SetText 2, sprReal(2).MaxRows, Trim(rs!Cod_prod)
                
                sprReal(2).SetText 3, sprReal(2).MaxRows, str(rs!est_u)
                sprReal(2).SetText 4, sprReal(2).MaxRows, str(rs!vta_u)
                sprReal(2).SetText 6, sprReal(2).MaxRows, Trim(rs!est_p)
                sprReal(2).SetText 7, sprReal(2).MaxRows, str(rs!vta_p)
            
                sprReal(2).SetText 9, sprReal(2).MaxRows, str(rs!stk_u)
            Case "EZE"
                sprReal(1).MaxRows = sprReal(1).MaxRows + 1
                sprReal(1).SetText 1, sprReal(1).MaxRows, Trim(rs!Descrip)
                sprReal(1).SetText 2, sprReal(1).MaxRows, Trim(rs!Cod_prod)
                
                sprReal(1).SetText 3, sprReal(1).MaxRows, str(rs!est_u)
                sprReal(1).SetText 4, sprReal(1).MaxRows, str(rs!vta_u)
                sprReal(1).SetText 6, sprReal(1).MaxRows, Trim(rs!est_p)
                sprReal(1).SetText 7, sprReal(1).MaxRows, str(rs!vta_p)
                sprReal(1).SetText 9, sprReal(1).MaxRows, str(rs!stk_u)
            Case "CIA"
                sprReal(0).MaxRows = sprReal(0).MaxRows + 1
                sprReal(0).SetText 1, sprReal(0).MaxRows, Trim(rs!Descrip)
                sprReal(0).SetText 2, sprReal(0).MaxRows, Trim(rs!Cod_prod)
                
                sprReal(0).SetText 3, sprReal(0).MaxRows, str(rs!est_u)
                sprReal(0).SetText 4, sprReal(0).MaxRows, str(rs!vta_u)
                sprReal(0).SetText 6, sprReal(0).MaxRows, Trim(rs!est_p)
                sprReal(0).SetText 7, sprReal(0).MaxRows, str(rs!vta_p)
                sprReal(0).SetText 9, sprReal(0).MaxRows, str(rs!stk_u)
            Case "INT"
                sprReal(3).MaxRows = sprReal(3).MaxRows + 1
                sprReal(3).SetText 1, sprReal(3).MaxRows, Trim(rs!Descrip)
                sprReal(3).SetText 2, sprReal(3).MaxRows, Trim(rs!Cod_prod)
                
                sprReal(3).SetText 3, sprReal(3).MaxRows, str(rs!est_u)
                sprReal(3).SetText 4, sprReal(3).MaxRows, str(rs!vta_u)
                sprReal(3).SetText 6, sprReal(3).MaxRows, Trim(rs!est_p)
                sprReal(3).SetText 7, sprReal(3).MaxRows, str(rs!vta_p)
                sprReal(3).SetText 9, sprReal(3).MaxRows, str(rs!stk_u)
        End Select
        rs.MoveNext
        If rs.EOF Then
            Exit Do
        End If
    Loop

End Sub

Private Sub L_LlenarGrillasProy()

Do While Not rs.EOF
        Select Case rs!depn
            Case "AEP"
                sprProy(2).MaxRows = sprProy(2).MaxRows + 1
                sprProy(2).SetText 1, sprProy(2).MaxRows, Trim(rs!Descrip)
                sprProy(2).SetText 2, sprProy(2).MaxRows, Trim(rs!Cod_prod)
                
                sprProy(2).SetText 3, sprProy(2).MaxRows, str(rs!est_u)
                sprProy(2).SetText 4, sprProy(2).MaxRows, str(rs!vta_u)
                sprProy(2).SetText 6, sprProy(2).MaxRows, Trim(rs!est_p)
                sprProy(2).SetText 7, sprProy(2).MaxRows, str(rs!vta_p)
            
                sprProy(2).SetText 9, sprProy(2).MaxRows, str(rs!stk_u)
            Case "EZE"
                sprProy(1).MaxRows = sprProy(1).MaxRows + 1
                sprProy(1).SetText 1, sprProy(1).MaxRows, Trim(rs!Descrip)
                sprProy(1).SetText 2, sprProy(1).MaxRows, Trim(rs!Cod_prod)
                
                sprProy(1).SetText 3, sprProy(1).MaxRows, str(rs!est_u)
                sprProy(1).SetText 4, sprProy(1).MaxRows, str(rs!vta_u)
                sprProy(1).SetText 6, sprProy(1).MaxRows, Trim(rs!est_p)
                sprProy(1).SetText 7, sprProy(1).MaxRows, str(rs!vta_p)
            
                sprProy(1).SetText 9, sprProy(1).MaxRows, str(rs!stk_u)
            
            Case "CIA"
                sprProy(0).MaxRows = sprProy(0).MaxRows + 1
                sprProy(0).SetText 1, sprProy(0).MaxRows, Trim(rs!Descrip)
                sprProy(0).SetText 2, sprProy(0).MaxRows, Trim(rs!Cod_prod)
                
                sprProy(0).SetText 3, sprProy(0).MaxRows, str(rs!est_u)
                sprProy(0).SetText 4, sprProy(0).MaxRows, str(rs!vta_u)
                sprProy(0).SetText 6, sprProy(0).MaxRows, Trim(rs!est_p)
                sprProy(0).SetText 7, sprProy(0).MaxRows, str(rs!vta_p)
            
                sprProy(0).SetText 9, sprProy(0).MaxRows, str(rs!stk_u)
                
            Case "INT"
                sprProy(3).MaxRows = sprProy(3).MaxRows + 1
                sprProy(3).SetText 1, sprProy(3).MaxRows, Trim(rs!Descrip)
                sprProy(3).SetText 2, sprProy(3).MaxRows, Trim(rs!Cod_prod)
                
                sprProy(3).SetText 3, sprProy(3).MaxRows, str(rs!est_u)
                sprProy(3).SetText 4, sprProy(3).MaxRows, str(rs!vta_u)
                sprProy(3).SetText 6, sprProy(3).MaxRows, Trim(rs!est_p)
                sprProy(3).SetText 7, sprProy(3).MaxRows, str(rs!vta_p)
            
                sprProy(3).SetText 9, sprProy(3).MaxRows, str(rs!stk_u)
                
        End Select
        rs.MoveNext
        If rs.EOF Then
            Exit Do
        End If
    Loop

End Sub


Private Sub botExcel_Click()


    L_PasarGrillaExel sprReal(tabEst.Tab)
    L_TratarExcelTot tabEst.Tab


End Sub



Private Sub L_TratarExcelTot(ind As Integer)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim i As Integer
Dim tit As Variant
Dim ESPIGON As String

Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmMonitoreoProducto.caption = Aplicacion.SeteoProceso(frmMonitoreoProducto.caption)

If NOMBRE <> "" Then '

    Set AppExcel = CreateObject("excel.sheet")

    'AppExcel.Application.Visible = True

    ReDim titCol(sprEXEL.MaxCols)
    Col = 1
    fila = 3

    Select Case ind
       Case 0
            ESPIGON = "Total Compañía"
       Case 1
            ESPIGON = "Ezeiza"
       Case 2
            ESPIGON = "Aeroparque"
       Case 3
            ESPIGON = "Interior"
    End Select
    For i = 1 To sprEXEL.MaxCols
        sprEXEL.GetText i, 0, tit
        titCol(i) = tit
    Next

    Exl_PonerValor AppExcel, 1, 1, "INFORME PRODUCTOS MONITOREADOS CON MAS DE " & mskCota.Text & "% DE DESVIOS"
    Exl_PonerValor AppExcel, 2, 1, ESPIGON
    
    rango = Exl_rangos(1, 2, 1, sprEXEL.MaxCols)
    
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    'AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Blanco
    

    If cboRubro.Text <> "" Then
        Exl_PonerValor AppExcel, fila, Col, "RUBRO : " & cboRubro.Text
    Else
        Exl_PonerValor AppExcel, fila, Col, "RUBRO : Todos "
    End If
    fila = fila + 1
    If cboSubRubro.Text <> "" Then
        Exl_PonerValor AppExcel, fila, Col, "SUB RUBRO : " & LstEspigon.List(cboSubRubro.ListIndex)
    Else
        Exl_PonerValor AppExcel, fila, Col, "SUB RUBRO : Todos "
    End If
    fila = fila + 1
    If cboSubRubro.Text <> "" Then
        Exl_PonerValor AppExcel, fila, Col, "FAMILIA : " & lstFamilia.List(cboFamilia.ListIndex)
    Else
        Exl_PonerValor AppExcel, fila, Col, "FAMILIA : Todas "
    End If

    rango = Exl_rangos(fila - 2, fila, 1, sprEXEL.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False

    'AppExcel.Application.Range(rango).Merge
    'Exl_Lineas AppExcel, rango, Exl_Linsimple
    '---------------------------------------
'For i = 0 To 4
'
'    If sprTotal(i).MaxRows > 0 Then
'
'    fila = fila + 2
'
'    Exl_PonerValor AppExcel, fila, col, L_NombreDato(i + 1)
'    rango = Exl_rangos(fila, fila, 1, sprTotal(i).MaxCols)
'    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
'    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
'    Exl_ColorInt AppExcel, rango, Exl_Blanco
'    AppExcel.Application.Range(rango).Merge
'
    fila = fila + 1

    Exl_BajarGrillaExel sprEXEL, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEXEL.MaxRows, Col + 1, sprEXEL.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEXEL.MaxRows
    rango = Exl_rangos(fila, fila, Col, sprEXEL.MaxCols)
    'Exl_ColorInt AppExcel, rango, Exl_Gris
    'fila = fila + 1
''    rango = Exl_rangos(fila, fila, col, sprTotal(i).MaxCols)
''    Exl_ColorInt AppExcel, rango, Exl_
'
'    rango = Exl_rangos(fila - sprTotal(i).MaxRows, fila - 1, 6, sprTotal(i).MaxCols)
'    Exl_ColorInt AppExcel, rango, Exl_Gris
'
''    AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
'    End If
'
'Next
'    Exl.Exl_AnchoCol AppExcel, sprTotal(1).MaxCols, sprTotal(1).MaxCols, 1
    Exl.Exl_AnchoCol AppExcel, 2, 2, 25
    Exl.Exl_AnchoCol AppExcel, 1, 1, 19
    Exl.Exl_AnchoCol AppExcel, 5, 5, 8
    Exl.Exl_AnchoCol AppExcel, 8, 8, 8
'
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.Orientation = 2
'    AppExcel.Application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
'    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
'
    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmMonitoreoProducto.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub




Private Sub cboFamilia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    cboFamilia.ListIndex = -1
End If


End Sub

Private Sub cboProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    cboProv.ListIndex = -1
End If
End Sub

Private Sub cboRubro_Click()
Dim sql As String

sql = " SELECT cod_srub,descr FROM baires.subrubro "
sql = sql & " WHERE cod_rubr = '" & cboRubro.Text & "'"
sql = sql & " ORDER BY descr"
 
FuncCbos_LlenarCboLst cboSubRubro, LstEspigon, sql
    
cboFamilia.Clear

sql = " select distinct pv.cod_prov , pv.descrip "
sql = sql & " from  proveedor pv, producto p "
sql = sql & " where cod_prod in (select cod_prod from monitoreo_productos ) "
sql = sql & " and p.cod_prov = pv.cod_prov "
sql = sql & " and cod_rubr = '" & cboRubro.Text & "' "

FuncCbos_LlenarCboLst cboProv, lstProv, sql

End Sub


Private Sub cboRubro_KeyPress(KeyAscii As Integer)
Dim sql As String

If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    cboRubro.ListIndex = -1
    cboFamilia.ListIndex = -1

    sql = " select distinct pv.cod_prov , pv.descrip "
    sql = sql & " from  proveedor pv, producto p "
    sql = sql & " where cod_prod in (select cod_prod from monitoreo_productos ) "
    sql = sql & " and p.cod_prov = pv.cod_prov "
    
    FuncCbos_LlenarCboLst cboProv, lstProv, sql

End If

End Sub


Private Sub CboSubRubro_Click()
Dim sql As String

sql = " SELECT cod_FAMILIA,descrip FROM baires.familia "
sql = sql & " WHERE cod_rubr = '" & cboRubro.Text & "'"
sql = sql & " And cod_srub = '" & LstEspigon.List(cboSubRubro.ListIndex) & "'"
sql = sql & " ORDER BY descrip "
 
FuncCbos_LlenarCboLst cboFamilia, lstFamilia, sql

End Sub

Private Sub cboSubRubro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    cboSubRubro.ListIndex = -1
    cboFamilia.ListIndex = -1
End If

End Sub


Private Sub Form_Load()
Dim sql As String

Top = 30
Left = 250
Height = 7500
Width = 11000

sql = " SELECT cod_rubr,descrip FROM baires.rubro "
sql = sql & " ORDER BY cod_rubr"

FuncCbos_LlenarCbo cboRubro, sql

sql = " select distinct pv.cod_prov , pv.descrip "
sql = sql & " from  proveedor pv, producto p "
sql = sql & " where cod_prod in (select cod_prod from monitoreo_productos ) "
sql = sql & " and p.cod_prov = pv.cod_prov "

FuncCbos_LlenarCboLst cboProv, lstProv, sql

mskAnio.Text = Year(Date)
mskMes.Text = Month(Date)


L_LimpiarGrillas
'frmPrincipal.lstForms.AddItem "frmVsDia"

mskCota.Text = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)
    FuncLocal_SacarForm "frmVsDia"
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


Private Sub mskMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub mskMes_LostFocus()
If Val(mskMes.Text) < 1 Or Val(mskMes.Text) > 12 Then
    mskMes.Text = Month(Date)
End If

End Sub


Private Sub sprProy_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim prod As Variant
Dim desc As String
Dim sql As String

sprProy(Index).GetText 2, NewRow, prod
If prod <> "" Then
sql = "Select descrip From baires.producto where cod_prod = " & prod

Func.Func_ObtenerDesc sql, desc

lblProdPy(Index).caption = desc
End If

End Sub


Private Sub sprReal_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim prod As Variant
Dim desc As String
Dim sql As String
Dim rs As Recordset

sprReal(Index).GetText 2, NewRow, prod
If prod <> "" Then
    sql = "Select descrip From baires.producto where cod_prod = " & prod

    Func.Func_ObtenerDesc sql, desc

    lblProd(Index).caption = desc

     sql = "Select baires.Stock_loc(" & prod & ",'BOD') stk From dual "
     
     If Aplicacion.ObtenerRsDAO(sql, rs) Then
        lblSTKBOD(Index).caption = rs!stk
     End If
End If
End Sub


Private Sub tabEst_Click(PreviousTab As Integer)
On Error GoTo ErrT:

Call sprReal_LeaveCell(tabEst.Tab, 1, 1, 1, sprReal(tabEst.Tab).ActiveRow, True)
Call sprProy_LeaveCell(tabEst.Tab, 1, 1, 1, sprProy(tabEst.Tab).ActiveRow, True)
ErrT:
    Exit Sub

End Sub

Private Sub tabGA_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    'sprGA(tabGA.Tab).SetFocus
    
ErrT:
    Exit Sub

End Sub

Private Sub tabTotal_Click(PreviousTab As Integer)

On Error GoTo ErrT:

    'sprTotal(tabTotal.Tab).SetFocus
    
ErrT:
    Exit Sub
End Sub

