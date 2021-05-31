VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPMP 
   Caption         =   "Ventas por Proveedor-Marca"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   9195
   Begin TabDlg.SSTab tabEspigon 
      Height          =   3810
      Left            =   75
      TabIndex        =   5
      Top             =   1560
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   6720
      _Version        =   327680
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   459
      Enabled         =   0   'False
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "RESUMEN"
      TabPicture(0)   =   "frmPMP.frx":0000
      Tab(0).ControlCount=   0
      Tab(0).ControlEnabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmPMP.frx":001C
      Tab(1).ControlCount=   7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ExcelIntA"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "botGrafEvoEZEA"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "botGrTortaEA"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "botPorc(0)"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "botPorc(1)"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "botPorc(2)"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "tabEzeA"
      Tab(1).Control(6).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmPMP.frx":0038
      Tab(2).ControlCount=   7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "botExelEzeB"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "botGrafEvoEZEB"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "botGrTortaEB"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "botPorcB(2)"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "botPorcB(1)"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "botPorcB(0)"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "tabEzeB"
      Tab(2).Control(6).Enabled=   0   'False
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmPMP.frx":0054
      Tab(3).ControlCount=   7
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "tabAep"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "botPorcaep(2)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "botPorcaep(1)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "botPorcaep(0)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "botGrTortaAEP"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "botGrafEvoAEP"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "botExcelAep"
      Tab(3).Control(6).Enabled=   0   'False
      Begin VB.CommandButton botExcelAep 
         Caption         =   "Excel"
         Height          =   510
         Left            =   7800
         Picture         =   "frmPMP.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3150
         Width           =   690
      End
      Begin VB.CommandButton botExelEzeB 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67200
         Picture         =   "frmPMP.frx":0602
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3195
         Width           =   885
      End
      Begin VB.CommandButton ExcelIntA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67200
         Picture         =   "frmPMP.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3150
         Width           =   885
      End
      Begin VB.CommandButton botGrafEvoAEP 
         Height          =   540
         Left            =   7800
         Picture         =   "frmPMP.frx":1126
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   2565
         Width           =   690
      End
      Begin VB.CommandButton botGrafEvoEZEB 
         Height          =   540
         Left            =   -67200
         Picture         =   "frmPMP.frx":1568
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2610
         Width           =   880
      End
      Begin VB.CommandButton botGrafEvoEZEA 
         Height          =   540
         Left            =   -67200
         Picture         =   "frmPMP.frx":19AA
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2565
         Width           =   880
      End
      Begin VB.CommandButton botGrTortaAEP 
         Height          =   495
         Left            =   7800
         Picture         =   "frmPMP.frx":1DEC
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2025
         Width           =   690
      End
      Begin VB.CommandButton botGrTortaEB 
         Height          =   495
         Left            =   -67200
         Picture         =   "frmPMP.frx":222E
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2070
         Width           =   880
      End
      Begin VB.CommandButton botGrTortaEA 
         Height          =   495
         Left            =   -67200
         Picture         =   "frmPMP.frx":2670
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2025
         Width           =   880
      End
      Begin VB.CommandButton botPorcaep 
         Caption         =   "  Estandar  "
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
         Height          =   495
         Index           =   0
         Left            =   7800
         TabIndex        =   41
         Top             =   435
         Width           =   675
      End
      Begin VB.CommandButton botPorcaep 
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
         Index           =   1
         Left            =   7800
         Picture         =   "frmPMP.frx":2AB2
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   960
         Width           =   690
      End
      Begin VB.CommandButton botPorcaep 
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
         Index           =   2
         Left            =   7800
         Picture         =   "frmPMP.frx":3020
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1500
         Width           =   690
      End
      Begin VB.CommandButton botPorcB 
         Caption         =   " % Por Columnas "
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
         Index           =   2
         Left            =   -67200
         TabIndex        =   35
         Top             =   1530
         Width           =   880
      End
      Begin VB.CommandButton botPorcB 
         Caption         =   "  % Por       Filas   "
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
         Index           =   1
         Left            =   -67200
         TabIndex        =   34
         Top             =   990
         Width           =   880
      End
      Begin VB.CommandButton botPorcB 
         Caption         =   "  Estandar  "
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
         Height          =   495
         Index           =   0
         Left            =   -67200
         TabIndex        =   33
         Top             =   450
         Width           =   880
      End
      Begin VB.CommandButton botPorc 
         Caption         =   "  Estandar  "
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
         Height          =   495
         Index           =   0
         Left            =   -67200
         TabIndex        =   11
         Top             =   405
         Width           =   880
      End
      Begin VB.CommandButton botPorc 
         Caption         =   "  % Por       Filas   "
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
         Index           =   1
         Left            =   -67200
         TabIndex        =   10
         Top             =   945
         Width           =   880
      End
      Begin VB.CommandButton botPorc 
         Caption         =   " % Por Columnas "
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
         Index           =   2
         Left            =   -67200
         TabIndex        =   9
         Top             =   1485
         Width           =   880
      End
      Begin TabDlg.SSTab tabEzeA 
         Height          =   3330
         Left            =   -74895
         TabIndex        =   6
         Top             =   345
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   5874
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   441
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmPMP.frx":35A2
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frEzeAPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frEzeA0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmPMP.frx":35BE
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeAPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frEzeA1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmPMP.frx":35DA
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frEzeA2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "frEzeAPorc2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick."
         TabPicture(3)   =   "frmPMP.frx":35F6
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprEzeA(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom. Imp. x Pax"
         TabPicture(4)   =   "frmPMP.frx":3612
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprEzeA(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2700
            Index           =   3
            Left            =   -74895
            OleObjectBlob   =   "frmPMP.frx":362E
            TabIndex        =   60
            Top             =   150
            Width           =   7140
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2700
            Index           =   4
            Left            =   -74835
            OleObjectBlob   =   "frmPMP.frx":3B31
            TabIndex        =   61
            Top             =   165
            Width           =   7140
         End
         Begin VB.Frame frEzeA0 
            Height          =   2985
            Left            =   75
            TabIndex        =   12
            Top             =   30
            Width           =   7320
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmPMP.frx":4034
               TabIndex        =   24
               Top             =   210
               Width           =   7140
            End
         End
         Begin VB.Frame frEzeAPorc0 
            Height          =   2985
            Left            =   75
            TabIndex        =   18
            Top             =   30
            Width           =   7335
            Begin FPSpread.vaSpread sprEzeAPorc 
               Height          =   2700
               Index           =   0
               Left            =   75
               OleObjectBlob   =   "frmPMP.frx":4563
               TabIndex        =   29
               Top             =   195
               Width           =   7200
            End
         End
         Begin VB.Frame frEzeA1 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   15
            Top             =   45
            Width           =   7350
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   1
               Left            =   75
               OleObjectBlob   =   "frmPMP.frx":49CE
               TabIndex        =   27
               Top             =   180
               Width           =   7200
            End
         End
         Begin VB.Frame frEzeAPorc1 
            Height          =   2985
            Left            =   -74925
            TabIndex        =   19
            Top             =   45
            Width           =   7305
            Begin FPSpread.vaSpread sprEzeAPorc 
               Height          =   2700
               Index           =   1
               Left            =   60
               OleObjectBlob   =   "frmPMP.frx":4E8D
               TabIndex        =   30
               Top             =   210
               Width           =   7155
            End
         End
         Begin VB.Frame frEzeA2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   43
            Top             =   45
            Width           =   7335
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   2
               Left            =   60
               OleObjectBlob   =   "frmPMP.frx":52F8
               TabIndex        =   46
               Top             =   180
               Width           =   7200
            End
         End
         Begin VB.Frame frEzeAPorc2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   44
            Top             =   45
            Width           =   7335
            Begin FPSpread.vaSpread sprEzeAPorc 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmPMP.frx":57B7
               TabIndex        =   45
               Top             =   200
               Width           =   7155
            End
         End
      End
      Begin TabDlg.SSTab tabEzeB 
         Height          =   3330
         Left            =   -74925
         TabIndex        =   7
         Top             =   345
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5874
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   441
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmPMP.frx":5C22
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "frEzeBPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frEzeB0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmPMP.frx":5C3E
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeB1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frEzeBPorc1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmPMP.frx":5C5A
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frEzeBPorc2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "frEzeB2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom. Imp x Tick"
         TabPicture(3)   =   "frmPMP.frx":5C76
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprEzeB(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom. Imp x Pax"
         TabPicture(4)   =   "frmPMP.frx":5C92
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "sprEzeB(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2700
            Index           =   3
            Left            =   -74820
            OleObjectBlob   =   "frmPMP.frx":5CAE
            TabIndex        =   62
            Top             =   240
            Width           =   6600
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2700
            Index           =   4
            Left            =   165
            OleObjectBlob   =   "frmPMP.frx":60F2
            TabIndex        =   63
            Top             =   255
            Width           =   6600
         End
         Begin VB.Frame frEzeB2 
            Height          =   3000
            Left            =   -74910
            TabIndex        =   47
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmPMP.frx":6536
               TabIndex        =   50
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeB1 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   16
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmPMP.frx":6960
               TabIndex        =   28
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeBPorc1 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   21
            Top             =   60
            Width           =   6810
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   1
               Left            =   60
               OleObjectBlob   =   "frmPMP.frx":6D8A
               TabIndex        =   32
               Top             =   165
               Width           =   6645
            End
         End
         Begin VB.Frame frEzeB0 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   13
            Top             =   60
            Width           =   6795
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   0
               Left            =   75
               OleObjectBlob   =   "frmPMP.frx":716F
               TabIndex        =   25
               Top             =   165
               Width           =   6645
            End
         End
         Begin VB.Frame frEzeBPorc0 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   20
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmPMP.frx":75DF
               TabIndex        =   31
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeBPorc2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   48
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmPMP.frx":79C4
               TabIndex        =   49
               Top             =   200
               Width           =   6645
            End
         End
      End
      Begin TabDlg.SSTab tabAep 
         Height          =   3330
         Left            =   60
         TabIndex        =   8
         Top             =   315
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   5874
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   441
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmPMP.frx":7DA9
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "frAepPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frAep0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmPMP.frx":7DC5
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frAepPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frAep1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmPMP.frx":7DE1
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frAepPorc2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "frAep2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick"
         TabPicture(3)   =   "frmPMP.frx":7DFD
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprAep(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom Imp x Pax"
         TabPicture(4)   =   "frmPMP.frx":7E19
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "sprAep(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprAep 
            Height          =   2700
            Index           =   4
            Left            =   165
            OleObjectBlob   =   "frmPMP.frx":7E35
            TabIndex        =   65
            Top             =   225
            Width           =   6600
         End
         Begin FPSpread.vaSpread sprAep 
            Height          =   2700
            Index           =   3
            Left            =   -74820
            OleObjectBlob   =   "frmPMP.frx":81D8
            TabIndex        =   64
            Top             =   225
            Width           =   6600
         End
         Begin VB.Frame frAep2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   51
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmPMP.frx":857B
               TabIndex        =   54
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frAep1 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   17
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmPMP.frx":894A
               TabIndex        =   36
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frAepPorc1 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   23
            Top             =   30
            Width           =   6405
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   1
               Left            =   105
               OleObjectBlob   =   "frmPMP.frx":8D19
               TabIndex        =   38
               Top             =   210
               Width           =   6180
            End
         End
         Begin VB.Frame frAep0 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   14
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmPMP.frx":9083
               TabIndex        =   26
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frAepPorc0 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   22
            Top             =   45
            Width           =   6345
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   0
               Left            =   100
               OleObjectBlob   =   "frmPMP.frx":9452
               TabIndex        =   37
               Top             =   200
               Width           =   6180
            End
         End
         Begin VB.Frame frAepPorc2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   52
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmPMP.frx":97BC
               TabIndex        =   53
               Top             =   200
               Width           =   6600
            End
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1500
      Left            =   8100
      TabIndex        =   1
      Top             =   -45
      Width           =   810
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
         Height          =   405
         Index           =   0
         Left            =   120
         Picture         =   "frmPMP.frx":9B26
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
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
         Height          =   405
         Index           =   1
         Left            =   120
         Picture         =   "frmPMP.frx":9C28
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   570
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
         Height          =   450
         Index           =   2
         Left            =   120
         Picture         =   "frmPMP.frx":9D2A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   990
         Width           =   570
      End
   End
   Begin VB.Frame frdatos 
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   -75
      Width           =   8025
      Begin ComctlLib.Toolbar Toolbar3 
         Height          =   390
         Left            =   1095
         TabIndex        =   80
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         ImageList       =   "ImageList2"
         _Version        =   327680
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   3
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "a"
               Object.ToolTipText     =   "Ayuda Provvedores"
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "B"
               Object.ToolTipText     =   "Limpiar Prov."
               Object.Tag             =   ""
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   1080
         TabIndex        =   75
         Top             =   630
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         ImageList       =   "ImageList2"
         _Version        =   327680
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   3
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "a"
               Object.ToolTipText     =   "Ayuda Provvedores"
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "B"
               Object.ToolTipText     =   "Limpiar Prov."
               Object.Tag             =   ""
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   5895
         TabIndex        =   84
         Top             =   645
         Width           =   1890
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   6420
         TabIndex        =   83
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox txtDescripMr 
         Height          =   315
         Left            =   1995
         TabIndex        =   82
         Top             =   1110
         Width           =   3060
      End
      Begin VB.TextBox txtDescrip 
         Height          =   330
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   660
         Width           =   3060
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2430
         Picture         =   "frmPMP.frx":A54C
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   5265
         Picture         =   "frmPMP.frx":A6BE
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   210
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1275
         TabIndex        =   71
         Top             =   225
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFHasta 
         Height          =   285
         Left            =   4125
         TabIndex        =   72
         Top             =   225
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtCodPv 
         BackColor       =   &H80000018&
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
         Left            =   1140
         TabIndex        =   77
         Top             =   705
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtCodMr 
         Height          =   300
         Left            =   1185
         TabIndex        =   81
         Top             =   1110
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Marca"
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
         Left            =   120
         TabIndex        =   79
         Top             =   1110
         Width           =   945
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
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
         Index           =   3
         Left            =   105
         TabIndex        =   78
         Top             =   645
         Width           =   960
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
         Left            =   90
         TabIndex        =   74
         Top             =   225
         Width           =   1155
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
         Left            =   2970
         TabIndex        =   73
         Top             =   225
         Width           =   1125
      End
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   6465
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPMP.frx":A830
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPMP.frx":A942
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPMP.frx":AA54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPMP.frx":B286
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPMP.frx":B804
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPMP.frx":BD96
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPMP.frx":C0B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPMP.frx":C3CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPMP.frx":C6E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPMP.frx":C9FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPMP.frx":D230
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset
Dim RsDataPax  As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas
                       'con porcentajes
Dim UsoPorIntB As Boolean
Dim UsoPorAero As Boolean

Dim HELPProv As Boolean
Private Function L_Armarcondicion() As String
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant

'func_Dia1SegunMes_Anio(mskMes.Text, mskAnio.Text)
'func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text)

fechaDesde = mskFDesde.FormattedText
fechaHasta = mskFHasta.FormattedText

Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

L_Armarcondicion = Cond

End Function

Private Function L_ArmarcondicionPax() As String
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant

fechaDesde = mskFDesde.FormattedText
fechaHasta = mskFHasta.FormattedText

Cond = " WHERE fch_vta between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

L_ArmarcondicionPax = Cond

End Function


Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim Fila As Integer

Fila = 0
Do While Not RsData.EOF
    fecha = Format$(RsData!fch_ticket, FTOFECHA)
    Fila = Fila + 1
    Do While fecha = Format$(RsData!fch_ticket, FTOFECHA)

        Select Case RsData!cod_depn
            Case DSLoc(1).Dep
                L_LlenarGrillaAEP fecha, Fila
            Case DSLoc(2).Dep
                Select Case RsData!cod_sdep
                    Case DSLoc(2).SDep
                        L_LlenarGrillaINTA fecha, Fila
                    Case DSLoc(3).SDep
                        L_LlenarGrillaINTB fecha, Fila
                End Select
        End Select
        If RsData.EOF Then
            Exit Do
        End If
    Loop
Loop
For i = 0 To 1
    Spread_TotalesGrillas sprEzeA(i), sprEzeA(i).MaxCols - 1
    Spread_TotalesGrillas sprEzeB(i), sprEzeB(i).MaxCols - 1
    Spread_TotalesGrillas sprAep(i), sprAep(i).MaxCols - 1
    
Next
FuncLocal_PromediosFilaSPR sprEzeA(0), sprEzeA(1), sprEzeA(3)
FuncLocal_PromediosFilaSPR sprEzeA(0), sprEzeA(2), sprEzeA(4)
FuncLocal_PromediosFilaSPR sprAep(0), sprAep(1), sprAep(3)
FuncLocal_PromediosFilaSPR sprAep(0), sprAep(2), sprAep(4)
FuncLocal_PromediosFilaSPR sprEzeB(0), sprEzeB(1), sprEzeB(3)
FuncLocal_PromediosFilaSPR sprEzeB(0), sprEzeB(2), sprEzeB(4)

L_RefrescarFrames

End Sub
Private Sub L_DecoEspigonPax()
'Los pasajeros tienen un tratamiento especial
'Es otro recordset por eso se hace un seguimiento aparte
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim Fila As Integer

Fila = 0
Do While Not RsDataPax.EOF
    fecha = Format$(RsDataPax!fch_vta, FTOFECHA)
    Fila = Fila + 1
    Do While fecha = Format$(RsDataPax!fch_vta, FTOFECHA)

        Select Case RsDataPax!cod_depn
            Case DSLoc(1).Dep
                L_LlenarGrillaAEPPx fecha, Fila
            Case DSLoc(2).Dep
                Select Case RsDataPax!cod_sdep
                    Case DSLoc(2).SDep
                        L_LlenarGrillaINTAPax fecha, Fila
                    Case DSLoc(3).SDep
                        L_LlenarGrillaINTBPax fecha, Fila
                End Select
        End Select
        If RsDataPax.EOF Then
            Exit Do
        End If

    Loop
Loop

    Spread_TotalesGrillas sprEzeA(2), sprEzeA(2).MaxCols - 1
    Spread_TotalesGrillas sprEzeB(2), sprEzeB(2).MaxCols - 1
    Spread_TotalesGrillas sprAep(2), sprAep(2).MaxCols - 1


L_RefrescarFrames

End Sub

Private Sub L_LimpiarGrillas()
Dim i
    
For i = 0 To 4
    sprEzeA(i).MaxRows = 0
    sprEzeB(i).MaxRows = 0
    sprAep(i).MaxRows = 0

Next
    
End Sub

Private Sub L_PintarGrillas()
Dim i

For i = 0 To 4
    FuncLocal_PintarfinSemana sprEzeA(i)
Next

End Sub

Private Sub L_Porcentages(ByRef sprPorc As Control, sprDato As Control, Orientacion As String)
Dim i As Integer, j As Integer
Dim valor As Variant
Dim Tot As Variant, fecha As Variant
Dim Result As Single

On Error GoTo ErrPorc:

sprPorc.MaxRows = 0

Select Case Orientacion
    Case "FILAS"
        For i = 1 To sprDato.MaxRows
            sprDato.GetText sprDato.MaxCols, i, Tot
            sprDato.GetText 1, i, fecha
            Result = 0
            sprPorc.MaxRows = i
            
            sprPorc.SetText 1, i, Format$(fecha, "dd-mm-yy")
            sprPorc.SetText TotalINTA, i, "100"
            
            For j = 2 To sprDato.MaxCols - 1
                sprDato.GetText j, i, valor
                If Tot <> "" Then
                    Result = valor * 100 / Tot
                End If
                sprPorc.SetText j, i, Format$(Result, "% 0.00")
            Next
            
        Next
    Case "COL"
        For i = 1 To sprDato.MaxRows
         If i = sprDato.MaxRows Then
            sprPorc.MaxRows = i
            For j = 2 To sprDato.MaxCols
                sprPorc.SetText j, i, "100"
            Next
         Else
            sprDato.GetText 1, i, fecha
            Result = 0
            sprPorc.MaxRows = i
        
            sprPorc.SetText 1, i, Format$(fecha, "dd-mm-yy")
            
            For j = 2 To sprDato.MaxCols
                sprDato.GetText j, sprDato.MaxRows, Tot
                sprDato.GetText j, i, valor
                If Tot <> 0 Then
                    Result = valor * 100 / Tot
                End If
                sprPorc.SetText j, i, Format$(Result, "% 0.00")
            Next
         End If
        Next

End Select
        Spread_TotalesLinea sprPorc
ErrPorc:
    Exit Sub
End Sub

Private Sub L_RefrescarFrames()
    frEzeAPorc0.Refresh
    frEzeAPorc1.Refresh
    frEzeAPorc2.Refresh
    frEzeBPorc0.Refresh
    frEzeBPorc1.Refresh
    frEzeBPorc2.Refresh
    frAepPorc0.Refresh
    frAepPorc1.Refresh
    frAepPorc2.Refresh
End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim RS As Recordset

On Error GoTo ErrLGC:

sql = " SELECT "
sql = sql & " fch_ticket, "
sql = sql & " cod_depn, "
sql = sql & " cod_sdep, "
sql = sql & " cod_local, "
sql = sql & " cod_sloc, "
sql = sql & " sum(cantidad) cant, "
sql = sql & " sum(importe) imp "
sql = sql & "FROM " & funcLocal_Vista("venta_lgi", Year(CDate(mskFDesde.FormattedText)))
sql = sql & L_Armarcondicion
sql = sql & "group by fch_ticket,cod_depn,cod_sdep,cod_local,cod_sloc "
sql = sql & " order by fch_ticket,cod_depn,cod_sdep,cod_local,cod_sloc"

If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        L_DecoEspigon
    End If
    L_PintarGrillas
    Aplicacion.CerrarDAO RsData
    
End If

ErrLGC:
    Exit Sub
End Sub

Private Sub L_SeteoEjecutar(valor As Boolean, dato As String)

If valor Then
    Select Case dato
        Case EZEA
            UsoPorIntA = True
        Case EZEB
            UsoPorIntB = True
        Case AERO
            UsoPorAero = True
    End Select
    botEjecutar(1).Enabled = False
Else
    Select Case dato
        Case EZEA
            UsoPorIntA = False
        Case EZEB
            UsoPorIntB = False
        Case AERO
            UsoPorAero = False
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorAero) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub


Private Sub L_SeteoPorc(Index As Integer)
Dim i

Select Case Index
    Case 0
        frEzeA0.Visible = True
        frEzeA1.Visible = True
        frEzeA2.Visible = True
        botPorc(0).Enabled = False
        L_SeteoEjecutar False, EZEA
        frEzeA0.Refresh
        frEzeA1.Refresh
        frEzeA2.Refresh
    Case 1
        frEzeA0.Visible = False
        frEzeA1.Visible = False
        frEzeA2.Visible = False
        L_Porcentages sprEzeAPorc(0), sprEzeA(0), "FILAS"
        L_Porcentages sprEzeAPorc(1), sprEzeA(1), "FILAS"
        L_Porcentages sprEzeAPorc(2), sprEzeA(2), "FILAS"
        botPorc(0).Enabled = True
        frEzeAPorc0.Refresh
        frEzeAPorc1.Refresh
        frEzeAPorc2.Refresh
        L_SeteoEjecutar True, EZEA
    Case 2
        frEzeA0.Visible = False
        frEzeA1.Visible = False
        frEzeA2.Visible = False
        L_Porcentages sprEzeAPorc(0), sprEzeA(0), "COL"
        L_Porcentages sprEzeAPorc(1), sprEzeA(1), "COL"
        L_Porcentages sprEzeAPorc(2), sprEzeA(2), "COL"
        
        botPorc(0).Enabled = True
        frEzeAPorc0.Refresh
        frEzeAPorc1.Refresh
        frEzeAPorc2.Refresh
        L_SeteoEjecutar True, EZEA
        
End Select


End Sub

Private Sub L_SeteoTabPax(valor As Boolean)
    tabEzeA.TabEnabled(2) = valor
    tabEzeA.TabEnabled(4) = valor
    tabEzeB.TabEnabled(2) = valor
    tabEzeB.TabEnabled(4) = valor
    tabAep.TabEnabled(2) = valor
    tabAep.TabEnabled(4) = valor
End Sub

Private Sub L_TratarExcel(spr As Control, Titulo As String, subTit As String, info As String)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim Fila As Integer
Dim i
Dim tit As Variant
Dim nombre As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()

frmLGC.caption = Aplicacion.SeteoProceso(frmLGC.caption)

If nombre <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
'    AppExcel.application.Visible = True
    
    ReDim titCol(spr.MaxCols)
    col = 1
    Fila = 3
    
    For i = 1 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 1, 1, Titulo
    rango = Exl_rangos(1, 1, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro
    AppExcel.application.range(rango).merge
    Exl_Lineas AppExcel, rango
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, Fila, col, subTit
    rango = Exl_rangos(Fila, Fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq
    
    AppExcel.application.range(rango).merge
    Exl_Lineas AppExcel, rango
        
    Fila = Fila + 2
    rango = Exl_rangos(Fila, Fila + 2, 1, 1)
    
    Exl_PonerValor AppExcel, Fila, col, "Información sobre :" & info
    
    Fila = Fila + 2
    
    Exl_BajarGrillaExel spr, AppExcel, Fila, col, titCol
    Fila = Fila + spr.MaxRows
    rango = Exl_rangos(Fila, Fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    AppExcel.saveas nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmLGC.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub

Private Sub botEjecutar_Click(Index As Integer)

Select Case Index
    Case 0
        L_Refrescar
    Case 1
        Call botPorc_Click(0)
        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        tabEspigon.Enabled = False
        L_LimpiarGrillas
        L_RefrescarFrames
    Case 2
        Unload Me
End Select
End Sub

Private Sub botExcelAep_Click()
Select Case tabAep.Tab
    Case 0
        L_TratarExcel sprAep(0), "AEROPARQUE ", "Informe por Local", "Importes"
    Case 1
        L_TratarExcel sprAep(1), "AEROPARQUE", "Informe por Local", "Tickets"
    Case 2
        L_TratarExcel sprAep(2), "AEROPARQUE", "Informe por Local", "Pasajeros"
    Case 3
        L_TratarExcel sprAep(3), "AEROPARQUE", "Informe por Local", "Promedios por Tickets"
    Case 4
        L_TratarExcel sprAep(4), "AEROPARQUE", "Informe por Local", "Promedios por Pasajeros"
End Select
End Sub

Private Sub botExelEzeB_Click()
Select Case tabEzeB.Tab
    Case 0
        L_TratarExcel sprEzeB(0), "ESPIGON INTERNACIONAL B", "Informe por Local", "Importes"
    Case 1
        L_TratarExcel sprEzeB(1), "ESPIGON INTERNACIONAL B", "Informe por Local", "Tickets"
    Case 2
        L_TratarExcel sprEzeB(2), "ESPIGON INTERNACIONAL B", "Informe por Local", "Pasajeros"
    Case 3
        L_TratarExcel sprEzeB(3), "ESPIGON INTERNACIONAL B", "Informe por Local", "Promedios por Tickets"
    Case 4
        L_TratarExcel sprEzeB(4), "ESPIGON INTERNACIONAL B", "Informe por Local", "Promedios por Pasajeros"
End Select
End Sub

Private Sub botGrafEvoAEP_Click()
    Select Case tabAep.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico AERO, "Importes", sprAep(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico AERO, "Tickets", sprAep(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico AERO, "Pasajeros", sprAep(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select

End Sub

Private Sub botGrafEvoEZEA_Click()
    
    Select Case tabEzeA.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico EZEA, "Importes", sprEzeA(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico EZEA, "Tickets", sprEzeA(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico EZEA, "Pasajeros", sprEzeA(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select




End Sub

Private Sub botGrafEvoEZEB_Click()
    Select Case tabEzeB.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico EZEB, "Importes", sprEzeB(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico EZEB, "Tickets", sprEzeB(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico EZEB, "Pasajeros", sprEzeB(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select

End Sub


Private Sub botGrTortaAEP_Click()
Dim col_datos As Collection

    Set col_datos = New Collection
    
    Select Case tabAep.Tab
        Case 0
            L_LLenarColeccion col_datos, sprAep(0)
            FrmGraficos.CargarGrafico AERO, "Importes", col_datos
        Case 1
            L_LLenarColeccion col_datos, sprAep(1)
            FrmGraficos.CargarGrafico AERO, "Tickets", col_datos
        Case 2
            L_LLenarColeccion col_datos, sprAep(2)
            FrmGraficos.CargarGrafico AERO, "Pasajeros", col_datos
    End Select

End Sub

Private Sub botGrTortaEA_Click()
Dim col_datos As Collection

    Set col_datos = New Collection
    
    Select Case tabEzeA.Tab
        Case 0
            L_LLenarColeccion col_datos, sprEzeA(0)
            FrmGraficos.CargarGrafico EZEA, "Importes", col_datos
        Case 1
            L_LLenarColeccion col_datos, sprEzeA(1)
            FrmGraficos.CargarGrafico EZEA, "Tickets", col_datos
        Case 2
            L_LLenarColeccion col_datos, sprEzeA(2)
            FrmGraficos.CargarGrafico EZEA, "Pasajeros", col_datos

    End Select
End Sub


Private Sub botGrTortaEB_Click()
Dim col_datos As Collection

    Set col_datos = New Collection
    
    Select Case tabEzeB.Tab
        Case 0
            L_LLenarColeccion col_datos, sprEzeB(0)
            FrmGraficos.CargarGrafico EZEB, "Importes", col_datos
        Case 1
            L_LLenarColeccion col_datos, sprEzeB(1)
            FrmGraficos.CargarGrafico EZEB, "Tickets", col_datos
        Case 2
            L_LLenarColeccion col_datos, sprEzeB(2)
            FrmGraficos.CargarGrafico EZEB, "Pasajeros", col_datos

    End Select

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


Private Sub botPorc_Click(Index As Integer)
Dim i

Select Case Index
    Case 0
        frEzeA0.Visible = True
        frEzeA1.Visible = True
        frEzeA2.Visible = True
        botPorc(0).Enabled = False
        L_SeteoEjecutar False, EZEA
        frEzeA0.Refresh
        frEzeA1.Refresh
        frEzeA2.Refresh
    Case 1
        frEzeA0.Visible = False
        frEzeA1.Visible = False
        frEzeA2.Visible = False
        L_Porcentages sprEzeAPorc(0), sprEzeA(0), "FILAS"
        L_Porcentages sprEzeAPorc(1), sprEzeA(1), "FILAS"
        L_Porcentages sprEzeAPorc(2), sprEzeA(2), "FILAS"
        botPorc(0).Enabled = True
        frEzeAPorc0.Refresh
        frEzeAPorc1.Refresh
        frEzeAPorc2.Refresh
        L_SeteoEjecutar True, EZEA
    Case 2
        frEzeA0.Visible = False
        frEzeA1.Visible = False
        frEzeA2.Visible = False
        L_Porcentages sprEzeAPorc(0), sprEzeA(0), "COL"
        L_Porcentages sprEzeAPorc(1), sprEzeA(1), "COL"
        L_Porcentages sprEzeAPorc(2), sprEzeA(2), "COL"
        
        botPorc(0).Enabled = True
        frEzeAPorc0.Refresh
        frEzeAPorc1.Refresh
        frEzeAPorc2.Refresh
        L_SeteoEjecutar True, EZEA
        
End Select

End Sub

Private Sub botPorcaep_Click(Index As Integer)

Select Case Index
    Case 0
        frAep0.Visible = True
        frAep1.Visible = True
        frAep2.Visible = True
        botPorcaep(0).Enabled = False
        L_SeteoEjecutar False, AERO
        frAep0.Refresh
        frAep1.Refresh
        frAep2.Refresh
    Case 1
        frAep0.Visible = False
        frAep1.Visible = False
        frAep2.Visible = False
        L_Porcentages sprAepPorc(0), sprAep(0), "FILAS"
        L_Porcentages sprAepPorc(1), sprAep(1), "FILAS"
        L_Porcentages sprAepPorc(2), sprAep(2), "FILAS"
        botPorcaep(0).Enabled = True
        frAepPorc0.Refresh
        frAepPorc1.Refresh
        frAepPorc2.Refresh
        L_SeteoEjecutar True, AERO
    Case 2
        frAep0.Visible = False
        frAep1.Visible = False
        frAep2.Visible = False
        L_Porcentages sprAepPorc(0), sprAep(0), "COL"
        L_Porcentages sprAepPorc(1), sprAep(1), "COL"
        L_Porcentages sprAepPorc(2), sprAep(2), "COL"
        
        botPorcaep(0).Enabled = True
        frAepPorc0.Refresh
        frAepPorc1.Refresh
        frAepPorc2.Refresh
        L_SeteoEjecutar True, AERO
        
End Select


End Sub

Private Sub botPorcB_Click(Index As Integer)

Select Case Index
    Case 0
        frEzeB0.Visible = True
        frEzeB1.Visible = True
        frEzeB2.Visible = True
        botPorcB(0).Enabled = False
        L_SeteoEjecutar False, EZEB
        frEzeB0.Refresh
        frEzeB1.Refresh
        frEzeB2.Refresh
    Case 1
        frEzeB0.Visible = False
        frEzeB1.Visible = False
        frEzeB2.Visible = False
        L_Porcentages sprEzeBPorc(0), sprEzeB(0), "FILAS"
        L_Porcentages sprEzeBPorc(1), sprEzeB(1), "FILAS"
        L_Porcentages sprEzeBPorc(2), sprEzeB(2), "FILAS"
        botPorcB(0).Enabled = True
        frEzeBPorc0.Refresh
        frEzeBPorc1.Refresh
        frEzeBPorc2.Refresh
        L_SeteoEjecutar True, EZEB
    Case 2
        frEzeB0.Visible = False
        frEzeB1.Visible = False
        frEzeB2.Visible = False
        L_Porcentages sprEzeBPorc(0), sprEzeB(0), "COL"
        L_Porcentages sprEzeBPorc(1), sprEzeB(1), "COL"
        L_Porcentages sprEzeBPorc(2), sprEzeB(2), "COL"
        
        botPorcB(0).Enabled = True
        frEzeBPorc0.Refresh
        frEzeBPorc1.Refresh
        frEzeBPorc2.Refresh
        L_SeteoEjecutar True, EZEB
        
End Select

End Sub




Private Sub ExcelIntA_Click()

Select Case tabEzeA.Tab
    Case 0
        L_TratarExcel sprEzeA(0), "ESPIGON INTERNACIONAL A", "Informe por Local", "Importes"
    Case 1
        L_TratarExcel sprEzeA(1), "ESPIGON INTERNACIONAL A", "Informe por Local", "Tickets"
    Case 2
        L_TratarExcel sprEzeA(2), "ESPIGON INTERNACIONAL A", "Informe por Local", "Pasajeros"
    Case 3
        L_TratarExcel sprEzeA(3), "ESPIGON INTERNACIONAL A", "Informe por Local", "Promedios por Tickets"
    Case 4
        L_TratarExcel sprEzeA(4), "ESPIGON INTERNACIONAL A", "Informe por Local", "Promedios por Pasajeros"
        
End Select
    
End Sub


Private Sub Form_Activate()
FuncLocal_SeteoTABS tabEspigon
End Sub

Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6000
Me.Width = 9300

mskFDesde.Text = func_Dia1SegunMes_Anio(Month(Date), Year(Date))
mskFHasta.Text = Format$(Date - 1, FTOFECHA)

'FuncLocal_SeteoTABS tabEspigon

UsoPorIntA = False
UsoPorIntB = False
UsoPorAero = False

L_RefrescarFrames
End Sub




Private Sub mskFDesde_LostFocus()
    
    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Date
    End If
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
    
    If mskFHasta.Text <> "" Then
        If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Or Year(mskFHasta.FormattedText) <> Year(mskFDesde.FormattedText) Then
            mskFHasta.Text = mskFDesde.Text
        End If
    End If

End Sub


Private Sub mskFHasta_LostFocus()
If Not IsDate(mskFHasta.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
ElseIf CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Or Year(mskFHasta.FormattedText) <> Year(mskFDesde.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
End If

mskFHasta.Text = Format$(mskFHasta.FormattedText, FTOFECHA)
End Sub



Private Sub L_LlenarGrillaINTA(fecha As String, Fila As Integer)
Dim Depn As String, SDep As String
Dim i As Integer
Dim totCant As Long
Dim TOTImp As Long


Depn = RsData!cod_depn
SDep = RsData!cod_sdep

sprEzeA(0).MaxRows = sprEzeA(0).MaxRows + 1
sprEzeA(1).MaxRows = sprEzeA(1).MaxRows + 1
sprEzeA(3).MaxRows = sprEzeA(3).MaxRows + 1

sprEzeA(0).SetText 1, sprEzeA(0).MaxRows, Format$(fecha, "dd-mm-yy")
sprEzeA(1).SetText 1, sprEzeA(1).MaxRows, Format$(fecha, "dd-mm-yy")
sprEzeA(3).SetText 1, sprEzeA(3).MaxRows, Format$(fecha, "dd-mm-yy")


totCant = 0
TOTImp = 0

     Do While fecha = Format$(RsData!fch_ticket, FTOFECHA) And _
        Depn = RsData!cod_depn And SDep = RsData!cod_sdep
        
            For i = 1 To 8
                If DSLoc(2).Locales(i) = RsData!cod_local & Trim(str(RsData!cod_sloc)) Then
                    sprEzeA(0).SetText i + 1, sprEzeA(0).MaxRows, str(RsData!imp)
                    sprEzeA(1).SetText i + 1, sprEzeA(1).MaxRows, str(RsData!cant)
                    sprEzeA(3).SetText i + 1, sprEzeA(3).MaxRows, str(RsData!imp / RsData!cant)
                 
                    totCant = totCant + RsData!cant
                    TOTImp = TOTImp + RsData!imp
                    
                    RsData.MoveNext
                End If
                    If RsData.EOF Then
                        Exit For
                    End If
            
            Next
            sprEzeA(3).SetText sprEzeA(3).MaxCols, sprEzeA(3).MaxRows, Format$((TOTImp / totCant), "#.0")
            
            If RsData.EOF Then
                Exit Do
            End If
            
        Loop
End Sub

Private Sub L_LlenarGrillaINTAPax(fecha As String, Fila As Integer)
Dim Depn As String, SDep As String
Dim i As Integer
Dim TOTAL As Variant
Dim totCant As Long
Dim TOTImp As Long

Depn = RsDataPax!cod_depn
SDep = RsDataPax!cod_sdep

sprEzeA(2).MaxRows = sprEzeA(2).MaxRows + 1
sprEzeA(4).MaxRows = sprEzeA(4).MaxRows + 1

sprEzeA(2).SetText 1, sprEzeA(2).MaxRows, Format$(fecha, "dd-mm-yy")
sprEzeA(4).SetText 1, sprEzeA(4).MaxRows, Format$(fecha, "dd-mm-yy")

totCant = 0
TOTImp = 0
     Do While fecha = Format$(RsDataPax!fch_vta, FTOFECHA) And _
        Depn = RsDataPax!cod_depn And SDep = RsDataPax!cod_sdep
        
            For i = 1 To 8
                If DSLoc(2).Locales(i) = RsDataPax!cod_local & Trim(str(RsDataPax!cod_sloc)) Then
                    sprEzeA(2).SetText i + 1, sprEzeA(2).MaxRows, str(RsDataPax!cant)
                    sprEzeA(4).SetText i + 1, sprEzeA(4).MaxRows, str(RsDataPax!imp / RsDataPax!cant)
                    
                    totCant = totCant + RsDataPax!cant
                    TOTImp = TOTImp + RsDataPax!imp
                    
                    RsDataPax.MoveNext
                End If
                    If RsDataPax.EOF Then
                        Exit For
                    End If
            
            Next
            
            sprEzeA(4).SetText sprEzeA(4).MaxCols, sprEzeA(4).MaxRows, Format$((TOTImp / totCant), "#.0")
            If RsDataPax.EOF Then
                Exit Do
            End If
            
        Loop
        
End Sub


Private Sub L_LlenarGrillaINTB(fecha As String, Fila As Integer)
Dim Depn As String, SDep As String
Dim i As Integer
Dim totCant As Long
Dim TOTImp As Long


Depn = RsData!cod_depn
SDep = RsData!cod_sdep

sprEzeB(0).MaxRows = sprEzeB(0).MaxRows + 1
sprEzeB(1).MaxRows = sprEzeB(1).MaxRows + 1
sprEzeB(3).MaxRows = sprEzeB(3).MaxRows + 1

sprEzeB(0).SetText 1, sprEzeB(0).MaxRows, Format$(fecha, "dd-mm-yy")
sprEzeB(1).SetText 1, sprEzeB(1).MaxRows, Format$(fecha, "dd-mm-yy")
sprEzeB(3).SetText 1, sprEzeB(3).MaxRows, Format$(fecha, "dd-mm-yy")

totCant = 0
TOTImp = 0

     Do While fecha = Format$(RsData!fch_ticket, FTOFECHA) And _
        Depn = RsData!cod_depn And SDep = RsData!cod_sdep
        
            For i = 1 To 4
                If DSLoc(3).Locales(i) = RsData!cod_local & Trim(str(RsData!cod_sloc)) Then
                    sprEzeB(0).SetText i + 1, sprEzeB(0).MaxRows, str(RsData!imp)
                    sprEzeB(1).SetText i + 1, sprEzeB(1).MaxRows, str(RsData!cant)
                    sprEzeB(3).SetText i + 1, sprEzeB(3).MaxRows, str(RsData!imp / RsData!cant)
                    
                    totCant = totCant + RsData!cant
                    TOTImp = TOTImp + RsData!imp
                    
                    RsData.MoveNext
                End If
                    If RsData.EOF Then
                        Exit For
                    End If
            
            Next
            sprEzeB(3).SetText sprEzeB(3).MaxCols, sprEzeB(3).MaxRows, Format$((TOTImp / totCant), "#.0")
            
            If RsData.EOF Then
                Exit Do
            End If
        
        Loop
End Sub
Private Sub L_LlenarGrillaINTBPax(fecha As String, Fila As Integer)
Dim Depn As String, SDep As String
Dim i As Integer
Dim TOTAL As Variant
Dim totCant As Long
Dim TOTImp As Long

Depn = RsDataPax!cod_depn
SDep = RsDataPax!cod_sdep

sprEzeB(2).MaxRows = sprEzeB(2).MaxRows + 1
sprEzeB(4).MaxRows = sprEzeB(4).MaxRows + 1

sprEzeB(2).SetText 1, sprEzeB(2).MaxRows, Format$(fecha, "dd-mm-yy")
sprEzeB(4).SetText 1, sprEzeB(4).MaxRows, Format$(fecha, "dd-mm-yy")

totCant = 0
TOTImp = 0
     Do While fecha = Format$(RsDataPax!fch_vta, FTOFECHA) And _
        Depn = RsDataPax!cod_depn And SDep = RsDataPax!cod_sdep
        
            For i = 1 To 4
                If DSLoc(3).Locales(i) = RsDataPax!cod_local & Trim(str(RsDataPax!cod_sloc)) Then
                    sprEzeB(2).SetText i + 1, sprEzeB(2).MaxRows, str(RsDataPax!cant)
                    sprEzeB(4).SetText i + 1, sprEzeB(4).MaxRows, str(RsDataPax!imp / RsDataPax!cant)
                    
                    totCant = totCant + RsDataPax!cant
                    TOTImp = TOTImp + RsDataPax!imp
                    
                    RsDataPax.MoveNext
                End If
                    If RsDataPax.EOF Then
                        Exit For
                    End If
            
            Next
            sprEzeB(4).SetText sprEzeB(4).MaxCols, sprEzeB(4).MaxRows, Format$((TOTImp / totCant), "#.0")
            
            If RsDataPax.EOF Then
                Exit Do
            End If
        
        Loop
        

End Sub


Private Sub L_LlenarGrillaAEP(fecha As String, Fila As Integer)
Dim Depn As String, SDep As String
Dim i As Integer
Dim TOTALImp As Variant
Dim TOTALCant As Variant

Depn = RsData!cod_depn
SDep = RsData!cod_sdep

sprAep(0).MaxRows = sprAep(0).MaxRows + 1
sprAep(1).MaxRows = sprAep(1).MaxRows + 1
sprAep(3).MaxRows = sprAep(3).MaxRows + 1

sprAep(0).SetText 1, sprAep(0).MaxRows, Format$(fecha, "dd-mm-yy")
sprAep(1).SetText 1, sprAep(1).MaxRows, Format$(fecha, "dd-mm-yy")
sprAep(3).SetText 1, sprAep(3).MaxRows, Format$(fecha, "dd-mm-yy")

    TOTALImp = 0
    TOTALCant = 0
     Do While fecha = Format$(RsData!fch_ticket, FTOFECHA) And _
        Depn = RsData!cod_depn And SDep = RsData!cod_sdep
        
            For i = 1 To 2
                If DSLoc(1).Locales(i) = RsData!cod_local & Trim(str(RsData!cod_sloc)) Then
                    sprAep(0).SetText i + 1, sprAep(0).MaxRows, str(RsData!imp)
                    sprAep(1).SetText i + 1, sprAep(1).MaxRows, str(RsData!cant)
                    sprAep(3).SetText i + 1, sprAep(3).MaxRows, str(RsData!imp / RsData!cant)
                    
                    TOTALImp = TOTALImp + RsData!imp
                    TOTALCant = TOTALCant + RsData!cant
                    RsData.MoveNext
                End If
                    If RsData.EOF Then
                        Exit For
                    End If
            
            Next
            sprAep(3).SetText sprAep(3).MaxCols, sprAep(3).MaxRows, Format$(TOTALImp / TOTALCant, "#.0")
            
            If RsData.EOF Then
                Exit Do
            End If
        
        Loop
End Sub


Private Sub L_LlenarGrillaAEPPx(fecha As String, Fila As Integer)
Dim Depn As String, SDep As String
Dim i As Integer
Dim TOTAL As Variant
Dim totCant As Long
Dim TOTImp As Long

Depn = RsDataPax!cod_depn
SDep = RsDataPax!cod_sdep

sprAep(2).MaxRows = sprAep(2).MaxRows + 1
sprAep(4).MaxRows = sprAep(4).MaxRows + 1

sprAep(2).SetText 1, sprAep(2).MaxRows, Format$(fecha, "dd-mm-yy")
sprAep(4).SetText 1, sprAep(4).MaxRows, Format$(fecha, "dd-mm-yy")

totCant = 0
TOTImp = 0
     Do While fecha = Format$(RsDataPax!fch_vta, FTOFECHA) And _
        Depn = RsDataPax!cod_depn And SDep = RsDataPax!cod_sdep
        
            For i = 1 To 2
                If DSLoc(1).Locales(i) = RsDataPax!cod_local & Trim(str(RsDataPax!cod_sloc)) Then
                    sprAep(2).SetText i + 1, sprAep(2).MaxRows, str(RsDataPax!cant)
                    sprAep(4).SetText i + 1, sprAep(4).MaxRows, str(RsDataPax!imp / RsDataPax!cant)
                    
                    totCant = totCant + RsDataPax!cant
                    TOTImp = TOTImp + RsDataPax!imp
                    
                    RsDataPax.MoveNext
                
                End If
                    If RsDataPax.EOF Then
                        Exit For
                    End If
            
            Next

            sprAep(4).SetText sprAep(4).MaxCols, sprAep(4).MaxRows, Format$((TOTImp / totCant), "#.0")
            
            If RsDataPax.EOF Then
                Exit Do
            End If
        
        Loop
        

End Sub



Private Sub tabAep_Click(PreviousTab As Integer)
frAep0.Refresh
frAep1.Refresh
frAep2.Refresh
Select Case tabAep.Tab
    Case 0, 1, 2
        botPorcaep(0).Visible = True
        botPorcaep(1).Visible = True
        botPorcaep(2).Visible = True
        botGrTortaAEP.Visible = True
        botGrafEvoAEP.Visible = True
    Case 3, 4
        botPorcaep(0).Visible = False
        botPorcaep(1).Visible = False
        botPorcaep(2).Visible = False
        botGrTortaAEP.Visible = False
        botGrafEvoAEP.Visible = False
End Select

End Sub

Private Sub tabEzeA_Click(PreviousTab As Integer)
frEzeA0.Refresh
frEzeA1.Refresh
frEzeA2.Refresh
Select Case tabEzeA.Tab
    Case 0, 1, 2
        botPorc(0).Visible = True
        botPorc(1).Visible = True
        botPorc(2).Visible = True
        botGrTortaEA.Visible = True
        botGrafEvoEZEA.Visible = True
    Case 3, 4
        botPorc(0).Visible = False
        botPorc(1).Visible = False
        botPorc(2).Visible = False
        botGrTortaEA.Visible = False
        botGrafEvoEZEA.Visible = False
End Select

End Sub

Private Sub tabEzeB_Click(PreviousTab As Integer)
frEzeB0.Refresh
frEzeB1.Refresh
frEzeB2.Refresh
Select Case tabEzeB.Tab
    Case 0, 1, 2
        botPorcB(0).Visible = True
        botPorcB(1).Visible = True
        botPorcB(2).Visible = True
        botGrTortaEB.Visible = True
        botGrafEvoEZEB.Visible = True
    Case 3, 4
        botPorcB(0).Visible = False
        botPorcB(1).Visible = False
        botPorcB(2).Visible = False
        botGrTortaEB.Visible = False
        botGrafEvoEZEB.Visible = False
End Select
End Sub

Private Sub L_LLenarColeccion(ByRef col As Collection, spr As Control)
Dim cl_dato As CLlgi
Dim valor As Variant
Dim fecha As Variant
Dim i

spr.Row = spr.ActiveRow

spr.GetText 1, spr.Row, fecha


For i = 2 To spr.MaxCols - 1
    Set cl_dato = New CLlgi
        
    cl_dato.fch = fecha
    
    spr.GetText i, 0, valor
    cl_dato.locale = valor
    
    spr.GetText i, spr.Row, valor
    cl_dato.DatoGral = valor
    
    col.Add cl_dato
    
Next

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim codigo$, desc$
Dim sqlProv$
    
Select Case Button.Key
    Case "a"
        If frmHelpPv.MuestraHlp(codigo$, desc$, "PROVEEDOR", "Select cod_prov, descrip from baires.proveedor ") = vbOK Then
            txtCodPv.Text = codigo$
            txtDescrip.Text = desc$
    
            HELPProv = True
            SendKeys "{TAB}"
    
        End If
'        txtCodPv.SetFocus
    Case "B"
            txtCodPv.Text = "-"
            txtDescrip.Text = ""
            txtCodMr.Text = ""
            txtDescripMr.Text = ""

End Select

End Sub


Private Sub Toolbar3_ButtonClick(ByVal Button As ComctlLib.Button)
Dim codigo$, desc$
Dim sqlProv$
    
Select Case Button.Key
    Case "a"
        If frmHelpPv.MuestraHlp(codigo$, desc$, "MARCA", "Select cod_marca, descrip from baires.marca WHERE cod_prov = '" & txtCodPv.Text & "' ") = vbOK Then
            txtCodMr.Text = codigo$
            txtDescripMr.Text = desc$
            SendKeys "{TAB}"
    
        End If
    Case "b"
            txtCodMr.Text = ""
            txtDescripMr.Text = ""
End Select


End Sub


Private Sub txtCodPv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtCodPv.Text = "" Then
        Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
    Else
        SendKeys "{TAB}"
    End If
End If
End Sub


Private Sub txtCodPv_LostFocus()
Dim sql$
Dim desc$

If txtCodPv.Text <> "" Then
      If Not HELPProv Then
'        sql$ = "SELECT  descrip FROM proveedor WHERE cod_prov = '" & UCase(txtCodPv.Text) & "'"
'        If FuncLocales_ObtenerDesc(sql$, desc$) Then
'            MeLlenarcboMarca
'        Else
        '    cboMarca.Clear
        '    txtDescMarca.Text = ""
'        End If
'        txtDescrip.Text = desc$
      Else
        'MeLlenarcboMarca
      End If
Else
    txtDescrip.Text = ""
    
End If
  
    txtCodPv.Text = UCase(txtCodPv.Text)

    HELPProv = False


End Sub


