VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmComp 
   Caption         =   "Comparativo de Ventas -Inf. por Local-"
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
      Left            =   60
      TabIndex        =   2
      Top             =   1560
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   6720
      _Version        =   327680
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   459
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
      TabPicture(0)   =   "frmComp.frx":0000
      Tab(0).ControlCount=   0
      Tab(0).ControlEnabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmComp.frx":001C
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "tabEzeA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ExcelIntA"
      Tab(1).Control(1).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmComp.frx":0038
      Tab(2).ControlCount=   2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "botExelEzeB"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "tabEzeB"
      Tab(2).Control(1).Enabled=   0   'False
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmComp.frx":0054
      Tab(3).ControlCount=   2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tabAep"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "botExcelAep"
      Tab(3).Control(1).Enabled=   -1  'True
      TabCaption(4)   =   "INTERIOR"
      TabPicture(4)   =   "frmComp.frx":0070
      Tab(4).ControlCount=   2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TabInt"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "botExcelInt"
      Tab(4).Control(1).Enabled=   -1  'True
      Begin TabDlg.SSTab TabInt 
         Height          =   3330
         Left            =   -74835
         TabIndex        =   52
         Top             =   300
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   5874
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmComp.frx":008C
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FrInt0"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmComp.frx":00A8
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FrInt1"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmComp.frx":00C4
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FrInt2"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Prom. Imp x Tick"
         TabPicture(3)   =   "frmComp.frx":00E0
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprInt(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom Imp x Pax"
         TabPicture(4)   =   "frmComp.frx":00FC
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprInt(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprInt 
            Height          =   2700
            Index           =   3
            Left            =   -74730
            OleObjectBlob   =   "frmComp.frx":0118
            TabIndex        =   59
            Top             =   165
            Width           =   7140
         End
         Begin FPSpread.vaSpread sprInt 
            Height          =   2700
            Index           =   4
            Left            =   -74730
            OleObjectBlob   =   "frmComp.frx":0577
            TabIndex        =   60
            Top             =   180
            Width           =   7140
         End
         Begin VB.Frame FrInt0 
            Height          =   2910
            Left            =   75
            TabIndex        =   53
            Top             =   45
            Width           =   7260
            Begin FPSpread.vaSpread sprInt 
               Height          =   2700
               Index           =   0
               Left            =   60
               OleObjectBlob   =   "frmComp.frx":09D6
               TabIndex        =   54
               Top             =   165
               Width           =   7140
            End
         End
         Begin VB.Frame FrInt2 
            Height          =   2895
            Left            =   -74850
            TabIndex        =   57
            Top             =   45
            Width           =   7275
            Begin FPSpread.vaSpread sprInt 
               Height          =   2655
               Index           =   2
               Left            =   60
               OleObjectBlob   =   "frmComp.frx":0E2F
               TabIndex        =   58
               Top             =   195
               Width           =   7140
            End
         End
         Begin VB.Frame FrInt1 
            Height          =   2970
            Left            =   -75000
            TabIndex        =   55
            Top             =   0
            Width           =   7260
            Begin FPSpread.vaSpread sprInt 
               Height          =   2700
               Index           =   1
               Left            =   60
               OleObjectBlob   =   "frmComp.frx":128E
               TabIndex        =   56
               Top             =   165
               Width           =   7140
            End
         End
      End
      Begin VB.CommandButton botExcelInt 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67035
         Picture         =   "frmComp.frx":16ED
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2940
         Width           =   765
      End
      Begin VB.CommandButton ExcelIntA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   7845
         Picture         =   "frmComp.frx":1C7F
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2910
         Width           =   780
      End
      Begin VB.CommandButton botExelEzeB 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67125
         Picture         =   "frmComp.frx":2211
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2910
         Width           =   765
      End
      Begin VB.CommandButton botExcelAep 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67185
         Picture         =   "frmComp.frx":27A3
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2925
         Width           =   765
      End
      Begin TabDlg.SSTab tabEzeA 
         Height          =   3330
         Left            =   165
         TabIndex        =   4
         Top             =   360
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
         TabPicture(0)   =   "frmComp.frx":2D35
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frEzeA0"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmComp.frx":2D51
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeA1"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmComp.frx":2D6D
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frEzeA2"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick."
         TabPicture(3)   =   "frmComp.frx":2D89
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprEzeA(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom. Imp. x Pax"
         TabPicture(4)   =   "frmComp.frx":2DA5
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprEzeA(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2700
            Index           =   4
            Left            =   -74895
            OleObjectBlob   =   "frmComp.frx":2DC1
            TabIndex        =   36
            Top             =   270
            Width           =   7140
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2700
            Index           =   3
            Left            =   -74850
            OleObjectBlob   =   "frmComp.frx":3220
            TabIndex        =   35
            Top             =   255
            Width           =   7140
         End
         Begin VB.Frame frEzeA0 
            Height          =   2985
            Left            =   90
            TabIndex        =   5
            Top             =   30
            Width           =   7305
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   0
               Left            =   105
               OleObjectBlob   =   "frmComp.frx":367F
               TabIndex        =   8
               Top             =   195
               Width           =   7140
            End
         End
         Begin VB.Frame frEzeA1 
            Height          =   2985
            Left            =   -74880
            TabIndex        =   6
            Top             =   30
            Width           =   7305
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmComp.frx":3AD8
               TabIndex        =   33
               Top             =   210
               Width           =   7140
            End
         End
         Begin VB.Frame frEzeA2 
            Height          =   3000
            Left            =   -74940
            TabIndex        =   7
            Top             =   45
            Width           =   7335
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmComp.frx":3F37
               TabIndex        =   34
               Top             =   210
               Width           =   7140
            End
         End
      End
      Begin TabDlg.SSTab tabEzeB 
         Height          =   3330
         Left            =   -74850
         TabIndex        =   9
         Top             =   360
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
         TabPicture(0)   =   "frmComp.frx":4396
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frEzeB0"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmComp.frx":43B2
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeb1"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmComp.frx":43CE
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frEzeB2"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick."
         TabPicture(3)   =   "frmComp.frx":43EA
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprEzeB(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom. Imp. x Pax"
         TabPicture(4)   =   "frmComp.frx":4406
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprEzeB(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2700
            Index           =   3
            Left            =   -74805
            OleObjectBlob   =   "frmComp.frx":4422
            TabIndex        =   39
            Top             =   195
            Width           =   7140
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2700
            Index           =   4
            Left            =   -74865
            OleObjectBlob   =   "frmComp.frx":4881
            TabIndex        =   40
            Top             =   210
            Width           =   7140
         End
         Begin VB.Frame frEzeB2 
            Height          =   2910
            Left            =   -74895
            TabIndex        =   13
            Top             =   45
            Width           =   7320
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmComp.frx":4CE0
               TabIndex        =   38
               Top             =   135
               Width           =   7140
            End
         End
         Begin VB.Frame frEzeb1 
            Height          =   2925
            Left            =   -74895
            TabIndex        =   12
            Top             =   45
            Width           =   7305
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   1
               Left            =   75
               OleObjectBlob   =   "frmComp.frx":513F
               TabIndex        =   37
               Top             =   165
               Width           =   7140
            End
         End
         Begin VB.Frame frEzeB0 
            Height          =   2985
            Left            =   90
            TabIndex        =   10
            Top             =   30
            Width           =   7305
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   0
               Left            =   105
               OleObjectBlob   =   "frmComp.frx":559E
               TabIndex        =   11
               Top             =   210
               Width           =   7140
            End
         End
      End
      Begin TabDlg.SSTab tabAep 
         Height          =   3390
         Left            =   -74865
         TabIndex        =   14
         Top             =   330
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   5980
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
         TabPicture(0)   =   "frmComp.frx":59F7
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frAep0"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmComp.frx":5A13
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frAep1"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmComp.frx":5A2F
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frAep2"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick."
         TabPicture(3)   =   "frmComp.frx":5A4B
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprAep(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom. Imp. x Pax"
         TabPicture(4)   =   "frmComp.frx":5A67
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprAep(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprAep 
            Height          =   2700
            Index           =   3
            Left            =   -74880
            OleObjectBlob   =   "frmComp.frx":5A83
            TabIndex        =   43
            Top             =   255
            Width           =   7140
         End
         Begin FPSpread.vaSpread sprAep 
            Height          =   2700
            Index           =   4
            Left            =   -74895
            OleObjectBlob   =   "frmComp.frx":5EE2
            TabIndex        =   44
            Top             =   225
            Width           =   7140
         End
         Begin VB.Frame frAep0 
            Height          =   2985
            Left            =   105
            TabIndex        =   15
            Top             =   30
            Width           =   7305
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmComp.frx":6341
               TabIndex        =   16
               Top             =   195
               Width           =   7140
            End
         End
         Begin VB.Frame frAep2 
            Height          =   2985
            Left            =   -74910
            TabIndex        =   18
            Top             =   45
            Width           =   7275
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   2
               Left            =   75
               OleObjectBlob   =   "frmComp.frx":679A
               TabIndex        =   42
               Top             =   195
               Width           =   7140
            End
         End
         Begin VB.Frame frAep1 
            Height          =   2970
            Left            =   -74895
            TabIndex        =   17
            Top             =   45
            Width           =   7260
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   1
               Left            =   60
               OleObjectBlob   =   "frmComp.frx":6BF9
               TabIndex        =   41
               Top             =   165
               Width           =   7140
            End
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1500
      Left            =   7845
      TabIndex        =   1
      Top             =   -45
      Width           =   1080
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
         Left            =   225
         Picture         =   "frmComp.frx":7058
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   180
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
         Left            =   225
         Picture         =   "frmComp.frx":715A
         Style           =   1  'Graphical
         TabIndex        =   46
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
         Left            =   225
         Picture         =   "frmComp.frx":725C
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   990
         Width           =   570
      End
   End
   Begin VB.Frame frdatos 
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   7710
      Begin VB.Frame Frame1 
         Caption         =   "Fechas de  Consultas para Comparación"
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
         Height          =   1485
         Left            =   105
         TabIndex        =   3
         Top             =   30
         Width           =   7485
         Begin VB.Frame Frame4 
            Caption         =   "Periodo Actual"
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
            Height          =   1095
            Left            =   255
            TabIndex        =   26
            Top             =   240
            Width           =   3540
            Begin VB.CommandButton botHelpFH 
               Height          =   345
               Left            =   2865
               Picture         =   "frmComp.frx":7A7E
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2880
               Picture         =   "frmComp.frx":7BF0
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   255
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesde 
               Height          =   285
               Left            =   1590
               TabIndex        =   29
               Top             =   270
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-##-####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFHasta 
               Height          =   285
               Left            =   1605
               TabIndex        =   30
               Top             =   615
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-##-####"
               PromptChar      =   " "
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
               Left            =   345
               TabIndex        =   32
               Top             =   615
               Width           =   1185
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
               Left            =   360
               TabIndex        =   31
               Top             =   270
               Width           =   1185
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Periodo Anterior"
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
            Height          =   1080
            Left            =   3900
            TabIndex        =   19
            Top             =   240
            Width           =   3330
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2655
               Picture         =   "frmComp.frx":7D62
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   600
               Width           =   375
            End
            Begin VB.CommandButton botHelpFDAnt 
               Height          =   345
               Left            =   2655
               Picture         =   "frmComp.frx":7ED4
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   240
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesdeAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   22
               Top             =   255
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-##-####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFHastaAnt 
               Height          =   285
               Left            =   1395
               TabIndex        =   23
               Top             =   600
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-##-####"
               PromptChar      =   " "
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
               Left            =   135
               TabIndex        =   25
               Top             =   600
               Width           =   1185
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
               Index           =   2
               Left            =   150
               TabIndex        =   24
               Top             =   255
               Width           =   1185
            End
         End
      End
   End
End
Attribute VB_Name = "frmComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset
Dim RsDataAnt  As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas
                       'con porcentajes
Dim UsoPorIntB As Boolean
Dim UsoPorAero As Boolean
Dim UsoPorInterior As Boolean
Dim UsoPorTotal As Boolean


Private Function L_Armarcondicion(anio As Integer, Mes As Integer) As String
'Anio = 0 es actual
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant

'If optFechas(0).Value Then
    If anio = 0 Then
        fechaDesde = mskFDesde.FormattedText
        fechaHasta = mskFHasta.FormattedText
    Else
        fechaDesde = mskFDesdeAnt.FormattedText
        fechaHasta = mskFHastaAnt.FormattedText
    End If
'End If


Cond = " WHERE fch_vta between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

L_Armarcondicion = Cond

End Function


Private Sub L_DecoEspigonAnt()
Dim fecha As String
Dim i As Integer, indEzea As Long, indEzeB As Long, indAep As Long, indInt As Long

Dim j As Integer
Dim dato As Variant
Dim Locale As Variant

For i = 0 To 2
    sprEzeA(i).Col = 4
    sprEzeB(i).Col = 4
    sprAEP(i).Col = 4
    sprInt(i).Col = 4
Next

indEzea = 1
indEzeB = 1
indAep = 1
indInt = 1
Do While Not RsDataAnt.EOF

        Select Case RsDataAnt!cod_depn
            Case DSLoc(1).Dep
                sprAEP(0).GetText 1, indAep, Locale
                If Locale = Trim(RsDataAnt!cod_local) & "-" & Trim(str(RsDataAnt!Cod_Sloc)) Then
                    sprAEP(0).SetText 3, indAep, str(RsDataAnt!imp)
                    sprAEP(1).SetText 3, indAep, str(RsDataAnt!cant_t)
                    sprAEP(2).SetText 3, indAep, str(RsDataAnt!cant_p)
                    sprAEP(3).SetText 3, indAep, str(L_Division(RsDataAnt!imp, RsDataAnt!cant_t))
                    sprAEP(4).SetText 3, indAep, str(L_Division(RsDataAnt!imp, RsDataAnt!cant_p))
                                  
                    For i = 0 To 2
                        Spread.spread_ResaltarCelda sprAEP(i), 4, indAep
                        Spread.spread_ResaltarCelda sprAEP(i), 5, indAep
                    Next
                RsDataAnt.MoveNext
                End If
                indAep = indAep + 1
                If indAep > sprAEP(0).MaxRows Then
                    RsDataAnt.MoveNext
                End If
            Case DSLoc(2).Dep
                Select Case RsDataAnt!Cod_Sdep
                    Case DSLoc(2).Sdep
                        sprEzeA(0).GetText 1, indEzea, Locale
                        If Locale = Trim(RsDataAnt!cod_local) & "-" & Trim(str(RsDataAnt!Cod_Sloc)) Then
                            sprEzeA(0).SetText 3, indEzea, str(RsDataAnt!imp)
                            sprEzeA(1).SetText 3, indEzea, str(RsDataAnt!cant_t)
                            sprEzeA(2).SetText 3, indEzea, str(RsDataAnt!cant_p)
                            sprEzeA(3).SetText 3, indEzea, str(L_Division(RsDataAnt!imp, RsDataAnt!cant_t))
                            sprEzeA(4).SetText 3, indEzea, str(L_Division(RsDataAnt!imp, RsDataAnt!cant_p))
                        
                        For i = 0 To 2
                            Spread.spread_ResaltarCelda sprEzeA(i), 4, indEzea
                            Spread.spread_ResaltarCelda sprEzeA(i), 5, indEzea
                        Next
                        
                        RsDataAnt.MoveNext
                        End If
                        indEzea = indEzea + 1
                        If indEzea > sprEzeA(0).MaxRows Then
                            RsDataAnt.MoveNext
                        End If
                    
                    Case DSLoc(3).Sdep
                        sprEzeB(0).GetText 1, indEzeB, Locale
                        If Locale = Trim(RsDataAnt!cod_local) & "-" & Trim(str(RsDataAnt!Cod_Sloc)) Then
                            sprEzeB(0).SetText 3, indEzeB, str(RsDataAnt!imp)
                            sprEzeB(1).SetText 3, indEzeB, str(RsDataAnt!cant_t)
                            sprEzeB(2).SetText 3, indEzeB, str(RsDataAnt!cant_p)
                            sprEzeB(3).SetText 3, indEzeB, str(L_Division(RsDataAnt!imp, RsDataAnt!cant_t))
                            sprEzeB(4).SetText 3, indEzeB, str(L_Division(RsDataAnt!imp, RsDataAnt!cant_p))
                        
                        For i = 0 To 2
                            Spread.spread_ResaltarCelda sprEzeB(i), 4, indEzeB
                            Spread.spread_ResaltarCelda sprEzeB(i), 5, indEzeB
                        Next
                        
                        RsDataAnt.MoveNext
                        If indEzeB > sprEzeB(0).MaxRows Then
                            RsDataAnt.MoveNext
                        End If
                        
                        End If
                        indEzeB = indEzeB + 1
                End Select
                
            Case DSLoc(4).Dep
                sprInt(0).GetText 1, indInt, Locale
                If Locale = Trim(RsDataAnt!cod_local) & "-" & Trim(str(RsDataAnt!Cod_Sloc)) & " " & RsDataAnt!Cod_Sdep Then
                    sprInt(0).SetText 3, indInt, str(RsDataAnt!imp)
                    sprInt(1).SetText 3, indInt, str(RsDataAnt!cant_t)
                    sprInt(2).SetText 3, indInt, str(RsDataAnt!cant_p)
                    sprInt(3).SetText 3, indInt, str(L_Division(RsDataAnt!imp, RsDataAnt!cant_t))
                    sprInt(4).SetText 3, indInt, str(L_Division(RsDataAnt!imp, RsDataAnt!cant_p))
                                  
                    For i = 0 To 2
                        Spread.spread_ResaltarCelda sprInt(i), 4, indInt
                        Spread.spread_ResaltarCelda sprInt(i), 5, indInt
                    Next
                RsDataAnt.MoveNext
                End If
                indInt = indInt + 1
                If indInt > sprInt(0).MaxRows Then
                    RsDataAnt.MoveNext
                End If
                
        End Select
                
Loop
FuncLocal_PromediosFilaSPR sprEzeA(0), sprEzeA(2), sprEzeA(4), sprEzeA(0).MaxCols - 2
FuncLocal_PromediosFilaSPR sprEzeA(0), sprEzeA(1), sprEzeA(3), sprEzeA(0).MaxCols - 2

FuncLocal_PromediosFilaSPR sprEzeB(0), sprEzeB(2), sprEzeB(4), sprEzeB(0).MaxCols - 2
FuncLocal_PromediosFilaSPR sprEzeB(0), sprEzeB(1), sprEzeB(3), sprEzeB(0).MaxCols - 2

FuncLocal_PromediosFilaSPR sprAEP(0), sprAEP(2), sprAEP(4), sprAEP(0).MaxCols - 2
FuncLocal_PromediosFilaSPR sprAEP(0), sprAEP(1), sprAEP(3), sprAEP(0).MaxCols - 2

FuncLocal_PromediosFilaSPR sprInt(0), sprInt(2), sprInt(4), sprInt(0).MaxCols - 2
FuncLocal_PromediosFilaSPR sprInt(0), sprInt(1), sprInt(3), sprInt(0).MaxCols - 2

For i = 3 To 4
    For j = 1 To sprAEP(i).MaxRows
        Spread.spread_ResaltarCelda sprAEP(i), 5, j
        Spread.spread_ResaltarCelda sprAEP(i), 4, j
    Next
    For j = 1 To sprEzeA(i).MaxRows
        Spread.spread_ResaltarCelda sprEzeA(i), 5, j
        Spread.spread_ResaltarCelda sprEzeA(i), 4, j
    Next
    For j = 1 To sprEzeB(i).MaxRows
        Spread.spread_ResaltarCelda sprEzeB(i), 5, j
        Spread.spread_ResaltarCelda sprEzeB(i), 4, j
    Next

    For j = 1 To sprInt(i).MaxRows
        Spread.spread_ResaltarCelda sprInt(i), 5, j
        Spread.spread_ResaltarCelda sprInt(i), 4, j
    Next

Next

For i = 0 To 4
    Spread.spread_ResaltarCelda sprEzeA(i), 5, sprEzeA(i).MaxRows
    Spread.spread_ResaltarCelda sprEzeB(i), 5, sprEzeB(i).MaxRows
    Spread.spread_ResaltarCelda sprAEP(i), 5, sprAEP(i).MaxRows
    Spread.spread_ResaltarCelda sprInt(i), 5, sprInt(i).MaxRows
    
    Spread.spread_ResaltarCelda sprEzeA(i), 4, sprEzeA(i).MaxRows
    Spread.spread_ResaltarCelda sprEzeB(i), 4, sprEzeB(i).MaxRows
    Spread.spread_ResaltarCelda sprAEP(i), 4, sprAEP(i).MaxRows
    Spread.spread_ResaltarCelda sprInt(i), 4, sprInt(i).MaxRows
Next

End Sub

Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indEzea As Long, indEzeB As Long, indAep As Long, indInt As Long

Dim j As Integer
Dim dato As Variant
Dim Locale As Variant

For i = 0 To 2
    sprEzeA(i).Col = 4
    sprEzeB(i).Col = 4
    sprAEP(i).Col = 4
    sprInt(i).Col = 4
Next

indEzea = 1
indEzeB = 1
indAep = 1
indInt = 1
Do While Not RsData.EOF

        Select Case RsData!cod_depn
            Case DSLoc(1).Dep
                sprAEP(0).GetText 1, indAep, Locale
                If Locale = Trim(RsData!cod_local) & "-" & Trim(str(RsData!Cod_Sloc)) Then
                    sprAEP(0).SetText 2, indAep, str(RsData!imp)
                    sprAEP(1).SetText 2, indAep, str(RsData!cant_t)
                    sprAEP(2).SetText 2, indAep, str(RsData!cant_p)
                    sprAEP(3).SetText 2, indAep, 0
                    sprAEP(3).SetText 2, indAep, str(L_Division(RsData!imp, RsData!cant_t))
                    sprAEP(4).SetText 2, indAep, str(L_Division(RsData!imp, RsData!cant_p))
                                  

                RsData.MoveNext
                End If
                indAep = indAep + 1

            Case DSLoc(2).Dep
                Select Case RsData!Cod_Sdep
                    Case DSLoc(2).Sdep
                        sprEzeA(0).GetText 1, indEzea, Locale
                        If Locale = Trim(RsData!cod_local) & "-" & Trim(str(RsData!Cod_Sloc)) Then
                            sprEzeA(0).SetText 2, indEzea, str(RsData!imp)
                            sprEzeA(1).SetText 2, indEzea, str(RsData!cant_t)
                            sprEzeA(2).SetText 2, indEzea, str(RsData!cant_p)
                            sprEzeA(3).SetText 2, indEzea, str(L_Division(RsData!imp, RsData!cant_t))
                            sprEzeA(4).SetText 2, indEzea, str(L_Division(RsData!imp, RsData!cant_p))
                        
                        RsData.MoveNext
                        End If
                        indEzea = indEzea + 1
                    
                    Case DSLoc(3).Sdep
                        sprEzeB(0).GetText 1, indEzeB, Locale
                        If Locale = Trim(RsData!cod_local) & "-" & Trim(str(RsData!Cod_Sloc)) Then
                            sprEzeB(0).SetText 2, indEzeB, str(RsData!imp)
                            sprEzeB(1).SetText 2, indEzeB, str(RsData!cant_t)
                            sprEzeB(2).SetText 2, indEzeB, str(RsData!cant_p)
                            sprEzeB(3).SetText 2, indEzeB, str(L_Division(RsData!imp, RsData!cant_t))
                            sprEzeB(4).SetText 2, indEzeB, str(L_Division(RsData!imp, RsData!cant_p))
                        
                            RsData.MoveNext

                        End If
                        indEzeB = indEzeB + 1
                End Select
                
            Case DSLoc(4).Dep
                sprInt(0).GetText 1, indInt, Locale
                If Locale = Trim(RsData!cod_local) & "-" & Trim(str(RsData!Cod_Sloc)) & " " & RsData!Cod_Sdep Then
                    sprInt(0).SetText 2, indInt, str(RsData!imp)
                    sprInt(1).SetText 2, indInt, str(RsData!cant_t)
                    sprInt(2).SetText 2, indInt, str(RsData!cant_p)
                    sprInt(3).SetText 2, indInt, str(L_Division(RsData!imp, RsData!cant_t))
                    sprInt(4).SetText 2, indInt, str(L_Division(RsData!imp, RsData!cant_p))
                                  
                RsData.MoveNext
                End If
                indInt = indInt + 1

        End Select
                
Loop
For i = 0 To 2
Spread_TotalesGrillas sprEzeA(i), sprEzeA(i).MaxCols - 2, 2
Spread_TotalesGrillas sprEzeB(i), sprEzeB(i).MaxCols - 2, 2
Spread_TotalesGrillas sprAEP(i), sprAEP(i).MaxCols - 2, 2
Spread_TotalesGrillas sprInt(i), sprInt(i).MaxCols - 2, 2

Next

L_RefrescarFrames
End Sub



Public Function L_Division(divisor As Single, dividendo As Long) As Single

If dividendo = 0 Or divisor = 0 Then
   L_Division = 0
 Else
   L_Division = divisor / dividendo
End If

End Function

Private Sub L_LimpiarGrillas()
Dim i
    
For i = 0 To 4
    sprEzeA(i).MaxRows = 0
    sprEzeB(i).MaxRows = 0
    sprAEP(i).MaxRows = 0
    sprInt(i).MaxRows = 0
Next


    'sprTot.MaxRows = 0

    
End Sub

Public Sub L_LlenarLocales()
Dim sqlL As String
Dim n As Integer
Dim RsDataLoc As Recordset
Dim Sdep As String
Dim depn As String
Dim linea As Integer

sqlL = " SELECT distinct cod_depn,cod_sdep, decode(cod_loc,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22','L08','L22', cod_loc) cod_loc, decode(cod_loc,'L05',0,'L06',0,'L21',0,'L07',0,'L08',0,cod_sloc) cod_sloc "
sqlL = sqlL & " FROM ventas.local_sublocal Where cod_loc not in ('L05','L06')"
sqlL = sqlL & " ORDER BY cod_depn,cod_sdep,cod_loc,cod_sloc "

If Aplicacion.ObtenerRsDAO(sqlL, RsDataLoc) Then
   If Aplicacion.CantReg(RsData) > 0 Then
      For n = 0 To 4
       sprEzeA(n).MaxRows = 0
       sprEzeB(n).MaxRows = 0
       sprAEP(n).MaxRows = 0
       sprInt(n).MaxRows = 0
       'linea = SprRubro(n).MaxRows
      Next
      Do While Not RsDataLoc.EOF
          Select Case RsDataLoc!cod_depn
            Case DSLoc(1).Dep
                
                For n = 0 To 4
                 sprAEP(n).MaxRows = sprAEP(n).MaxRows + 1
                 sprAEP(n).SetText 1, sprAEP(n).MaxRows, Trim(RsDataLoc!cod_loc) & "-" & Trim(RsDataLoc!Cod_Sloc)
                Next n
            
            Case DSLoc(2).Dep
               Select Case RsDataLoc!Cod_Sdep
                  Case DSLoc(2).Sdep
                     For n = 0 To 4
                        sprEzeA(n).MaxRows = sprEzeA(n).MaxRows + 1
                        sprEzeA(n).SetText 1, sprEzeA(n).MaxRows, Trim(RsDataLoc!cod_loc) & "-" & Trim(RsDataLoc!Cod_Sloc)
                     Next n
                  Case DSLoc(3).Sdep
                
                     For n = 0 To 4
                        sprEzeB(n).MaxRows = sprEzeB(n).MaxRows + 1
                        sprEzeB(n).SetText 1, sprEzeB(n).MaxRows, Trim(RsDataLoc!cod_loc) & "-" & Trim(RsDataLoc!Cod_Sloc)
                     Next n
               End Select
               
            Case DSLoc(4).Dep
         
                For n = 0 To 4
                 sprInt(n).MaxRows = sprInt(n).MaxRows + 1
                 sprInt(n).SetText 1, sprInt(n).MaxRows, Trim(RsDataLoc!cod_loc) & "-" & Trim(RsDataLoc!Cod_Sloc) & " " & Trim(RsDataLoc!Cod_Sdep)
                Next n
          
          End Select
          
          RsDataLoc.MoveNext
     Loop
   End If
End If

End Sub


Private Sub L_Porcentages(ByRef sprPorc As control, sprDato As control, Orientacion As String)
Dim i As Integer, j As Integer
Dim valor As Variant
Dim tot As Variant, dato As Variant
Dim Result As Double

On Error GoTo ErrPorc:

sprPorc.MaxRows = 0

Select Case Orientacion
    Case "COL"
        For i = 1 To sprDato.MaxRows
         If i = sprDato.MaxRows Then
            sprPorc.MaxRows = i
            For j = 2 To sprDato.MaxCols
                sprPorc.SetText j, i, "100"
            Next
         Else
            sprDato.GetText 1, i, dato
            Result = 0
            sprPorc.MaxRows = i
        
            sprPorc.SetText 1, i, Trim(str(dato))
            
            For j = 2 To sprDato.MaxCols
                sprDato.GetText j, sprDato.MaxRows, tot
                sprDato.GetText j, i, valor
                If tot <> 0 Then
                    Result = valor * 100 / tot
                End If
                sprPorc.SetText j, i, Format$(Result, "% 0.00")
            Next
         End If
        Next

End Select
        L_TotalesLinea sprPorc
ErrPorc:
    Exit Sub
End Sub

Private Sub L_RefrescarFrames()
    frEzeA0.Refresh
    frEzeA1.Refresh
    frEzeA2.Refresh
    frEzeB0.Refresh
    frEzeB1.Refresh
    frEzeB2.Refresh
    frAep0.Refresh
    frAep1.Refresh
    frAep2.Refresh
    FrInt0.Refresh
    FrInt1.Refresh
    FrInt2.Refresh
End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset

Aplicacion.SeteoProceso ""


sql = " SELECT "
sql = sql & " cod_depn, "
sql = sql & " cod_sdep, "
sql = sql & "  decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local) cod_local, "
sql = sql & " decode(cod_local,'L05',0,'L06',0,'L21',0,'L07',0,cod_sloc) cod_sloc, "
sql = sql & " sum(cant_tickets) cant_t, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cant_pax) cant_p "
sql = sql & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)))
sql = sql & L_Armarcondicion(0, 1)
sql = sql & "group by cod_depn,cod_sdep, decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local),decode(cod_local,'L05',0,'L06',0,'L21',0,'L07',0,cod_sloc) "
sql = sql & " order by cod_depn,cod_sdep,cod_local,cod_sloc"

sqlX = " SELECT "
sqlX = sqlX & " cod_depn, "
sqlX = sqlX & " cod_sdep, "
sqlX = sqlX & "  decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local) cod_local, "
sqlX = sqlX & " decode(cod_local,'L05',0,'L06',0,'L21',0,'L07',0,cod_sloc) cod_sloc, "
sqlX = sqlX & " sum(cant_tickets) cant_t, "
sqlX = sqlX & " sum(importe) imp, "
sqlX = sqlX & " sum(cant_pax) cant_p "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesdeAnt.FormattedText)))
sqlX = sqlX & L_Armarcondicion(1, 1)
sqlX = sqlX & "group by cod_depn,cod_sdep, decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local),decode(cod_local,'L05',0,'L06',0,'L21',0,'L07',0,cod_sloc) "
sqlX = sqlX & " order by cod_depn,cod_sdep,cod_local,cod_sloc"


If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        L_LlenarLocales
        L_DecoEspigon
        If Aplicacion.ObtenerRsDAO(sqlX, RsDataAnt) Then
            L_DecoEspigonAnt
        End If
        Aplicacion.CerrarDAO RsDataAnt
    End If
    
    Aplicacion.CerrarDAO RsData

    
End If

Aplicacion.SeteoFin

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
        Case INTE
            UsoPorInterior = True
        Case Total
            UsoPorTotal = True
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
        Case INTE
            UsoPorInterior = False
        Case Total
            UsoPorTotal = False
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorAero Or UsoPorInterior Or UsoPorTotal) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub


Private Sub L_TotalesLinea(spr As control)

    'Select a block of cells
    spr.Col = 1
    spr.Row = spr.MaxRows
    spr.Col2 = 9
    spr.Row2 = spr.MaxRows
    spr.BlockMode = True

    'Determine the color of background, foreground and border color
    spr.ForeColor = RGB(0, 0, 255)
    spr.BackColor = RGB(242, 242, 242)
    spr.CellBorderColor = RGB(255, 255, 255)
    
    'Turn block mode off
    spr.BlockMode = False

End Sub


Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        L_Refrescar
    Case 1
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
Select Case TabAep.Tab
    Case 0
        L_TratarExcel sprAEP(0), "AEROPARQUE ", "Informe Comparativo ", "Importes"
    Case 1
        L_TratarExcel sprAEP(1), "AEROPARQUE", "Informe Comparativo ", "Tickets"
    Case 2
        L_TratarExcel sprAEP(2), "AEROPARQUE", "Informe Comparativo ", "Pasajeros"
    Case 3
        L_TratarExcel sprAEP(3), "AEROPARQUE", "Informe Comparativo ", "Promedios por Tickets"
    Case 4
        L_TratarExcel sprAEP(4), "AEROPARQUE", "Informe Comparativo ", "Promedios por Pasajeros"
End Select

End Sub

Private Sub botExcelInt_Click()
Select Case TabINT.Tab
    Case 0
        L_TratarExcel sprInt(0), "INTERIOR", "Informe Comparativo ", "Importes"
    Case 1
        L_TratarExcel sprInt(1), "INTERIOR", "Informe Comparativo ", "Tickets"
    Case 2
        L_TratarExcel sprInt(2), "INTERIOR", "Informe Comparativo ", "Pasajeros"
    Case 3
        L_TratarExcel sprInt(3), "INTERIOR", "Informe Comparativo ", "Promedios por Tickets"
    Case 4
        L_TratarExcel sprInt(4), "INTERIOR", "Informe Comparativo ", "Promedios por Pasajeros"
End Select

End Sub


Private Sub botExelEzeB_Click()
Select Case TabEzeB.Tab
    Case 0
        L_TratarExcel sprEzeB(0), "ESPIGON INTERNACIONAL B", "Informe Comparativo ", "Importes"
    Case 1
        L_TratarExcel sprEzeB(1), "ESPIGON INTERNACIONAL B", "Informe Comparativo ", "Tickets"
    Case 2
        L_TratarExcel sprEzeB(2), "ESPIGON INTERNACIONAL B", "Informe Comparativo ", "Pasajeros"
    Case 3
        L_TratarExcel sprEzeB(3), "ESPIGON INTERNACIONAL B", "Informe Comparativo ", "Promedios por Tickets"
    Case 4
        L_TratarExcel sprEzeB(4), "ESPIGON INTERNACIONAL B", "Informe Comparativo ", "Promedios por Pasajeros"
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

Private Sub botHelpFDAnt_Click()
Dim fch As Date

If mskFDesdeAnt.Text <> "" Then
    fch = mskFDesdeAnt.FormattedText
Else
    fch = Date
End If

frmFecha.MuestroFormFecha fch

mskFDesdeAnt.Text = Format$(fch, FTOFECHA)

mskFDesdeAnt.SetFocus

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


Private Sub botHelpFHAnt_Click()
Dim fch As Date

If mskFHastaAnt.Text <> "" Then
    fch = mskFHastaAnt.FormattedText
Else
    fch = Date
End If

frmFecha.MuestroFormFecha fch

mskFHastaAnt.Text = Format$(fch, FTOFECHA)

mskFHastaAnt.SetFocus

End Sub




Private Sub ExcelIntA_Click()
Select Case TabEzeA.Tab
    Case 0
        L_TratarExcel sprEzeA(0), "ESPIGON INTERNACIONAL A", "Informe Comparativo ", "Importes"
    Case 1
        L_TratarExcel sprEzeA(1), "ESPIGON INTERNACIONAL A", "Informe Comparativo ", "Tickets"
    Case 2
        L_TratarExcel sprEzeA(2), "ESPIGON INTERNACIONAL A", "Informe Comparativo ", "Pasajeros"
    Case 3
        L_TratarExcel sprEzeA(3), "ESPIGON INTERNACIONAL A", "Informe Comparativo ", "Promedios por Tickets"
    Case 4
        L_TratarExcel sprEzeA(4), "ESPIGON INTERNACIONAL A", "Informe Comparativo ", "Promedios por Pasajeros"
End Select


End Sub

Private Sub Form_Activate()
FuncLocal_SeteoTABS tabEspigon

End Sub

Private Sub L_TratarExcel(spr As control, titulo As String, subTit As String, Info As String)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim i
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmComp.caption = Aplicacion.SeteoProceso(frmComp.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    AppExcel.Application.Visible = True
    
    ReDim titCol(spr.MaxCols)
    Col = 1
    fila = 3
    
    For i = 1 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 1, 1, titulo
    rango = Exl_rangos(1, 1, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    rango = Exl_rangos(fila, fila + 5, 1, 1)
    
    Exl_PonerValor AppExcel, fila, Col, "Información sobre :" & Info
    fila = fila + 2
    
    Exl_PonerValor AppExcel, fila, Col, "Periodo Anrerior : del " & mskFDesdeAnt.FormattedText & " al " & mskFHastaAnt.FormattedText
    fila = fila + 1
    
    Exl_PonerValor AppExcel, fila, Col, "Periodo Actual :  del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText
    
    Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
    
    fila = fila + 2
    
    Exl_BajarGrillaExel spr, AppExcel, fila, Col, titCol
    fila = fila + spr.MaxRows
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    
    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmComp.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6000
Me.Width = 9300

mskFDesde.Text = "01-" & Format$(Month(Date), "0#") & "-" & Format$(Year(Date), "####")
mskFHasta.Text = Format$(Date - 1, FTOFECHA)

mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio, FTOFECHA)

L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmComp"

End Sub


Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmComp"
End Sub


Private Sub mskFDesde_LostFocus()

    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    'ElseIf (Year(mskFDesde.FormattedText) < Year(Date)) Then
    '    mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    End If
    
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
    If mskFHasta.Text <> "" Then
        If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
            mskFHasta.Text = mskFDesde.Text
        End If
    End If

End Sub


Private Sub mskFDesdeAnt_LostFocus()
    If Not IsDate(mskFDesdeAnt.FormattedText) Then
        mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - 7, FTOFECHA)
    'ElseIf (CDate(mskFDesdeAnt.FormattedText) > CDate(mskFDesde.FormattedText) - 7) Then
    '    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - 7, FTOFECHA)
    End If
    
    mskFDesdeAnt.Text = Format$(mskFDesdeAnt.FormattedText, FTOFECHA)
    If mskFHastaAnt.Text <> "" Then
        If CDate(mskFHastaAnt.FormattedText) < CDate(mskFDesdeAnt.FormattedText) _
        Or Year(mskFHastaAnt.FormattedText) <> Year(mskFDesdeAnt.FormattedText) Then
            mskFHastaAnt.Text = mskFDesdeAnt.Text
        End If
    End If
End Sub


Private Sub mskFHasta_LostFocus()

If Not IsDate(mskFHasta.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
ElseIf CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
End If

mskFHasta.Text = Format$(mskFHasta.FormattedText, FTOFECHA)
mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio, FTOFECHA)

End Sub


Private Sub mskFHastaAnt_LostFocus()

If Not IsDate(mskFHastaAnt.FormattedText) Then
        mskFHastaAnt.Text = mskFDesdeAnt.Text
ElseIf CDate(mskFHastaAnt.FormattedText) < CDate(mskFDesdeAnt.FormattedText) _
    Or Year(mskFHastaAnt.FormattedText) <> Year(mskFDesdeAnt.FormattedText) Then
        mskFHastaAnt.Text = mskFDesdeAnt.Text
End If

mskFHastaAnt.Text = Format$(mskFHastaAnt.FormattedText, FTOFECHA)

End Sub


