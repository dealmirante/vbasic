VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLGC 
   Caption         =   "Ventas por Local"
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
      TabIndex        =   12
      Top             =   1575
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   6720
      _Version        =   327680
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   6
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
      TabPicture(0)   =   "frmLGC.frx":0000
      Tab(0).ControlCount=   0
      Tab(0).ControlEnabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmLGC.frx":001C
      Tab(1).ControlCount=   7
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "botGrTortaEA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "botGrafEvoEZEA"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ExcelIntA"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "botPorc(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "botPorc(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "botPorc(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "tabEzeA"
      Tab(1).Control(6).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmLGC.frx":0038
      Tab(2).ControlCount=   8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "botImpr"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "botPorcB(2)"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "botPorcB(1)"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "botPorcB(0)"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "botExelEzeB"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "botGrafEvoEZEB"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "botGrTortaEB"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "tabEzeB"
      Tab(2).Control(7).Enabled=   0   'False
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmLGC.frx":0054
      Tab(3).ControlCount=   7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "botPorcAep(2)"
      Tab(3).Control(0).Enabled=   -1  'True
      Tab(3).Control(1)=   "botPorcAep(1)"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "botPorcAep(0)"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "botExcelAep"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "botGrafEvoAEP"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "botGrTortaAEP"
      Tab(3).Control(5).Enabled=   -1  'True
      Tab(3).Control(6)=   "tabAep"
      Tab(3).Control(6).Enabled=   0   'False
      TabCaption(4)   =   "INTERIOR"
      TabPicture(4)   =   "frmLGC.frx":0070
      Tab(4).ControlCount=   7
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TabInt"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "botGrTortaINT"
      Tab(4).Control(1).Enabled=   -1  'True
      Tab(4).Control(2)=   "botGrafEvoINT"
      Tab(4).Control(2).Enabled=   -1  'True
      Tab(4).Control(3)=   "botExcelINT"
      Tab(4).Control(3).Enabled=   -1  'True
      Tab(4).Control(4)=   "botPorcInt(0)"
      Tab(4).Control(4).Enabled=   -1  'True
      Tab(4).Control(5)=   "botPorcInt(1)"
      Tab(4).Control(5).Enabled=   -1  'True
      Tab(4).Control(6)=   "botPorcInt(2)"
      Tab(4).Control(6).Enabled=   -1  'True
      Begin TabDlg.SSTab TabInt 
         Height          =   3315
         Left            =   -74865
         TabIndex        =   82
         Top             =   330
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   5847
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   441
         ForeColor       =   16711680
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmLGC.frx":008C
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FrINTPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "FrINT0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmLGC.frx":00A8
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FrINTPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "FrINT1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmLGC.frx":00C4
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FrINTPorc2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "FrINT2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick"
         TabPicture(3)   =   "frmLGC.frx":00E0
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprInt(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom Imp x pax"
         TabPicture(4)   =   "frmLGC.frx":00FC
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprInt(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprInt 
            Height          =   2805
            Index           =   3
            Left            =   -74910
            OleObjectBlob   =   "frmLGC.frx":0118
            TabIndex        =   104
            Top             =   150
            Width           =   7200
         End
         Begin FPSpread.vaSpread sprInt 
            Height          =   2835
            Index           =   4
            Left            =   -74850
            OleObjectBlob   =   "frmLGC.frx":07C4
            TabIndex        =   105
            Top             =   135
            Width           =   7140
         End
         Begin VB.Frame FrINT0 
            Height          =   2895
            Left            =   135
            TabIndex        =   91
            Top             =   90
            Width           =   7185
            Begin FPSpread.vaSpread sprInt 
               Height          =   2685
               Index           =   0
               Left            =   75
               OleObjectBlob   =   "frmLGC.frx":0E70
               TabIndex        =   92
               Top             =   150
               Width           =   7050
            End
         End
         Begin VB.Frame FrINT1 
            Height          =   2910
            Left            =   -74880
            TabIndex        =   95
            Top             =   60
            Width           =   7185
            Begin FPSpread.vaSpread sprInt 
               Height          =   2685
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmLGC.frx":157B
               TabIndex        =   96
               Top             =   165
               Width           =   7065
            End
         End
         Begin VB.Frame FrINTPorc0 
            Height          =   2940
            Left            =   135
            TabIndex        =   89
            Top             =   45
            Width           =   7155
            Begin FPSpread.vaSpread sprINTPorc 
               Height          =   2700
               Index           =   0
               Left            =   165
               OleObjectBlob   =   "frmLGC.frx":1BF1
               TabIndex        =   90
               Top             =   165
               Width           =   6915
            End
         End
         Begin VB.Frame FrINTPorc1 
            Height          =   2910
            Left            =   -74895
            TabIndex        =   93
            Top             =   60
            Width           =   7200
            Begin FPSpread.vaSpread sprINTPorc 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmLGC.frx":2031
               TabIndex        =   94
               Top             =   165
               Width           =   7050
            End
         End
         Begin VB.Frame FrINT2 
            Height          =   2925
            Left            =   -74925
            TabIndex        =   99
            Top             =   45
            Width           =   7245
            Begin FPSpread.vaSpread sprInt 
               Height          =   2700
               Index           =   2
               Left            =   105
               OleObjectBlob   =   "frmLGC.frx":2473
               TabIndex        =   100
               Top             =   150
               Width           =   7110
            End
         End
         Begin VB.Frame FrINTPorc2 
            Height          =   2910
            Left            =   -74910
            TabIndex        =   97
            Top             =   60
            Width           =   7230
            Begin FPSpread.vaSpread sprINTPorc 
               Height          =   2685
               Index           =   2
               Left            =   75
               OleObjectBlob   =   "frmLGC.frx":2AE9
               TabIndex        =   98
               Top             =   165
               Width           =   7110
            End
         End
      End
      Begin TabDlg.SSTab tabEzeA 
         Height          =   3330
         Left            =   120
         TabIndex        =   13
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
         TabPicture(0)   =   "frmLGC.frx":2F2B
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frEzeAPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frEzeA0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmLGC.frx":2F47
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeAPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frEzeA1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmLGC.frx":2F63
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FrEZEA2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "FrEZEAporc2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick."
         TabPicture(3)   =   "frmLGC.frx":2F7F
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprEzeA(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom. Imp. x Pax"
         TabPicture(4)   =   "frmLGC.frx":2F9B
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprEzeA(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2850
            Index           =   4
            Left            =   -74895
            OleObjectBlob   =   "frmLGC.frx":2FB7
            TabIndex        =   55
            Top             =   105
            Width           =   7200
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2865
            Index           =   3
            Left            =   -74910
            OleObjectBlob   =   "frmLGC.frx":378C
            TabIndex        =   54
            Top             =   105
            Width           =   7230
         End
         Begin VB.Frame frEzeA0 
            Height          =   2970
            Left            =   105
            TabIndex        =   16
            Top             =   30
            Width           =   7245
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   0
               Left            =   60
               OleObjectBlob   =   "frmLGC.frx":3F73
               TabIndex        =   28
               Top             =   180
               Width           =   7170
            End
         End
         Begin VB.Frame frEzeA1 
            Height          =   2970
            Left            =   -74910
            TabIndex        =   19
            Top             =   45
            Width           =   7305
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2715
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmLGC.frx":4773
               TabIndex        =   31
               Top             =   165
               Width           =   7155
            End
         End
         Begin VB.Frame frEzeAPorc1 
            Height          =   2940
            Left            =   -74910
            TabIndex        =   23
            Top             =   45
            Width           =   7305
            Begin FPSpread.vaSpread sprEzeAporc 
               Height          =   2715
               Index           =   1
               Left            =   60
               OleObjectBlob   =   "frmLGC.frx":4F78
               TabIndex        =   34
               Top             =   180
               Width           =   7155
            End
         End
         Begin VB.Frame FrEZEA2 
            Height          =   2970
            Left            =   -74895
            TabIndex        =   103
            Top             =   45
            Width           =   7275
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2715
               Index           =   2
               Left            =   60
               OleObjectBlob   =   "frmLGC.frx":54F1
               TabIndex        =   106
               Top             =   165
               Width           =   7155
            End
         End
         Begin VB.Frame FrEZEAporc2 
            Height          =   2940
            Left            =   -74895
            TabIndex        =   101
            Top             =   45
            Width           =   7275
            Begin FPSpread.vaSpread sprEzeAporc 
               Height          =   2730
               Index           =   2
               Left            =   60
               OleObjectBlob   =   "frmLGC.frx":5CF7
               TabIndex        =   102
               Top             =   150
               Width           =   7155
            End
         End
         Begin VB.Frame frEzeAPorc0 
            Height          =   2940
            Left            =   105
            TabIndex        =   22
            Top             =   45
            Width           =   7305
            Begin FPSpread.vaSpread sprEzeAporc 
               Height          =   2730
               Index           =   0
               Left            =   60
               OleObjectBlob   =   "frmLGC.frx":6269
               TabIndex        =   33
               Top             =   150
               Width           =   7185
            End
         End
      End
      Begin VB.CommandButton botGrTortaINT 
         Height          =   495
         Left            =   -67065
         Picture         =   "frmLGC.frx":67DA
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   1995
         Width           =   765
      End
      Begin VB.CommandButton botGrafEvoINT 
         Height          =   540
         Left            =   -67080
         Picture         =   "frmLGC.frx":6C1C
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   2535
         Width           =   765
      End
      Begin VB.CommandButton botExcelINT 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67080
         Picture         =   "frmLGC.frx":705E
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3120
         Width           =   765
      End
      Begin VB.CommandButton botPorcInt 
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
         Height          =   500
         Index           =   0
         Left            =   -67095
         Picture         =   "frmLGC.frx":75F0
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   390
         Width           =   765
      End
      Begin VB.CommandButton botPorcInt 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   -67080
         Picture         =   "frmLGC.frx":7F62
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   915
         Width           =   765
      End
      Begin VB.CommandButton botPorcInt 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   2
         Left            =   -67080
         Picture         =   "frmLGC.frx":85B8
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   1455
         Width           =   765
      End
      Begin VB.CommandButton botImpr 
         Caption         =   "Command1"
         Height          =   345
         Left            =   -66390
         TabIndex        =   81
         Top             =   3270
         Width           =   225
      End
      Begin VB.CommandButton botPorcAep 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   2
         Left            =   -67200
         Picture         =   "frmLGC.frx":8C7E
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1485
         Width           =   765
      End
      Begin VB.CommandButton botPorcAep 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   -67200
         Picture         =   "frmLGC.frx":9344
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   960
         Width           =   765
      End
      Begin VB.CommandButton botPorcAep 
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
         Height          =   500
         Index           =   0
         Left            =   -67200
         Picture         =   "frmLGC.frx":999A
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   420
         Width           =   765
      End
      Begin VB.CommandButton botPorcB 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   2
         Left            =   -67200
         Picture         =   "frmLGC.frx":A30C
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1530
         Width           =   765
      End
      Begin VB.CommandButton botPorcB 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   -67200
         Picture         =   "frmLGC.frx":A9D2
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   990
         Width           =   765
      End
      Begin VB.CommandButton botPorcB 
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
         Height          =   500
         Index           =   0
         Left            =   -67200
         Picture         =   "frmLGC.frx":B028
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   465
         Width           =   765
      End
      Begin VB.CommandButton botPorc 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   2
         Left            =   7800
         Picture         =   "frmLGC.frx":B99A
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1500
         Width           =   765
      End
      Begin VB.CommandButton botPorc 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   7800
         Picture         =   "frmLGC.frx":C060
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   945
         Width           =   765
      End
      Begin VB.CommandButton botPorc 
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
         Height          =   500
         Index           =   0
         Left            =   7815
         Picture         =   "frmLGC.frx":C6B6
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   435
         Width           =   765
      End
      Begin VB.CommandButton botExcelAep 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67185
         Picture         =   "frmLGC.frx":D028
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3150
         Width           =   765
      End
      Begin VB.CommandButton botExelEzeB 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67185
         Picture         =   "frmLGC.frx":D5BA
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3195
         Width           =   765
      End
      Begin VB.CommandButton ExcelIntA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   7785
         Picture         =   "frmLGC.frx":DB4C
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3150
         Width           =   780
      End
      Begin VB.CommandButton botGrafEvoAEP 
         Height          =   540
         Left            =   -67200
         Picture         =   "frmLGC.frx":E0DE
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2565
         Width           =   765
      End
      Begin VB.CommandButton botGrafEvoEZEB 
         Height          =   540
         Left            =   -67200
         Picture         =   "frmLGC.frx":E520
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2610
         Width           =   765
      End
      Begin VB.CommandButton botGrafEvoEZEA 
         Height          =   540
         Left            =   7785
         Picture         =   "frmLGC.frx":E962
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2565
         Width           =   780
      End
      Begin VB.CommandButton botGrTortaAEP 
         Height          =   495
         Left            =   -67200
         Picture         =   "frmLGC.frx":EDA4
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2025
         Width           =   765
      End
      Begin VB.CommandButton botGrTortaEB 
         Height          =   495
         Left            =   -67200
         Picture         =   "frmLGC.frx":F1E6
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2070
         Width           =   765
      End
      Begin VB.CommandButton botGrTortaEA 
         Height          =   495
         Left            =   7785
         Picture         =   "frmLGC.frx":F628
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2025
         Width           =   780
      End
      Begin TabDlg.SSTab tabEzeB 
         Height          =   3330
         Left            =   -74925
         TabIndex        =   14
         Top             =   345
         Width           =   7335
         _ExtentX        =   12938
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
         TabPicture(0)   =   "frmLGC.frx":FA6A
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frEzeBPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frEzeB0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmLGC.frx":FA86
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeBPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frEzeB1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmLGC.frx":FAA2
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frEzeBPorc2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "frEzeB2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom. Imp x Tick"
         TabPicture(3)   =   "frmLGC.frx":FABE
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprEzeB(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom. Imp x Pax"
         TabPicture(4)   =   "frmLGC.frx":FADA
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprEzeB(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2700
            Index           =   4
            Left            =   -74820
            OleObjectBlob   =   "frmLGC.frx":FAF6
            TabIndex        =   57
            Top             =   210
            Width           =   6600
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2700
            Index           =   3
            Left            =   -74790
            OleObjectBlob   =   "frmLGC.frx":1008F
            TabIndex        =   56
            Top             =   165
            Width           =   6600
         End
         Begin VB.Frame frEzeB2 
            Height          =   2940
            Left            =   -74925
            TabIndex        =   41
            Top             =   45
            Width           =   7110
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   2
               Left            =   225
               OleObjectBlob   =   "frmLGC.frx":10628
               TabIndex        =   44
               Top             =   150
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeB1 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   20
            Top             =   60
            Width           =   6810
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   1
               Left            =   240
               OleObjectBlob   =   "frmLGC.frx":10B99
               TabIndex        =   32
               Top             =   195
               Width           =   6420
            End
         End
         Begin VB.Frame frEzeBPorc1 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   25
            Top             =   60
            Width           =   6810
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   1
               Left            =   60
               OleObjectBlob   =   "frmLGC.frx":1110A
               TabIndex        =   36
               Top             =   165
               Width           =   6645
            End
         End
         Begin VB.Frame frEzeB0 
            Height          =   2925
            Left            =   180
            TabIndex        =   17
            Top             =   30
            Width           =   7035
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2670
               Index           =   0
               Left            =   180
               OleObjectBlob   =   "frmLGC.frx":1150E
               TabIndex        =   29
               Top             =   180
               Width           =   6645
            End
         End
         Begin VB.Frame frEzeBPorc2 
            Height          =   2940
            Left            =   -74925
            TabIndex        =   42
            Top             =   45
            Width           =   7110
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmLGC.frx":11AD3
               TabIndex        =   43
               Top             =   195
               Width           =   6840
            End
         End
         Begin VB.Frame frEzeBPorc0 
            Height          =   2910
            Left            =   165
            TabIndex        =   24
            Top             =   45
            Width           =   7035
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   0
               Left            =   195
               OleObjectBlob   =   "frmLGC.frx":11ED7
               TabIndex        =   35
               Top             =   150
               Width           =   6675
            End
         End
      End
      Begin TabDlg.SSTab tabAep 
         Height          =   3330
         Left            =   -74940
         TabIndex        =   15
         Top             =   315
         Width           =   7365
         _ExtentX        =   12991
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
         TabPicture(0)   =   "frmLGC.frx":122DC
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frAepPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frAep0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets"
         TabPicture(1)   =   "frmLGC.frx":122F8
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frAepPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frAep1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmLGC.frx":12314
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frAepPorc2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "frAep2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick"
         TabPicture(3)   =   "frmLGC.frx":12330
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprAep(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom Imp x Pax"
         TabPicture(4)   =   "frmLGC.frx":1234C
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprAep(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprAep 
            Height          =   2700
            Index           =   3
            Left            =   -74820
            OleObjectBlob   =   "frmLGC.frx":12368
            TabIndex        =   58
            Top             =   225
            Width           =   6600
         End
         Begin FPSpread.vaSpread sprAep 
            Height          =   2700
            Index           =   4
            Left            =   -74835
            OleObjectBlob   =   "frmLGC.frx":1281D
            TabIndex        =   59
            Top             =   225
            Width           =   6600
         End
         Begin VB.Frame frAep0 
            Height          =   3000
            Left            =   255
            TabIndex        =   18
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmLGC.frx":12CD2
               TabIndex        =   30
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frAepPorc0 
            Height          =   2985
            Left            =   255
            TabIndex        =   26
            Top             =   45
            Width           =   6795
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   0
               Left            =   105
               OleObjectBlob   =   "frmLGC.frx":131C0
               TabIndex        =   38
               Top             =   195
               Width           =   6195
            End
         End
         Begin VB.Frame frAep1 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   21
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmLGC.frx":13577
               TabIndex        =   37
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frAepPorc1 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   27
            Top             =   30
            Width           =   6405
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   1
               Left            =   105
               OleObjectBlob   =   "frmLGC.frx":13A57
               TabIndex        =   39
               Top             =   210
               Width           =   6180
            End
         End
         Begin VB.Frame frAep2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   45
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmLGC.frx":13E05
               TabIndex        =   48
               Top             =   180
               Width           =   6600
            End
         End
         Begin VB.Frame frAepPorc2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   46
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmLGC.frx":142E6
               TabIndex        =   47
               Top             =   210
               Width           =   6600
            End
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1500
      Left            =   7995
      TabIndex        =   11
      Top             =   -15
      Width           =   990
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
         Left            =   195
         Picture         =   "frmLGC.frx":14685
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   150
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
         Left            =   195
         Picture         =   "frmLGC.frx":14787
         Style           =   1  'Graphical
         TabIndex        =   79
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
         Left            =   195
         Picture         =   "frmLGC.frx":14889
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   990
         Width           =   570
      End
   End
   Begin VB.Frame frdatos 
      Height          =   1530
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   7860
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2835
         Picture         =   "frmLGC.frx":150AB
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   195
         Width           =   375
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   2835
         Picture         =   "frmLGC.frx":1521D
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   555
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Comitente"
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
         Height          =   1335
         Index           =   1
         Left            =   3345
         TabIndex        =   3
         Top             =   120
         Width           =   4305
         Begin VB.ComboBox cboComi 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1485
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   915
            Width           =   2475
         End
         Begin VB.TextBox txtComi 
            Height          =   285
            Left            =   2595
            TabIndex        =   9
            Top             =   855
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.ListBox lstComi 
            Height          =   255
            Left            =   3105
            TabIndex        =   8
            Top             =   870
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.OptionButton optComi 
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
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   0
            Left            =   345
            TabIndex        =   7
            Top             =   570
            Width           =   1080
         End
         Begin VB.OptionButton optComi 
            Caption         =   "IOSC y Propios"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   1
            Left            =   1815
            TabIndex        =   6
            Top             =   555
            Width           =   1500
         End
         Begin VB.OptionButton optComi 
            Caption         =   "Otros"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   2
            Left            =   345
            TabIndex        =   5
            Top             =   945
            Width           =   1005
         End
         Begin VB.OptionButton optComi 
            Caption         =   "General"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   3
            Left            =   345
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1005
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Grupos"
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
         Height          =   615
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   870
         Width           =   2940
         Begin VB.ComboBox cboGrupo 
            Height          =   315
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   225
            Width           =   1845
         End
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1560
         TabIndex        =   65
         Top             =   210
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
         Left            =   1575
         TabIndex        =   66
         Top             =   555
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
         Left            =   315
         TabIndex        =   68
         Top             =   210
         Width           =   1185
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
         Left            =   315
         TabIndex        =   67
         Top             =   555
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmLGC"
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

Dim UsoPorInte As Boolean

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

If cboGrupo.Text <> "TODOS" Then
    Cond = Cond & " and grupo_venta = '" & cboGrupo.Text & "'"
End If

If optComi(1).Value Then
    Cond = Cond & " and  Comitente IN ('IOSC','IBAIR') "
ElseIf optComi(0).Value Then
    Cond = Cond & " and  Comitente NOT IN ('IOSC','IBAIR','T') "
ElseIf optComi(2).Value Then
    Cond = Cond & " and  Comitente = '" & lstComi.List(cboComi.ListIndex) & "'"
Else
    Cond = Cond & " and  Comitente = 'T' "
End If


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

If cboGrupo.Text <> "TODOS" Then
    Cond = Cond & " and grupo_venta = '" & cboGrupo.Text & "'"
End If

L_ArmarcondicionPax = Cond

End Function


Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim fila As Integer

fila = 0
Do While Not RsData.EOF
    fecha = Format$(RsData!Fch_Ticket, FTOFECHA)
    fila = fila + 1
    Do While fecha = Format$(RsData!Fch_Ticket, FTOFECHA)

        Select Case RsData!cod_depn
            Case DSLoc(1).Dep
                L_LlenarGrillaAEP fecha, fila
            Case DSLoc(2).Dep
                Select Case RsData!Cod_Sdep
                    Case DSLoc(2).Sdep
                        L_LlenarGrillaINTA fecha, fila
                    Case DSLoc(3).Sdep
                        L_LlenarGrillaINTB fecha, fila
                End Select
         'Interior
            Case DSLoc(4).Dep
                L_LlenarGrillaINTERIOR fecha, fila
        End Select
        If RsData.EOF Then
            Exit Do
        End If
    Loop
Loop

For i = 0 To 1
    Spread_TotalesGrillas sprEzeA(i), sprEzeA(i).MaxCols - 1, 2
    Spread_TotalesGrillas sprEzeB(i), sprEzeB(i).MaxCols - 1, 2
    Spread_TotalesGrillas sprAEP(i), sprAEP(i).MaxCols - 1, 2
'Interior
    Spread_TotalesGrillas sprInt(i), sprInt(i).MaxCols - 1, 2
    
Next

FuncLocal_PromediosFilaSPR sprEzeA(0), sprEzeA(1), sprEzeA(3), sprEzeA(0).MaxCols
FuncLocal_PromediosFilaSPR sprEzeA(0), sprEzeA(2), sprEzeA(4), sprEzeA(0).MaxCols
FuncLocal_PromediosFilaSPR sprAEP(0), sprAEP(1), sprAEP(3), sprAEP(0).MaxCols
FuncLocal_PromediosFilaSPR sprAEP(0), sprAEP(2), sprAEP(4), sprAEP(0).MaxCols
FuncLocal_PromediosFilaSPR sprEzeB(0), sprEzeB(1), sprEzeB(3), sprEzeB(0).MaxCols
FuncLocal_PromediosFilaSPR sprEzeB(0), sprEzeB(2), sprEzeB(4), sprEzeB(0).MaxCols

'Interior
FuncLocal_PromediosFilaSPR sprInt(0), sprInt(1), sprInt(3), sprInt(0).MaxCols
FuncLocal_PromediosFilaSPR sprInt(0), sprInt(2), sprInt(4), sprInt(0).MaxCols

L_RefrescarFrames

End Sub
Private Sub L_DecoEspigonPax()
'Los pasajeros tienen un tratamiento especial
'Es otro recordset por eso se hace un seguimiento aparte
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim fila As Integer

fila = 0
Do While Not RsDataPax.EOF
    fecha = Format$(RsDataPax!fch_vta, FTOFECHA)
    fila = fila + 1
    Do While fecha = Format$(RsDataPax!fch_vta, FTOFECHA)

        Select Case RsDataPax!cod_depn
            Case DSLoc(1).Dep
                L_LlenarGrillaAEPPx fecha, fila
            Case DSLoc(2).Dep
                Select Case RsDataPax!Cod_Sdep
                    Case DSLoc(2).Sdep
                        L_LlenarGrillaINTAPax fecha, fila
                    Case DSLoc(3).Sdep
                        L_LlenarGrillaINTBPax fecha, fila
                End Select
         'Interior
            Case DSLoc(4).Dep
                L_LlenarGrillaINTERIORPax fecha, fila
        End Select
        If RsDataPax.EOF Then
            Exit Do
        End If

    Loop
Loop

    Spread_TotalesGrillas sprEzeA(2), sprEzeA(2).MaxCols - 1, 2
    Spread_TotalesGrillas sprEzeB(2), sprEzeB(2).MaxCols - 1, 2
    Spread_TotalesGrillas sprAEP(2), sprAEP(2).MaxCols - 1, 2
    Spread_TotalesGrillas sprInt(2), sprInt(2).MaxCols - 1, 2


L_RefrescarFrames

End Sub

Private Sub L_LimpiarGrillas()
Dim i
    
For i = 0 To 4
    sprEzeA(i).MaxRows = 0
    sprEzeB(i).MaxRows = 0
    sprAEP(i).MaxRows = 0
    sprInt(i).MaxRows = 0
Next
    
End Sub

Private Sub L_PintarGrillas()
Dim i

For i = 0 To 4
    Spread_PintarfinSemana sprEzeA(i)
    Spread_PintarfinSemana sprEzeB(i)
    Spread_PintarfinSemana sprAEP(i)
    Spread_PintarfinSemana sprInt(i)
Next

End Sub

Private Sub L_Porcentages(ByRef sprPorc As control, sprDato As control, Orientacion As String)
Dim i As Integer, j As Integer
Dim valor As Variant
Dim tot As Variant, fecha As Variant
Dim Result As Double

On Error GoTo ErrPorc:

sprPorc.MaxRows = 0

Select Case Orientacion
    Case "FILAS"
        For i = 1 To sprDato.MaxRows
            sprDato.GetText sprDato.MaxCols, i, tot
            sprDato.GetText 1, i, fecha
            Result = 0
            sprPorc.MaxRows = i
            
            sprPorc.SetText 1, i, Format$(fecha, "dd-mm-yy")
            sprPorc.SetText TotalINTA, i, "100"
            
            For j = 2 To sprDato.MaxCols - 1
                sprDato.GetText j, i, valor
                If tot <> 0 Then
                    Result = valor * 100 / tot
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
                sprDato.GetText j, sprDato.MaxRows, tot
                sprDato.GetText j, i, valor
                If tot <> 0 Then
                    Result = valor * 100 / tot
                Else
                    Result = 0
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
    'Para el INTERIOR
    FrIntPorc0.Refresh
    FrIntPorc1.Refresh
    FrINTPorc2.Refresh
    
End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset

'On Error GoTo ErrLGC:

frmLGC.caption = Aplicacion.SeteoProceso(frmLGC.caption)

sql = " SELECT "
sql = sql & " fch_ticket, "
sql = sql & " cod_depn, "
sql = sql & " cod_sdep, "
sql = sql & " decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local) cod_local, "
sql = sql & " decode(cod_local,'L05',0,'L06',0,'L21',0,'L07',0,cod_sloc) cod_sloc, "
sql = sql & " sum(cantidad) cant, "
sql = sql & " sum(importe) imp "
sql = sql & "FROM " & funcLocal_Vista("venta_lgi", Year(CDate(mskFDesde.FormattedText)))
sql = sql & L_Armarcondicion
sql = sql & "group by fch_ticket,cod_depn,cod_sdep,decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local),decode(cod_local,'L05',0,'L06',0,'L21',0,'L07',0,cod_sloc) "
sql = sql & " order by fch_ticket,cod_depn,cod_sdep,cod_local,cod_sloc"

sqlX = " SELECT "
sqlX = sqlX & " fch_vta, "
sqlX = sqlX & " cod_depn, "
sqlX = sqlX & " cod_sdep, "
sqlX = sqlX & " decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local) cod_local, "
sqlX = sqlX & " decode(cod_local,'L05',0,'L06',0,'L21',0,'L07',0,cod_sloc) cod_sloc, "
sqlX = sqlX & " sum(importe) imp, "
sqlX = sqlX & " sum(cant_pax) cant "
If cboGrupo.Text <> "TODOS" Then
    sqlX = sqlX & "FROM " & funcLocal_Vista("pax_grupo", Year(CDate(mskFDesde.FormattedText)))
Else
    sqlX = sqlX & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)))
End If
sqlX = sqlX & L_ArmarcondicionPax
sqlX = sqlX & "group by fch_vta,cod_depn,cod_sdep,decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local),decode(cod_local,'L05',0,'L06',0,'L21',0,'L07',0,cod_sloc) "
sqlX = sqlX & " order by fch_vta,cod_depn,cod_sdep,cod_local,cod_sloc"

If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        If optComi(3).Value Then
            Call Aplicacion.ObtenerRsDAO(sqlX, RsDataPax)
            L_DecoEspigonPax
            L_SeteoTabPax True
        Else
            L_SeteoTabPax False
        End If
        L_DecoEspigon
    End If
    L_PintarGrillas
    Aplicacion.CerrarDAO RsData
    
End If

ErrLGC:
    frmLGC.caption = Aplicacion.SeteoFin
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
        Case INTE
            UsoPorInte = True
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
            UsoPorInte = True
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorAero) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub


Private Sub L_SeteoTabPax(valor As Boolean)
Dim i

    For i = 1 To 4
        TabEzeA.TabEnabled(i) = valor
        TabEzeB.TabEnabled(i) = valor
        TabAep.TabEnabled(i) = valor
        TabINT.TabEnabled(i) = valor
    Next
    
End Sub

Private Sub L_TratarExcel(titulo As String, subTit As String, Esp As String, CantCol As Integer)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer, filaant
Dim i
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmLGC.caption = Aplicacion.SeteoProceso(frmLGC.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    AppExcel.Application.Visible = True
    
    ReDim titCol(CantCol)
    Col = 1
    fila = 3
    
    Exl_PonerValor AppExcel, 1, 1, titulo
    rango = Exl_rangos(1, 1, 1, CantCol)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, CantCol)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, "Grupo :" & cboGrupo.Text
    fila = fila + 2
    
    If optComi(1).Value Then
        Exl_PonerValor AppExcel, fila, Col, "Comitente : I.O.S.C y Propios"
    ElseIf optComi(0).Value Then
        Exl_PonerValor AppExcel, fila, Col, "Comitente : NO I.O.S.C "
    ElseIf optComi(3).Value Then
        Exl_PonerValor AppExcel, fila, Col, "Comitente : TODOS "
    Else
        Exl_PonerValor AppExcel, fila, Col, "Comitente : " & cboComi.Text
    End If
    Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
    
    fila = fila + 1
    Select Case Esp
        Case AERO
            For i = 1 To CantCol
                sprAEP(0).GetText i, 0, tit
                titCol(i) = tit
            Next

           For i = 0 To 4
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información sobre :" & L_NombreDato(i + 1)
               
               fila = fila + 2
               filaant = fila
               Exl_BajarGrillaExel sprAEP(i), AppExcel, fila, Col, titCol
               fila = fila + sprAEP(i).MaxRows
               rango = Exl_rangos(fila, fila, Col, sprAEP(i).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               If i > 2 Then
                   rango = Exl_rangos(filaant + 1, fila, Col + 1, sprAEP(i).MaxCols)
                   Exl_Format AppExcel, rango
               End If
'               AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
           Next
        Case EZEA
            For i = 1 To CantCol
                sprEzeA(0).GetText i, 0, tit
                titCol(i) = tit
            Next
       
           For i = 0 To 4
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información sobre :" & L_NombreDato(i + 1)
               
               fila = fila + 2
               filaant = fila
               Exl_BajarGrillaExel sprEzeA(i), AppExcel, fila, Col, titCol
               fila = fila + sprEzeA(i).MaxRows
               rango = Exl_rangos(fila, fila, Col, sprEzeA(i).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               If i > 2 Then
                   rango = Exl_rangos(filaant + 1, fila, Col + 1, sprEzeA(i).MaxCols)
                   Exl_Format AppExcel, rango
               End If
'               AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
           Next
        
        Case EZEB
            For i = 1 To CantCol
                sprEzeB(0).GetText i, 0, tit
                titCol(i) = tit
            Next
       
           For i = 0 To 4
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información sobre :" & L_NombreDato(i + 1)
               
               fila = fila + 2
               filaant = fila
               Exl_BajarGrillaExel sprEzeB(i), AppExcel, fila, Col, titCol
               fila = fila + sprEzeB(i).MaxRows
               rango = Exl_rangos(fila, fila, Col, sprEzeB(i).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               If i > 2 Then
                   rango = Exl_rangos(filaant + 1, fila, Col + 1, sprEzeB(i).MaxCols)
                   Exl_Format AppExcel, rango
               End If
'               rango = Exl_rangos(fila + 1, fila + 1, col, col)
'               AppExcel.application.cells(fila + 1, fila + 1).PageBreak = True
               'AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
           Next
    
        Case INTE
            For i = 1 To CantCol
                sprInt(0).GetText i, 0, tit
                titCol(i) = tit
            Next
       
           For i = 0 To 4
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información sobre :" & L_NombreDato(i + 1)
               
               fila = fila + 2
               filaant = fila
               Exl_BajarGrillaExel sprInt(i), AppExcel, fila, Col, titCol
               fila = fila + sprInt(i).MaxRows
               rango = Exl_rangos(fila, fila, Col, sprInt(i).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               If i > 2 Then
                   rango = Exl_rangos(filaant + 1, fila, Col + 1, sprInt(i).MaxCols)
                   Exl_Format AppExcel, rango
               End If
'               rango = Exl_rangos(fila + 1, fila + 1, col, col)
'               AppExcel.application.cells(fila + 1, fila + 1).PageBreak = True
               'AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
           Next
    
    
    End Select
    
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    'AppExcel.application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    
    AppExcel.SaveAs NOMBRE & ".xls"
'    AppExcel.Workbooks.Open nombre & ".xls"
    Set AppExcel.Application.ActiveSheet = Nothing
    Set AppExcel = Nothing
    
End If

ErrorExl:

    frmLGC.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub

Private Function L_NombreDato(i As Integer) As String
Select Case i
    Case 1
        L_NombreDato = "Importes"
    Case 2
        L_NombreDato = "Tickets"
    Case 3
        L_NombreDato = "Pasajeros"
    Case 4
        L_NombreDato = "Promedios por Ticket"
    Case 5
        L_NombreDato = "Promedios por Pasajeros"
End Select
End Function


Private Sub L_TratarImpresion(Esp As String)
Dim txtTit As String
Dim c_Lgi As CLlgi
Dim i
Dim valor As Variant

Set c_Lgi = New CLlgi
Set c_Lgi.col_prnLGI = New Collection

Select Case Esp
    Case EZEB
        txtTit = ""
        For i = 2 To sprEzeB(0).MaxCols
            sprEzeB(0).GetText i, 0, valor
            txtTit = txtTit & valor
        Next
        c_Lgi.depn = "EZE"
        c_Lgi.Sdep = "INTB"
        
End Select
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
    L_TratarExcel "AEROPARQUE ", "Informe por Local (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", AERO, sprAEP(0).MaxCols
End Sub

Private Sub botExcelInt_Click()
    L_TratarExcel "INTERIOR ", "Informe por Local (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", INTE, sprInt(0).MaxCols
End Sub


Private Sub botExelEzeB_Click()
    L_TratarExcel "ESPIGON INTERNACIONAL B", "Informe por Local (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", EZEB, sprEzeB(0).MaxCols
End Sub

Private Sub botGrafEvoAEP_Click()
    Select Case TabAep.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico AERO, "Importes", sprAEP(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico AERO, "Tickets", sprAEP(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico AERO, "Pasajeros", sprAEP(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select

End Sub

Private Sub botGrafEvoEZEA_Click()
    
    Select Case TabEzeA.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico EZEA, "Importes", sprEzeA(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico EZEA, "Tickets", sprEzeA(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico EZEA, "Pasajeros", sprEzeA(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select




End Sub

Private Sub botGrafEvoEZEB_Click()
    Select Case TabEzeB.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico EZEB, "Importes", sprEzeB(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico EZEB, "Tickets", sprEzeB(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico EZEB, "Pasajeros", sprEzeB(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select

End Sub


Private Sub botGrafEvoINT_Click()
    Select Case TabINT.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico INTE, "Importes", sprInt(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico INTE, "Tickets", sprInt(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico INTE, "Pasajeros", sprInt(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select
End Sub

Private Sub botGrTortaAEP_Click()
Dim col_datos As Collection

    Set col_datos = New Collection
    
    Select Case TabAep.Tab
        Case 0
            L_LLenarColeccion col_datos, sprAEP(0)
            FrmGraficos.CargarGrafico AERO, "Importes", col_datos
        Case 1
            L_LLenarColeccion col_datos, sprAEP(1)
            FrmGraficos.CargarGrafico AERO, "Tickets", col_datos
        Case 2
            L_LLenarColeccion col_datos, sprAEP(2)
            FrmGraficos.CargarGrafico AERO, "Pasajeros", col_datos
    End Select

End Sub

Private Sub botGrTortaEA_Click()
Dim col_datos As Collection

    Set col_datos = New Collection
    
    Select Case TabEzeA.Tab
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
    
    Select Case TabEzeB.Tab
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

Private Sub botGrTortaInt_Click()
Dim col_datos As Collection

    Set col_datos = New Collection
    
    Select Case TabINT.Tab
        Case 0
            L_LLenarColeccion col_datos, sprInt(0)
            FrmGraficos.CargarGrafico INTE, "Importes", col_datos
        Case 1
            L_LLenarColeccion col_datos, sprInt(1)
            FrmGraficos.CargarGrafico INTE, "Tickets", col_datos
        Case 2
            L_LLenarColeccion col_datos, sprInt(2)
            FrmGraficos.CargarGrafico INTE, "Pasajeros", col_datos
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


Private Sub botImpr_Click()

L_TratarImpresion EZEB

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
        botPorcAep(0).Enabled = False
        L_SeteoEjecutar False, AERO
        frAep0.Refresh
        frAep1.Refresh
        frAep2.Refresh
    Case 1
        frAep0.Visible = False
        frAep1.Visible = False
        frAep2.Visible = False
        L_Porcentages sprAepPorc(0), sprAEP(0), "FILAS"
        L_Porcentages sprAepPorc(1), sprAEP(1), "FILAS"
        L_Porcentages sprAepPorc(2), sprAEP(2), "FILAS"
        botPorcAep(0).Enabled = True
        frAepPorc0.Refresh
        frAepPorc1.Refresh
        frAepPorc2.Refresh
        L_SeteoEjecutar True, AERO
    Case 2
        frAep0.Visible = False
        frAep1.Visible = False
        frAep2.Visible = False
        L_Porcentages sprAepPorc(0), sprAEP(0), "COL"
        L_Porcentages sprAepPorc(1), sprAEP(1), "COL"
        L_Porcentages sprAepPorc(2), sprAEP(2), "COL"
        
        botPorcAep(0).Enabled = True
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





Private Sub botPorcInt_Click(Index As Integer)

Select Case Index
    Case 0
        FrInt0.Visible = True
        FrInt1.Visible = True
        FrInt2.Visible = True
        botPorcint(0).Enabled = False
        L_SeteoEjecutar False, INTE
        FrInt0.Refresh
        FrInt1.Refresh
        FrInt2.Refresh
    Case 1
        FrInt0.Visible = False
        FrInt1.Visible = False
        FrInt2.Visible = False
        L_Porcentages SprIntPorc(0), sprInt(0), "FILAS"
        L_Porcentages SprIntPorc(1), sprInt(1), "FILAS"
        L_Porcentages SprIntPorc(2), sprInt(2), "FILAS"
        botPorcint(0).Enabled = True
        FrIntPorc0.Refresh
        FrIntPorc1.Refresh
        FrINTPorc2.Refresh
        L_SeteoEjecutar True, INTE
    Case 2
        FrInt0.Visible = False
        FrInt1.Visible = False
        FrInt2.Visible = False
        L_Porcentages SprIntPorc(0), sprInt(0), "COL"
        L_Porcentages SprIntPorc(1), sprInt(1), "COL"
        L_Porcentages SprIntPorc(2), sprInt(2), "COL"
        
        botPorcint(0).Enabled = True
        FrIntPorc0.Refresh
        FrIntPorc1.Refresh
        FrINTPorc2.Refresh
        L_SeteoEjecutar True, INTE
        
End Select


End Sub



Private Sub ExcelIntA_Click()
    L_TratarExcel "ESPIGON INTERNACIONAL A", "Informe por Local (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", EZEA, sprEzeA(0).MaxCols
End Sub


Private Sub Form_Activate()
FuncLocal_SeteoTABS tabEspigon
tabEspigon.TabVisible(0) = False
End Sub

Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6000
Me.Width = 9300

cboGrupo.AddItem "TODOS"
cboGrupo.AddItem "A"
cboGrupo.AddItem "B"
cboGrupo.AddItem "C"

cboGrupo.ListIndex = 0

FuncCbo_LlenarCbosLST cboComi, lstComi, "COMITENTE"

mskFDesde.Text = func_Dia1SegunMes_Anio(Month(Date), Year(Date))
If CDate(mskFDesde.FormattedText) = Date Then
    mskFHasta.Text = Format$(Date, FTOFECHA)
Else
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
End If

'FuncLocal_SeteoTABS tabEspigon

UsoPorIntA = False
UsoPorIntB = False
UsoPorAero = False

L_RefrescarFrames
L_LimpiarGrillas
frmPrincipal.lstForms.AddItem "frmLGC"

End Sub




Private Sub Form_Unload(Cancel As Integer)
    FuncLocal_SacarForm "frmLGC"
End Sub


Private Sub mskFDesde_LostFocus()
    
    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
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


Private Sub optComi_Click(Index As Integer)

Select Case Index
    Case 0, 1, 3
        cboComi.ListIndex = -1
        cboComi.Enabled = False
    Case 2
        cboComi.ListIndex = 0
        cboComi.Enabled = True
    
End Select
End Sub




Private Sub L_LlenarGrillaINTA(fecha As String, fila As Integer)
Dim depn As String, Sdep As String
Dim i As Integer, Col As Integer
Dim totCant As Long
Dim TOTImp As Long
Dim encontro As String

depn = RsData!cod_depn
Sdep = RsData!Cod_Sdep

sprEzeA(0).MaxRows = sprEzeA(0).MaxRows + 1
sprEzeA(0).SetText 1, sprEzeA(0).MaxRows, Format$(fecha, "dd-mm-yy")

If optComi(3).Value Then
    sprEzeA(1).MaxRows = sprEzeA(1).MaxRows + 1
    sprEzeA(3).MaxRows = sprEzeA(3).MaxRows + 1

    sprEzeA(1).SetText 1, sprEzeA(1).MaxRows, Format$(fecha, "dd-mm-yy")
    sprEzeA(3).SetText 1, sprEzeA(3).MaxRows, Format$(fecha, "dd-mm-yy")
End If


totCant = 0
TOTImp = 0
        Do While fecha = Format$(RsData!Fch_Ticket, FTOFECHA) And _
            depn = RsData!cod_depn And Sdep = RsData!Cod_Sdep
            sprEzeA(0).SetText i + 1, sprEzeA(0).MaxRows, str(RsData!imp)
            encontro = "NO"
            For i = 1 To 9
                If DSLoc(2).locales(i) = RsData!cod_local & Trim(str(RsData!Cod_Sloc)) Then
                    sprEzeA(0).SetText i + 1, sprEzeA(0).MaxRows, str(RsData!imp)
                    If optComi(3).Value Then
                        sprEzeA(1).SetText i + 1, sprEzeA(1).MaxRows, str(RsData!cant)
                        If RsData!cant > 0 Then
                            sprEzeA(3).SetText i + 1, sprEzeA(3).MaxRows, str(RsData!imp / RsData!cant)
                        Else
                            sprEzeA(3).SetText i + 1, sprEzeA(3).MaxRows, "0"
                        End If
                    End If
                 
                    totCant = totCant + RsData!cant
                    TOTImp = TOTImp + RsData!imp
                    
                    encontro = "SI"
                    RsData.MoveNext
                Else
                
                End If
                    If encontro = "NO" Then
                        RsData.MoveNext
                    End If
                    
                    If RsData.EOF Then
                        Exit For
                    End If
            
            Next
            If optComi(3).Value Then
                sprEzeA(3).SetText sprEzeA(3).MaxCols, sprEzeA(3).MaxRows, Format$((TOTImp / totCant), "#.0")
            End If
            
            If RsData.EOF Then
                Exit Do
            End If
            
        Loop
End Sub

Private Sub L_LlenarGrillaINTERIOR(fecha As String, fila As Integer)
Dim depn As String, Sdep As String
Dim i As Integer
Dim totCant As Long
Dim TOTImp As Long


depn = RsData!cod_depn
Sdep = RsData!Cod_Sdep

sprInt(0).MaxRows = sprInt(0).MaxRows + 1
sprInt(0).SetText 1, sprInt(0).MaxRows, Format$(fecha, "dd-mm-yy")

If optComi(3).Value Then
    sprInt(1).MaxRows = sprInt(1).MaxRows + 1
    sprInt(3).MaxRows = sprInt(3).MaxRows + 1

    sprInt(1).SetText 1, sprInt(1).MaxRows, Format$(fecha, "dd-mm-yy")
    sprInt(3).SetText 1, sprInt(3).MaxRows, Format$(fecha, "dd-mm-yy")
End If


totCant = 0
TOTImp = 0

     Do While fecha = Format$(RsData!Fch_Ticket, FTOFECHA) And _
        depn = RsData!cod_depn
        
        Select Case RsData!Cod_Sdep
               
         Case "CORD"
            For i = 1 To 2
                If DSLoc(4).locales(i) = RsData!cod_local & Trim(str(RsData!Cod_Sloc)) Then
                    sprInt(0).SetText i + 1, sprInt(0).MaxRows, str(RsData!imp)
                    If optComi(3).Value Then
                        sprInt(1).SetText i + 1, sprInt(1).MaxRows, str(RsData!cant)
                        If RsData!cant > 0 Then
                            sprInt(3).SetText i + 1, sprInt(3).MaxRows, str(RsData!imp / RsData!cant)
                        Else
                            sprInt(3).SetText i + 1, sprInt(3).MaxRows, "0"
                        End If
                    End If
                    totCant = totCant + RsData!cant
                    TOTImp = TOTImp + RsData!imp
                    RsData.MoveNext
                End If
                    If RsData.EOF Then
                        Exit For
                    End If
            Next
            
            If RsData.EOF Then
                Exit Do
            End If

    Case "IGUA"
                If DSLoc(5).locales(1) = RsData!cod_local & Trim(str(RsData!Cod_Sloc)) Then
                    sprInt(0).SetText 6, sprInt(0).MaxRows, str(RsData!imp)
                    If optComi(3).Value Then
                        sprInt(1).SetText 6, sprInt(1).MaxRows, str(RsData!cant)
                        If RsData!cant > 0 Then
                            sprInt(3).SetText 6, sprInt(3).MaxRows, str(RsData!imp / RsData!cant)
                        Else
                            sprInt(3).SetText 6, sprInt(3).MaxRows, "0"
                        End If
                    End If
                    
                    totCant = totCant + RsData!cant
                    TOTImp = TOTImp + RsData!imp
                    
                    RsData.MoveNext
                End If
            
            If RsData.EOF Then
                Exit Do
            End If
            
        Case "BARI"

              If DSLoc(8).locales(1) = RsData!cod_local & Trim(str(RsData!Cod_Sloc)) Then
                    sprInt(0).SetText 8, sprInt(0).MaxRows, str(RsData!imp)
                    If optComi(3).Value Then
                        sprInt(1).SetText 8, sprInt(1).MaxRows, str(RsData!cant)
                        If RsData!cant > 0 Then
                            sprInt(3).SetText 8, sprInt(3).MaxRows, str(RsData!imp / RsData!cant)
                        Else
                            sprInt(3).SetText 8, sprInt(3).MaxRows, "0"
                        End If
                    End If
                    
                    totCant = totCant + RsData!cant
                    TOTImp = TOTImp + RsData!imp
                    
                    RsData.MoveNext
                End If
            
            If RsData.EOF Then
                Exit Do
            End If
            
       Case "MDPL"

             If DSLoc(7).locales(1) = RsData!cod_local & Trim(str(RsData!Cod_Sloc)) Then
                    sprInt(0).SetText 7, sprInt(0).MaxRows, str(RsData!imp)
                    If optComi(3).Value Then
                        sprInt(1).SetText 7, sprInt(1).MaxRows, str(RsData!cant)
                        If RsData!cant > 0 Then
                            sprInt(3).SetText 7, sprInt(3).MaxRows, str(RsData!imp / RsData!cant)
                        Else
                            sprInt(3).SetText 7, sprInt(3).MaxRows, "0"
                        End If
                    End If
                    
                    totCant = totCant + RsData!cant
                    TOTImp = TOTImp + RsData!imp
                    
                    RsData.MoveNext
                End If
            
            If RsData.EOF Then
                Exit Do
            End If
        
        
        Case "MEND"
            For i = 1 To 2
                If DSLoc(6).locales(i) = RsData!cod_local & Trim(str(RsData!Cod_Sloc)) Then
                    sprInt(0).SetText i + 3, sprInt(0).MaxRows, str(RsData!imp)
                    If optComi(3).Value Then
                        sprInt(1).SetText i + 3, sprInt(1).MaxRows, str(RsData!cant)
                        If RsData!cant > 0 Then
                            sprInt(3).SetText i + 3, sprInt(3).MaxRows, str(RsData!imp / RsData!cant)
                        Else
                            sprInt(3).SetText i + 3, sprInt(3).MaxRows, "0"
                        End If
                    End If
                    
                    totCant = totCant + RsData!cant
                    TOTImp = TOTImp + RsData!imp
                    
                    RsData.MoveNext
                End If
                    If RsData.EOF Then
                        Exit For
                    End If
            
            Next
            
            If RsData.EOF Then
                Exit Do
            End If
            
    
         End Select
               
            If RsData.EOF Then
                Exit Do
            End If
   
        Loop
        
        If optComi(3).Value Then
                sprInt(3).SetText sprInt(3).MaxCols, sprInt(3).MaxRows, Format$((TOTImp / totCant), "#.0")
        End If
            
            
End Sub

Private Sub L_LlenarGrillaINTAPax(fecha As String, fila As Integer)
Dim depn As String, Sdep As String
Dim i As Integer
Dim Total As Variant
Dim totCant As Long
Dim TOTImp As Long
Dim encontro As String

depn = RsDataPax!cod_depn
Sdep = RsDataPax!Cod_Sdep

sprEzeA(2).MaxRows = sprEzeA(2).MaxRows + 1
sprEzeA(4).MaxRows = sprEzeA(4).MaxRows + 1

sprEzeA(2).SetText 1, sprEzeA(2).MaxRows, Format$(fecha, "dd-mm-yy")
sprEzeA(4).SetText 1, sprEzeA(4).MaxRows, Format$(fecha, "dd-mm-yy")

totCant = 0
TOTImp = 0
     Do While fecha = Format$(RsDataPax!fch_vta, FTOFECHA) And _
        depn = RsDataPax!cod_depn And Sdep = RsDataPax!Cod_Sdep
        encontro = "NO"
            For i = 1 To 9
                If DSLoc(2).locales(i) = RsDataPax!cod_local & Trim(str(RsDataPax!Cod_Sloc)) Then
                    sprEzeA(2).SetText i + 1, sprEzeA(2).MaxRows, str(RsDataPax!cant)
                    If RsData!cant > 0 Then
                        sprEzeA(4).SetText i + 1, sprEzeA(4).MaxRows, str(RsDataPax!imp / RsDataPax!cant)
                    Else
                        sprEzeA(4).SetText i + 1, sprEzeA(4).MaxRows, "0"
                    End If
                    
                    totCant = totCant + RsDataPax!cant
                    TOTImp = TOTImp + RsDataPax!imp
                    
                    encontro = "SI"
                    RsDataPax.MoveNext
                Else
                    'RsDataPax.MoveNext
                End If
                    If encontro = "NO" Then
                        RsData.MoveNext
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


Private Sub L_LlenarGrillaINTB(fecha As String, fila As Integer)
Dim depn As String, Sdep As String
Dim i As Integer
Dim totCant As Long
Dim TOTImp As Long


depn = RsData!cod_depn
Sdep = RsData!Cod_Sdep

'If optComi(3).Value Then
    sprEzeB(0).MaxRows = sprEzeB(0).MaxRows + 1
    sprEzeB(0).SetText 1, sprEzeB(0).MaxRows, Format$(fecha, "dd-mm-yy")
'End If

If optComi(3).Value Then
    sprEzeB(1).MaxRows = sprEzeB(1).MaxRows + 1
    sprEzeB(3).MaxRows = sprEzeB(3).MaxRows + 1
    
    sprEzeB(1).SetText 1, sprEzeB(1).MaxRows, Format$(fecha, "dd-mm-yy")
    sprEzeB(3).SetText 1, sprEzeB(3).MaxRows, Format$(fecha, "dd-mm-yy")
End If

totCant = 0
TOTImp = 0

     Do While fecha = Format$(RsData!Fch_Ticket, FTOFECHA) And _
        depn = RsData!cod_depn And Sdep = RsData!Cod_Sdep
        
            For i = 1 To 5
                If DSLoc(3).locales(i) = RsData!cod_local & Trim(str(RsData!Cod_Sloc)) Then
                    sprEzeB(0).SetText i + 1, sprEzeB(0).MaxRows, str(RsData!imp)
                    If optComi(3).Value Then
                        sprEzeB(1).SetText i + 1, sprEzeB(1).MaxRows, str(RsData!cant)
                        If RsData!cant > 0 Then
                            sprEzeB(3).SetText i + 1, sprEzeB(3).MaxRows, str(RsData!imp / RsData!cant)
                        Else
                            sprEzeB(3).SetText i + 1, sprEzeB(3).MaxRows, "0"
                        End If
                    End If
                    
                    totCant = totCant + RsData!cant
                    TOTImp = TOTImp + RsData!imp
                    
                    RsData.MoveNext
                End If
                    If RsData.EOF Then
                        Exit For
                    End If
            
            Next
            If optComi(3).Value Then
                sprEzeB(3).SetText sprEzeB(3).MaxCols, sprEzeB(3).MaxRows, Format$((TOTImp / totCant), "#.0")
            End If
            
            If RsData.EOF Then
                Exit Do
            End If
        
        Loop
End Sub
Private Sub L_LlenarGrillaINTBPax(fecha As String, fila As Integer)
Dim depn As String, Sdep As String
Dim i As Integer
Dim Total As Variant
Dim totCant As Long
Dim TOTImp As Long

depn = RsDataPax!cod_depn
Sdep = RsDataPax!Cod_Sdep

sprEzeB(2).MaxRows = sprEzeB(2).MaxRows + 1
sprEzeB(4).MaxRows = sprEzeB(4).MaxRows + 1

sprEzeB(2).SetText 1, sprEzeB(2).MaxRows, Format$(fecha, "dd-mm-yy")
sprEzeB(4).SetText 1, sprEzeB(4).MaxRows, Format$(fecha, "dd-mm-yy")

totCant = 0
TOTImp = 0
     Do While fecha = Format$(RsDataPax!fch_vta, FTOFECHA) And _
        depn = RsDataPax!cod_depn And Sdep = RsDataPax!Cod_Sdep
        
            For i = 1 To 5
                If DSLoc(3).locales(i) = RsDataPax!cod_local & Trim(str(RsDataPax!Cod_Sloc)) Then
                    sprEzeB(2).SetText i + 1, sprEzeB(2).MaxRows, str(RsDataPax!cant)
                    If RsDataPax!cant > 0 Then
                       sprEzeB(4).SetText i + 1, sprEzeB(4).MaxRows, str(RsDataPax!imp / RsDataPax!cant)
                    Else
                        sprEzeB(4).SetText i + 1, sprEzeB(4).MaxRows, "0"
                    End If
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
Private Sub L_LlenarGrillaINTERIORPax(fecha As String, fila As Integer)
Dim depn As String, Sdep As String
Dim i As Integer
Dim Total As Variant
Dim totCant As Long
Dim TOTImp As Long

depn = RsDataPax!cod_depn
Sdep = RsDataPax!Cod_Sdep

sprInt(2).MaxRows = sprInt(2).MaxRows + 1
sprInt(4).MaxRows = sprInt(4).MaxRows + 1

sprInt(2).SetText 1, sprInt(2).MaxRows, Format$(fecha, "dd-mm-yy")
sprInt(4).SetText 1, sprInt(4).MaxRows, Format$(fecha, "dd-mm-yy")

totCant = 0
TOTImp = 0
     Do While fecha = Format$(RsDataPax!fch_vta, FTOFECHA) And _
        depn = RsDataPax!cod_depn
        
        Select Case RsDataPax!Cod_Sdep
               
         Case "CORD"
            For i = 1 To 2
                If DSLoc(4).locales(i) = RsDataPax!cod_local & Trim(str(RsDataPax!Cod_Sloc)) Then
                    sprInt(2).SetText i + 1, sprInt(2).MaxRows, str(RsDataPax!cant)
                    If optComi(3).Value Then
                        sprInt(4).SetText i + 1, sprInt(4).MaxRows, str(RsDataPax!cant)
                        If RsDataPax!cant > 0 Then
                            sprInt(4).SetText i + 1, sprInt(4).MaxRows, str(RsDataPax!imp / RsDataPax!cant)
                        Else
                            sprInt(4).SetText i + 1, sprInt(4).MaxRows, "0"
                        End If
                    End If
                 
                    totCant = totCant + RsDataPax!cant
                    TOTImp = TOTImp + RsDataPax!imp
                    
                    RsDataPax.MoveNext
                End If
                    If RsDataPax.EOF Then
                        Exit For
                    End If
            
            Next
            
            If RsDataPax.EOF Then
                Exit Do
            End If
            
        Case "IGUA"

                If DSLoc(5).locales(1) = RsDataPax!cod_local & Trim(str(RsDataPax!Cod_Sloc)) Then
                    sprInt(2).SetText 6, sprInt(2).MaxRows, str(RsDataPax!cant)
                    If optComi(3).Value Then
                        sprInt(4).SetText 6, sprInt(4).MaxRows, str(RsDataPax!cant)
                        If RsDataPax!cant > 0 Then
                            sprInt(4).SetText 6, sprInt(4).MaxRows, str(RsDataPax!imp / RsDataPax!cant)
                        Else
                            sprInt(4).SetText 6, sprInt(4).MaxRows, "0"
                        End If
                    End If
                    
                    totCant = totCant + RsDataPax!cant
                    TOTImp = TOTImp + RsDataPax!imp
                    
                    RsDataPax.MoveNext
                End If
            
            If RsDataPax.EOF Then
                Exit Do
            End If
            
        Case "BARI"

              If DSLoc(8).locales(1) = RsDataPax!cod_local & Trim(str(RsDataPax!Cod_Sloc)) Then
                    sprInt(2).SetText 8, sprInt(2).MaxRows, str(RsDataPax!cant)
                    If optComi(3).Value Then
                        sprInt(4).SetText 8, sprInt(4).MaxRows, str(RsDataPax!cant)
                        If RsDataPax!cant > 0 Then
                            sprInt(4).SetText 8, sprInt(4).MaxRows, str(RsDataPax!imp / RsDataPax!cant)
                        Else
                            sprInt(4).SetText 8, sprInt(4).MaxRows, "0"
                        End If
                    End If
                    
                    totCant = totCant + RsDataPax!cant
                    TOTImp = TOTImp + RsDataPax!imp
                    
                    RsDataPax.MoveNext
                End If
            
            If RsDataPax.EOF Then
                Exit Do
            End If
            
       Case "MDPL"

             If DSLoc(7).locales(1) = RsDataPax!cod_local & Trim(str(RsDataPax!Cod_Sloc)) Then
                    sprInt(2).SetText 7, sprInt(2).MaxRows, str(RsDataPax!cant)
                    If optComi(3).Value Then
                        sprInt(4).SetText 7, sprInt(4).MaxRows, str(RsDataPax!cant)
                        If RsDataPax!cant > 0 Then
                            sprInt(4).SetText 7, sprInt(4).MaxRows, str(RsDataPax!imp / RsDataPax!cant)
                        Else
                            sprInt(4).SetText 7, sprInt(4).MaxRows, "0"
                        End If
                    End If
                    
                    totCant = totCant + RsDataPax!cant
                    TOTImp = TOTImp + RsDataPax!imp
                    
                    RsDataPax.MoveNext
                End If
            
 '           If optComi(3).Value Then
 '               sprInt(2).SetText sprInt(2).MaxCols, sprInt(2).MaxRows, Format$((TOTImp / totCant), "#.0")
 '           End If
            
            If RsDataPax.EOF Then
                Exit Do
            End If
        
        Case "MEND"
            For i = 1 To 2
                If DSLoc(6).locales(i) = RsDataPax!cod_local & Trim(str(RsDataPax!Cod_Sloc)) Then
                    sprInt(2).SetText i + 3, sprInt(2).MaxRows, str(RsDataPax!cant)
                    If optComi(3).Value Then
                        sprInt(4).SetText i + 3, sprInt(4).MaxRows, str(RsDataPax!cant)
                        If RsDataPax!cant > 0 Then
                            sprInt(4).SetText i + 3, sprInt(4).MaxRows, str(RsDataPax!imp / RsDataPax!cant)
                        Else
                            sprInt(4).SetText i + 3, sprInt(4).MaxRows, "0"
                        End If
                    End If
                 
                    totCant = totCant + RsDataPax!cant
                    TOTImp = TOTImp + RsDataPax!imp
                    
                    RsDataPax.MoveNext
                End If
                    If RsDataPax.EOF Then
                        Exit For
                    End If
            
            Next
            
            If RsDataPax.EOF Then
                Exit Do
            End If
            
      
         End Select
            
'Calcula los promedios por fila
            sprInt(4).SetText sprInt(4).MaxCols, sprInt(4).MaxRows, Format$((TOTImp / totCant), "#.0")
            
'            If RsDataPax.EOF Then
'                Exit Do
'            End If
        
        
        
        Loop
        




End Sub


Private Sub L_LlenarGrillaAEP(fecha As String, fila As Integer)
Dim depn As String, Sdep As String
Dim i As Integer
Dim TOTALImp As Variant
Dim TOTALCant As Variant

depn = RsData!cod_depn
Sdep = RsData!Cod_Sdep

sprAEP(0).MaxRows = sprAEP(0).MaxRows + 1
sprAEP(0).SetText 1, sprAEP(0).MaxRows, Format$(fecha, "dd-mm-yy")

If optComi(3).Value Then
    sprAEP(1).MaxRows = sprAEP(1).MaxRows + 1
    sprAEP(3).MaxRows = sprAEP(3).MaxRows + 1

    sprAEP(1).SetText 1, sprAEP(1).MaxRows, Format$(fecha, "dd-mm-yy")
    sprAEP(3).SetText 1, sprAEP(3).MaxRows, Format$(fecha, "dd-mm-yy")
End If

    TOTALImp = 0
    TOTALCant = 0
     Do While fecha = Format$(RsData!Fch_Ticket, FTOFECHA) And _
        depn = RsData!cod_depn And Sdep = RsData!Cod_Sdep
        
            For i = 1 To 3
                If DSLoc(1).locales(i) = RsData!cod_local & Trim(str(RsData!Cod_Sloc)) Then
                    sprAEP(0).SetText i + 1, sprAEP(0).MaxRows, str(RsData!imp)
                    If optComi(3).Value Then
                        sprAEP(1).SetText i + 1, sprAEP(1).MaxRows, str(RsData!cant)
                        If RsData!cant <> 0 Then ' Modifico Marcelo 12/02/99
                           sprAEP(3).SetText i + 1, sprAEP(3).MaxRows, str(RsData!imp / RsData!cant)
                        Else
                           sprAEP(3).SetText i + 1, sprAEP(3).MaxRows, "0"
                        End If
                    End If
                    TOTALImp = TOTALImp + RsData!imp
                    TOTALCant = TOTALCant + RsData!cant
                    RsData.MoveNext
                End If
                    If RsData.EOF Then
                        Exit For
                    End If
            
            Next
            If optComi(3).Value Then
                sprAEP(3).SetText sprAEP(3).MaxCols, sprAEP(3).MaxRows, Format$(TOTALImp / TOTALCant, "#.0")
            End If
            
            If RsData.EOF Then
                Exit Do
            End If
        
        Loop
End Sub


Private Sub L_LlenarGrillaAEPPx(fecha As String, fila As Integer)
Dim depn As String, Sdep As String
Dim i As Integer
Dim Total As Variant
Dim totCant As Long
Dim TOTImp As Long

depn = RsDataPax!cod_depn
Sdep = RsDataPax!Cod_Sdep

sprAEP(2).MaxRows = sprAEP(2).MaxRows + 1
sprAEP(4).MaxRows = sprAEP(4).MaxRows + 1

sprAEP(2).SetText 1, sprAEP(2).MaxRows, Format$(fecha, "dd-mm-yy")
sprAEP(4).SetText 1, sprAEP(4).MaxRows, Format$(fecha, "dd-mm-yy")

totCant = 0
TOTImp = 0
     Do While fecha = Format$(RsDataPax!fch_vta, FTOFECHA) And _
        depn = RsDataPax!cod_depn And Sdep = RsDataPax!Cod_Sdep
        
            For i = 1 To 3
                If DSLoc(1).locales(i) = RsDataPax!cod_local & Trim(str(RsDataPax!Cod_Sloc)) Then
                    sprAEP(2).SetText i + 1, sprAEP(2).MaxRows, str(RsDataPax!cant)
                    If RsDataPax!cant <> 0 Then ' Modifico Marcelo 12/02/99
                      sprAEP(4).SetText i + 1, sprAEP(4).MaxRows, (str(RsDataPax!imp / RsDataPax!cant))
                    Else
                      sprAEP(4).SetText i + 1, sprAEP(4).MaxRows, "0"
                    End If
                    
                    totCant = totCant + RsDataPax!cant
                    TOTImp = TOTImp + RsDataPax!imp
                    
                    RsDataPax.MoveNext
                
                End If
                    If RsDataPax.EOF Then
                        Exit For
                    End If
            
            Next

            sprAEP(4).SetText sprAEP(4).MaxCols, sprAEP(4).MaxRows, Format$((TOTImp / totCant), "#.0")
            
            If RsDataPax.EOF Then
                Exit Do
            End If
        
        Loop
        

End Sub



Private Sub tabAep_Click(PreviousTab As Integer)
frAep0.Refresh
frAep1.Refresh
frAep2.Refresh
Select Case TabAep.Tab
    Case 0, 1, 2
        botPorcAep(0).Visible = True
        botPorcAep(1).Visible = True
        botPorcAep(2).Visible = True
        botGrTortaAEP.Visible = True
        botGrafEvoAEP.Visible = True
    Case 3, 4
        botPorcAep(0).Visible = False
        botPorcAep(1).Visible = False
        botPorcAep(2).Visible = False
        botGrTortaAEP.Visible = False
        botGrafEvoAEP.Visible = False
End Select

End Sub

Private Sub tabEzeA_Click(PreviousTab As Integer)
frEzeA0.Refresh
frEzeA1.Refresh
frEzeA2.Refresh
Select Case TabEzeA.Tab
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
Select Case TabEzeB.Tab
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

Private Sub L_LLenarColeccion(ByRef Col As Collection, spr As control)
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
    cl_dato.Locale = valor
    
    spr.GetText i, spr.Row, valor
    cl_dato.DatoGral = valor
    
    Col.Add cl_dato
    
Next

End Sub

