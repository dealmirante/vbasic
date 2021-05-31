VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProd 
   Caption         =   "Seguimiento de unidades vendidas por Producto"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   9195
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
      Left            =   8040
      Picture         =   "frmProd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   885
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
      Height          =   420
      Index           =   1
      Left            =   8040
      Picture         =   "frmProd.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   465
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
      Height          =   420
      Index           =   0
      Left            =   8040
      Picture         =   "frmProd.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   30
      Width           =   570
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   4350
      Left            =   90
      TabIndex        =   1
      Top             =   1200
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   7673
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
      TabPicture(0)   =   "frmProd.frx":0A26
      Tab(0).ControlCount=   6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "stabRes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "froptR"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "botPorcR(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "botPorcR(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "botPorcR(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "botExcelRes"
      Tab(0).Control(5).Enabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmProd.frx":0A42
      Tab(1).ControlCount=   7
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "tabEzeA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "botGrTortaEA"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "botExcelEzeA"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "botPorc(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "botPorc(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "botPorc(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "froptA"
      Tab(1).Control(6).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmProd.frx":0A5E
      Tab(2).ControlCount=   7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabEzeB"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "botGrTortaEB"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "botExcelEzeB"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "botPorcB(0)"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "botPorcB(1)"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "botPorcB(2)"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "frOptB"
      Tab(2).Control(6).Enabled=   0   'False
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmProd.frx":0A7A
      Tab(3).ControlCount=   7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tabAep"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "botGrTortaAEP"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "botExcelAep"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "botPorcAep(0)"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "botPorcAep(1)"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "botPorcAep(2)"
      Tab(3).Control(5).Enabled=   -1  'True
      Tab(3).Control(6)=   "FrOptAEP"
      Tab(3).Control(6).Enabled=   0   'False
      TabCaption(4)   =   "INTERIOR"
      TabPicture(4)   =   "frmProd.frx":0A96
      Tab(4).ControlCount=   7
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TabInte"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "botPorcint(2)"
      Tab(4).Control(1).Enabled=   -1  'True
      Tab(4).Control(2)=   "botPorcint(1)"
      Tab(4).Control(2).Enabled=   -1  'True
      Tab(4).Control(3)=   "botPorcint(0)"
      Tab(4).Control(3).Enabled=   -1  'True
      Tab(4).Control(4)=   "botExcelInt"
      Tab(4).Control(4).Enabled=   -1  'True
      Tab(4).Control(5)=   "botGrTortaInt"
      Tab(4).Control(5).Enabled=   -1  'True
      Tab(4).Control(6)=   "FrOptINT"
      Tab(4).Control(6).Enabled=   0   'False
      Begin VB.CommandButton botExcelRes 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67500
         Picture         =   "frmProd.frx":0AB2
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   3105
         Width           =   765
      End
      Begin VB.CommandButton botPorcR 
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
         Left            =   -67500
         Picture         =   "frmProd.frx":1044
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   1515
         Width           =   765
      End
      Begin VB.CommandButton botPorcR 
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
         Index           =   1
         Left            =   -67500
         Picture         =   "frmProd.frx":170A
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   975
         Width           =   765
      End
      Begin VB.CommandButton botPorcR 
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
         Left            =   -67500
         Picture         =   "frmProd.frx":1D60
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   450
         Width           =   765
      End
      Begin VB.Frame froptR 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   -74850
         TabIndex        =   99
         Top             =   315
         Width           =   5010
         Begin VB.OptionButton optRes 
            Caption         =   "Total"
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
            Height          =   195
            Index           =   3
            Left            =   3195
            TabIndex        =   103
            Top             =   135
            Value           =   -1  'True
            Width           =   750
         End
         Begin VB.OptionButton optRes 
            Caption         =   "INT"
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
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   102
            Top             =   135
            Width           =   690
         End
         Begin VB.OptionButton optRes 
            Caption         =   "AEP"
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
            Height          =   195
            Index           =   1
            Left            =   1155
            TabIndex        =   101
            Top             =   135
            Width           =   690
         End
         Begin VB.OptionButton optRes 
            Caption         =   "EZE"
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
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   100
            Top             =   135
            Width           =   690
         End
      End
      Begin TabDlg.SSTab stabRes 
         Height          =   3375
         Left            =   -74820
         TabIndex        =   94
         Top             =   780
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5953
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmProd.frx":26D2
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frResPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frRes0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmProd.frx":26EE
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frRes1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frResPorc1"
         Tab(1).Control(1).Enabled=   0   'False
         Begin VB.Frame frRes1 
            Height          =   2970
            Left            =   -74895
            TabIndex        =   97
            Top             =   30
            Width           =   6795
            Begin FPSpread.vaSpread sprRes 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":270A
               TabIndex        =   98
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frRes0 
            Height          =   2985
            Left            =   105
            TabIndex        =   95
            Top             =   30
            Width           =   6795
            Begin FPSpread.vaSpread sprRes 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":2BF0
               TabIndex        =   96
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frResPorc1 
            Height          =   2985
            Left            =   -74910
            TabIndex        =   106
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprResPorc 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":30D6
               TabIndex        =   107
               Top             =   180
               Width           =   6600
            End
         End
         Begin VB.Frame frResPorc0 
            Height          =   2985
            Left            =   90
            TabIndex        =   104
            Top             =   45
            Width           =   6780
            Begin FPSpread.vaSpread sprResPorc 
               Height          =   2700
               Index           =   0
               Left            =   120
               OleObjectBlob   =   "frmProd.frx":3494
               TabIndex        =   105
               Top             =   195
               Width           =   6600
            End
         End
      End
      Begin VB.Frame FrOptAEP 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -74790
         TabIndex        =   87
         Top             =   345
         Width           =   3270
         Begin VB.OptionButton optLocalAE 
            Caption         =   "Total"
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
            Height          =   195
            Index           =   2
            Left            =   2100
            TabIndex        =   90
            Top             =   150
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton optLocalAE 
            Caption         =   "L14"
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
            Height          =   195
            Index           =   1
            Left            =   1065
            TabIndex        =   89
            Top             =   150
            Width           =   690
         End
         Begin VB.OptionButton optLocalAE 
            Caption         =   "L09"
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
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   88
            Top             =   135
            Width           =   690
         End
      End
      Begin VB.Frame FrOptINT 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   -74775
         TabIndex        =   75
         Top             =   315
         Width           =   7005
         Begin VB.OptionButton OptlocalInt 
            Caption         =   "Total"
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
            Height          =   195
            Index           =   5
            Left            =   5100
            TabIndex        =   81
            Top             =   90
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton OptlocalInt 
            Caption         =   "L18"
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
            Height          =   195
            Index           =   4
            Left            =   4005
            TabIndex        =   80
            Top             =   90
            Width           =   630
         End
         Begin VB.OptionButton OptlocalInt 
            Caption         =   "L16"
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
            Height          =   195
            Index           =   3
            Left            =   2970
            TabIndex        =   79
            Top             =   90
            Width           =   870
         End
         Begin VB.OptionButton OptlocalInt 
            Caption         =   "L15"
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
            Height          =   195
            Index           =   2
            Left            =   2070
            TabIndex        =   78
            Top             =   75
            Width           =   600
         End
         Begin VB.OptionButton OptlocalInt 
            Caption         =   "L10"
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
            Height          =   195
            Index           =   1
            Left            =   1035
            TabIndex        =   77
            Top             =   75
            Width           =   645
         End
         Begin VB.OptionButton OptlocalInt 
            Caption         =   "L04"
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
            Height          =   255
            Index           =   0
            Left            =   195
            TabIndex        =   76
            Top             =   45
            Width           =   690
         End
         Begin VB.Label Label6 
            Caption         =   "BARI"
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
            Height          =   540
            Left            =   4260
            TabIndex        =   86
            Top             =   300
            Width           =   1500
         End
         Begin VB.Label Label5 
            Caption         =   "MEND"
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
            Height          =   600
            Left            =   3225
            TabIndex        =   85
            Top             =   300
            Width           =   1500
         End
         Begin VB.Label Label4 
            Caption         =   "MDPL"
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
            Height          =   600
            Left            =   2340
            TabIndex        =   84
            Top             =   285
            Width           =   1500
         End
         Begin VB.Label Label3 
            Caption         =   "MEND"
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
            Height          =   600
            Left            =   1290
            TabIndex        =   83
            Top             =   285
            Width           =   1500
         End
         Begin VB.Label Label2 
            Caption         =   "CORD"
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
            Height          =   600
            Left            =   450
            TabIndex        =   82
            Top             =   285
            Width           =   1500
         End
      End
      Begin VB.CommandButton botGrTortaInt 
         Height          =   500
         Left            =   -67125
         Picture         =   "frmProd.frx":3852
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   2055
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton botExcelInt 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67125
         Picture         =   "frmProd.frx":3C94
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   3135
         Width           =   765
      End
      Begin VB.CommandButton botPorcint 
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
         Left            =   -67125
         Picture         =   "frmProd.frx":4226
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   465
         Width           =   765
      End
      Begin VB.CommandButton botPorcint 
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
         Left            =   -67125
         Picture         =   "frmProd.frx":4B98
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   990
         Width           =   765
      End
      Begin VB.CommandButton botPorcint 
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
         Left            =   -67125
         Picture         =   "frmProd.frx":51EE
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1530
         Width           =   765
      End
      Begin TabDlg.SSTab TabInte 
         Height          =   3405
         Left            =   -74790
         TabIndex        =   61
         Top             =   825
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   6006
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmProd.frx":58B4
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FrIntPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "FrInt0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmProd.frx":58D0
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FrInt1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "FrIntPorc1"
         Tab(1).Control(1).Enabled=   0   'False
         Begin VB.Frame FrInt1 
            Height          =   3000
            Left            =   -74895
            TabIndex        =   68
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprInt 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":58EC
               TabIndex        =   69
               Top             =   180
               Width           =   6600
            End
         End
         Begin VB.Frame FrIntPorc1 
            Height          =   2985
            Left            =   -74895
            TabIndex        =   66
            Top             =   30
            Width           =   6795
            Begin FPSpread.vaSpread SprIntPorc 
               Height          =   2700
               Index           =   1
               Left            =   105
               OleObjectBlob   =   "frmProd.frx":5DC3
               TabIndex        =   67
               Top             =   195
               Width           =   6615
            End
         End
         Begin VB.Frame FrInt0 
            Height          =   3000
            Left            =   105
            TabIndex        =   64
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprInt 
               Height          =   2700
               Index           =   0
               Left            =   105
               OleObjectBlob   =   "frmProd.frx":6166
               TabIndex        =   65
               Top             =   180
               Width           =   6600
            End
         End
         Begin VB.Frame FrIntPorc0 
            Height          =   3000
            Left            =   120
            TabIndex        =   62
            Top             =   30
            Width           =   6765
            Begin FPSpread.vaSpread SprIntPorc 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":663D
               TabIndex        =   63
               Top             =   180
               Width           =   6555
            End
         End
      End
      Begin VB.Frame frOptB 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   -74805
         TabIndex        =   56
         Top             =   300
         Width           =   5010
         Begin VB.OptionButton optLocalB 
            Caption         =   "L01"
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
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   60
            Top             =   135
            Width           =   690
         End
         Begin VB.OptionButton optLocalB 
            Caption         =   "L02"
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
            Height          =   195
            Index           =   1
            Left            =   1155
            TabIndex        =   59
            Top             =   135
            Width           =   690
         End
         Begin VB.OptionButton optLocalB 
            Caption         =   "L03"
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
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   58
            Top             =   135
            Width           =   690
         End
         Begin VB.OptionButton optLocalB 
            Caption         =   "Total"
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
            Height          =   195
            Index           =   3
            Left            =   3195
            TabIndex        =   57
            Top             =   135
            Value           =   -1  'True
            Width           =   750
         End
      End
      Begin VB.Frame froptA 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   255
         TabIndex        =   54
         Top             =   405
         Width           =   6810
         Begin VB.OptionButton optLocalA 
            Caption         =   "L08"
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
            Height          =   210
            Index           =   3
            Left            =   3885
            TabIndex        =   113
            Top             =   60
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.OptionButton optLocalA 
            Caption         =   "L06"
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
            Height          =   210
            Index           =   2
            Left            =   1695
            TabIndex        =   112
            Top             =   60
            Width           =   765
         End
         Begin VB.OptionButton optLocalA 
            Caption         =   "L22"
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
            Height          =   210
            Index           =   1
            Left            =   885
            TabIndex        =   93
            Top             =   60
            Width           =   765
         End
         Begin VB.OptionButton optLocalA 
            Caption         =   "Total"
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
            Height          =   210
            Index           =   4
            Left            =   2625
            TabIndex        =   91
            Top             =   60
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton optLocalA 
            Caption         =   "L21"
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
            Height          =   210
            Index           =   0
            Left            =   30
            TabIndex        =   55
            Top             =   45
            Width           =   765
         End
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
         Left            =   -67500
         Picture         =   "frmProd.frx":69E0
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1530
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
         Left            =   -67500
         Picture         =   "frmProd.frx":70A6
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   990
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
         Left            =   -67500
         Picture         =   "frmProd.frx":76FC
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   465
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
         Left            =   -67515
         Picture         =   "frmProd.frx":806E
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1515
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
         Left            =   -67515
         Picture         =   "frmProd.frx":8734
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   975
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
         Left            =   -67515
         Picture         =   "frmProd.frx":8D8A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   450
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
         Left            =   7500
         Picture         =   "frmProd.frx":96FC
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1515
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
         Left            =   7500
         Picture         =   "frmProd.frx":9DC2
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   975
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
         Left            =   7500
         Picture         =   "frmProd.frx":A418
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   450
         Width           =   765
      End
      Begin VB.CommandButton botExcelAep 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67500
         Picture         =   "frmProd.frx":AD8A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3135
         Width           =   765
      End
      Begin VB.CommandButton botExcelEzeB 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67500
         Picture         =   "frmProd.frx":B31C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3150
         Width           =   765
      End
      Begin VB.CommandButton botExcelEzeA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   7500
         Picture         =   "frmProd.frx":B8AE
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3135
         Width           =   765
      End
      Begin VB.CommandButton botGrTortaAEP 
         Height          =   500
         Left            =   -67500
         Picture         =   "frmProd.frx":BE40
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2055
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton botGrTortaEB 
         Height          =   510
         Left            =   -67500
         Picture         =   "frmProd.frx":C282
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2055
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton botGrTortaEA 
         Height          =   500
         Left            =   7500
         Picture         =   "frmProd.frx":C6C4
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   765
      End
      Begin TabDlg.SSTab tabEzeA 
         Height          =   3330
         Left            =   195
         TabIndex        =   2
         Top             =   780
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   5874
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   4
         TabHeight       =   441
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmProd.frx":CB06
         Tab(0).ControlCount=   3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label7"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frEzeAPorc0"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "frEzeA0"
         Tab(0).Control(2).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmProd.frx":CB22
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeAPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frEzeA1"
         Tab(1).Control(1).Enabled=   0   'False
         Begin VB.Frame frEzeA1 
            Height          =   2970
            Left            =   -74910
            TabIndex        =   6
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   1
               Left            =   105
               OleObjectBlob   =   "frmProd.frx":CB3E
               TabIndex        =   18
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeAPorc1 
            Height          =   2985
            Left            =   -74910
            TabIndex        =   9
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeAPorc 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":D023
               TabIndex        =   21
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeA0 
            Height          =   2970
            Left            =   105
            TabIndex        =   4
            Top             =   30
            Width           =   6795
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":D3E0
               TabIndex        =   12
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeAPorc0 
            Height          =   2985
            Left            =   90
            TabIndex        =   8
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeAPorc 
               Height          =   2700
               Index           =   0
               Left            =   120
               OleObjectBlob   =   "frmProd.frx":D8C6
               TabIndex        =   20
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Vtas anteriores a la reforma"
            Height          =   195
            Left            =   4620
            TabIndex        =   92
            Top             =   3105
            Visible         =   0   'False
            Width           =   2370
         End
      End
      Begin TabDlg.SSTab tabAep 
         Height          =   3330
         Left            =   -74835
         TabIndex        =   3
         Top             =   810
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   5874
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   4
         TabHeight       =   441
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmProd.frx":DC83
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frAepPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frAep0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmProd.frx":DC9F
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frAep1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frAepPorc1"
         Tab(1).Control(1).Enabled=   0   'False
         Begin VB.Frame frAep0 
            Height          =   2760
            Left            =   90
            TabIndex        =   5
            Top             =   150
            Width           =   6780
            Begin FPSpread.vaSpread sprAep 
               Height          =   2505
               Index           =   0
               Left            =   105
               OleObjectBlob   =   "frmProd.frx":DCBB
               TabIndex        =   13
               Top             =   180
               Width           =   6600
            End
         End
         Begin VB.Frame frAepPorc0 
            Height          =   2745
            Left            =   60
            TabIndex        =   10
            Top             =   150
            Width           =   6795
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2490
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":E192
               TabIndex        =   14
               Top             =   195
               Width           =   6615
            End
         End
         Begin VB.Frame frAep1 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   7
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":E535
               TabIndex        =   19
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frAepPorc1 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   11
            Top             =   30
            Width           =   6810
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":EA0C
               TabIndex        =   22
               Top             =   195
               Width           =   6615
            End
         End
      End
      Begin TabDlg.SSTab tabEzeB 
         Height          =   3330
         Left            =   -74835
         TabIndex        =   44
         Top             =   795
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   5874
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   4
         TabHeight       =   441
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmProd.frx":EDAF
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frEzeBPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frEzeB0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmProd.frx":EDCB
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeBPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frEzeB1"
         Tab(1).Control(1).Enabled=   0   'False
         Begin VB.Frame frEzeB1 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   47
            Top             =   15
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":EDE7
               TabIndex        =   48
               Top             =   195
               Width           =   6645
            End
         End
         Begin VB.Frame frEzeBPorc1 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   49
            Top             =   30
            Width           =   6810
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":F2CC
               TabIndex        =   50
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeB0 
            Height          =   2985
            Left            =   60
            TabIndex        =   45
            Top             =   45
            Width           =   6795
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":F695
               TabIndex        =   46
               Top             =   195
               Width           =   6645
            End
         End
         Begin VB.Frame frEzeBPorc0 
            Height          =   3000
            Left            =   75
            TabIndex        =   51
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmProd.frx":FB7B
               TabIndex        =   52
               Top             =   195
               Width           =   6600
            End
         End
      End
   End
   Begin VB.Frame frdatos 
      Height          =   1170
      Left            =   15
      TabIndex        =   0
      Top             =   -45
      Width           =   7890
      Begin VB.CommandButton botProd 
         Caption         =   "Códigos de Productos Monitoreados"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4095
         TabIndex        =   53
         Top             =   300
         Width           =   3255
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   2700
         Picture         =   "frmProd.frx":FF38
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   690
         Width           =   375
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2715
         Picture         =   "frmProd.frx":100AA
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   210
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1545
         TabIndex        =   28
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFHasta 
         Height          =   285
         Left            =   1545
         TabIndex        =   29
         Top             =   705
         Width           =   1170
         _ExtentX        =   2064
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
         TabIndex        =   31
         Top             =   705
         Width           =   1140
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
         Left            =   345
         TabIndex        =   30
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsDataEzeA  As Recordset
Dim RsDataEzeB  As Recordset
Dim RsDataAep  As Recordset
Dim RsDataInt  As Recordset
Dim RsDataRes  As Recordset

Dim RsDataEzeAPax  As Recordset
Dim RsDataEzeBPax  As Recordset
Dim RsDataAepPax  As Recordset
Dim RsDataIntPax  As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas
                       'con porcentajes
Dim UsoPorIntB As Boolean
Dim UsoPorAero As Boolean
Dim UsoPorInterior  As Boolean

Dim UsoPorTotal As Boolean

Dim col_Prod As Collection

Private Function L_Armarcondicion(Esp As String) As String
Dim Cond As String, condLoc As String
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant, i

fechaDesde = mskFDesde.FormattedText
fechaHasta = mskFHasta.FormattedText

Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

Cond = Cond & " and Cod_prod IN ( " & L_TodosLosProd & " )"

Select Case Esp
    Case EZEA
        condLoc = L_LocalOpt(EZEA)
        Cond = Cond & " and Cod_depn = 'EZE' and cod_sdep = 'INTA' "
        If condLoc <> "" Then
            Cond = Cond & " and cod_local = '" & condLoc & "'"
        End If
    Case EZEB
        condLoc = L_LocalOpt(EZEB)
        Cond = Cond & " and Cod_depn = 'EZE' and cod_sdep = 'INTB' "
        If condLoc <> "" Then
            Cond = Cond & " and cod_local = '" & condLoc & "'"
        End If
    Case AERO
        condLoc = L_LocalOpt(AERO)
        Cond = Cond & " and Cod_depn = 'AEP' and cod_sdep = 'AEP' "
        If condLoc <> "" Then
            Cond = Cond & " and cod_local = '" & condLoc & "'"
        End If
    Case INTE
        condLoc = L_LocalOpt(INTE)
        Cond = Cond & " and Cod_depn = 'INT' "
        ''and cod_sdep = 'AEP' "
        If condLoc <> "" Then
            Cond = Cond & " and cod_local = '" & condLoc & "'"
        End If
    Case "RESUMEN"
        condLoc = L_LocalOpt("RESUMEN")
        If condLoc <> "" Then
            Cond = Cond & " and Cod_depn = '" & condLoc & "'"
        End If
    
End Select

L_Armarcondicion = Cond

End Function

Private Sub L_LimpiarGrillas()
Dim i
For i = 0 To 1
    sprEzeA(i).MaxRows = 0
    sprEzeB(i).MaxRows = 0
    sprAEP(i).MaxRows = 0
    sprInt(i).MaxRows = 0
    sprRes(i).MaxRows = 0
Next


End Sub

Private Function L_LocalOpt(Esp As String) As String
Dim i
Dim Locale

Locale = ""

Select Case Esp
    Case EZEA
        For i = 0 To 3
            If optLocalA(i).Value Then
                Locale = optLocalA(i).caption
            End If
        Next
    Case EZEB
        For i = 0 To 2
            If optLocalB(i).Value Then
                Locale = optLocalB(i).caption
            End If
        Next
    Case AERO
        For i = 0 To 1
            If optLocalAE(i).Value Then
                Locale = optLocalAE(i).caption
            End If
        Next
    Case INTE
        For i = 0 To 4
            If OptlocalInt(i).Value Then
                Locale = OptlocalInt(i).caption
            End If
        Next
    Case "RESUMEN"
        For i = 0 To 2
            If optRes(i).Value Then
                Locale = optRes(i).caption
            End If
        Next
    
End Select
L_LocalOpt = Locale
End Function

Private Sub L_Porcentages(ByRef sprPorc As control, sprDato As control, Orientacion As String)
Dim i As Integer, j As Integer
Dim valor As Variant
Dim tot As Variant, prod As Variant
Dim Result As Double

On Error GoTo ErrPorc:

sprPorc.MaxRows = 0

Select Case Orientacion
    Case "FILAS"
        For i = 1 To sprDato.MaxRows
            sprDato.GetText sprDato.MaxCols, i, tot
            sprDato.GetText 1, i, prod
            Result = 0
            sprPorc.MaxRows = i
            
            sprPorc.SetText 1, i, str(prod)
            sprPorc.SetText sprPorc.MaxCols, i, "100"
            
            For j = 2 To sprDato.MaxCols - 1
                sprDato.GetText j, i, valor
                If tot <> "" Then
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
            sprDato.GetText 1, i, prod
            Result = 0
            sprPorc.MaxRows = i
        
            sprPorc.SetText 1, i, prod
            
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
        Spread_TotalesLinea sprPorc
ErrPorc:
    Exit Sub
    
End Sub

Private Sub L_RefrescarFrames()
    
    frEzeAPorc0.Refresh
    frEzeAPorc1.Refresh
    frEzeBPorc0.Refresh
    frEzeBPorc1.Refresh
    frAepPorc0.Refresh
    frAepPorc1.Refresh
    FrIntPorc0.Refresh
    FrIntPorc1.Refresh
    frEzeA0.Refresh
    frEzeA1.Refresh
    frEzeB0.Refresh
    frEzeB1.Refresh
    frAep0.Refresh
    frAep1.Refresh
    FrInt0.Refresh
    FrInt1.Refresh
    
    frRes0.Refresh
    frRes1.Refresh
    
    froptA.Refresh
    frOptB.Refresh
    FrOptINT.Refresh
    FrOptAEP.Refresh
    froptR.Refresh
    
End Sub

Private Sub L_RefrescarEzeA()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i
On Error GoTo ErrGLC:

frmProd.caption = Aplicacion.SeteoProceso(frmProd.caption)

sql = " SELECT "
sql = sql & " grupo_venta,"
sql = sql & " cod_depn, "
sql = sql & " cod_sdep, "
If Not optLocalA(4).Value Then
    sql = sql & " cod_local, "
End If
sql = sql & " cod_prod, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cantidad) cant "
sql = sql & "FROM " & funcLocal_Vista("venta_plg", Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion(EZEA)
If optLocalA(4).Value Then
    sql = sql & " group by grupo_venta,cod_depn,cod_sdep,cod_prod "
    sql = sql & " order by cod_prod,cod_depn,cod_sdep"
Else
    sql = sql & " group by grupo_venta,cod_depn,cod_sdep,cod_local,cod_prod "
    sql = sql & " order by cod_prod,cod_depn,cod_sdep,cod_local"
End If
If Aplicacion.ObtenerRsDAO(sql, RsDataEzeA) Then
    
    If Aplicacion.CantReg(RsDataEzeA) > 0 Then
        L_LlenarGrillaINTA
     Else
       'lIMPIO LA GRILLA PARA EL LOCAL QUE NO TUVO VENTAS
     For i = 0 To 1
       sprEzeB(i).MaxRows = 0
       sprEzeA(i).MaxRows = 0
     Next i
    End If

    Aplicacion.CerrarDAO RsDataEzeA
    
ErrGLC:
    frmProd.caption = Aplicacion.SeteoFin
    Exit Sub
    
End If

End Sub

Private Sub L_RefrescarResumen()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i
On Error GoTo ErrGLC:

frmProd.caption = Aplicacion.SeteoProceso(frmProd.caption)

sql = " SELECT "
sql = sql & " grupo_venta,"
If Not optRes(3).Value Then
    sql = sql & " cod_depn, "
End If
sql = sql & " cod_prod, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cantidad) cant "
sql = sql & "FROM " & funcLocal_Vista("venta_plg", Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion("RESUMEN")
If optRes(3).Value Then
    sql = sql & " group by grupo_venta,cod_prod "
    sql = sql & " order by cod_prod"
Else
    sql = sql & " group by grupo_venta,cod_depn,cod_prod "
    sql = sql & " order by cod_prod,cod_depn"
End If
If Aplicacion.ObtenerRsDAO(sql, RsDataRes) Then
    
    If Aplicacion.CantReg(RsDataRes) > 0 Then
        L_LlenarGrillaResumen
    Else
    End If

    Aplicacion.CerrarDAO RsDataRes
    
ErrGLC:
    frmProd.caption = Aplicacion.SeteoFin
    Exit Sub
    
End If

End Sub


Private Sub L_RefrescarEzeB()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrGLC:

frmProd.caption = Aplicacion.SeteoProceso(frmProd.caption)

sql = " SELECT "
sql = sql & " grupo_venta,"
sql = sql & " cod_depn, "
sql = sql & " cod_sdep, "
If Not optLocalB(3).Value Then
    sql = sql & " cod_local, "
End If
sql = sql & " cod_prod, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cantidad) cant "
sql = sql & "FROM " & funcLocal_Vista("venta_plg", Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion(EZEB)
If optLocalB(3).Value Then
    sql = sql & " group by grupo_venta,cod_depn,cod_sdep,cod_prod "
    sql = sql & " order by cod_prod,cod_depn,cod_sdep"
Else
    sql = sql & " group by grupo_venta,cod_depn,cod_sdep,cod_local,cod_prod "
    sql = sql & " order by cod_prod,cod_depn,cod_sdep,cod_local"
End If

If Aplicacion.ObtenerRsDAO(sql, RsDataEzeB) Then
    
    If Aplicacion.CantReg(RsDataEzeB) > 0 Then
        L_LlenarGrillaINTB
        'FuncLocal_PromediosFilaSPR sprEzeB(0), sprEzeB(1), sprEzeB(3)
        'FuncLocal_PromediosFilaSPR sprEzeB(0), sprEzeB(2), sprEzeB(4)
      Else
        'lIMPIO LA GRILLA PARA EL LOCAL QUE NO TUVO VENTAS
       For i = 0 To 1
        sprEzeB(i).MaxRows = 0
       Next i
    End If
    Aplicacion.CerrarDAO RsDataEzeB
    
End If

ErrGLC:
    frmProd.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub

Private Sub L_RefrescarAep()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrGLC:

frmProd.caption = Aplicacion.SeteoProceso(frmProd.caption)

sql = " SELECT "
sql = sql & " grupo_venta,"
sql = sql & " cod_depn, "
sql = sql & " cod_sdep, "
If Not optLocalAE(2).Value Then
    sql = sql & " cod_local, "
End If
sql = sql & " cod_prod, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cantidad) cant "
sql = sql & "FROM " & funcLocal_Vista("venta_plg", Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion(AERO)
If optLocalAE(2).Value Then
    sql = sql & " group by grupo_venta,cod_depn,cod_sdep,cod_prod "
    sql = sql & " order by cod_prod,cod_depn,cod_sdep"
Else
    sql = sql & " group by grupo_venta,cod_depn,cod_sdep,cod_local,cod_prod "
    sql = sql & " order by cod_prod,cod_depn,cod_sdep,cod_local"
End If

If Aplicacion.ObtenerRsDAO(sql, RsDataAep) Then
    
    If Aplicacion.CantReg(RsDataAep) > 0 Then
        L_LlenarGrillaAEP
        'FuncLocal_PromediosFilaSPR sprAep(0), sprAep(1), sprAep(3)
        'FuncLocal_PromediosFilaSPR sprAep(0), sprAep(2), sprAep(4)
      Else
        'lIMPIO LA GRILLA PARA EL LOCAL QUE NO TUVO VENTAS
        
        sprAEP(i).MaxRows = 0
    End If
        
    Aplicacion.CerrarDAO RsDataAep
    
End If

ErrGLC:
    frmProd.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub

Private Sub L_RefrescarInt()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrGLC:

frmProd.caption = Aplicacion.SeteoProceso(frmProd.caption)

sql = " SELECT "
sql = sql & " grupo_venta,"
sql = sql & " cod_depn, "
If Not OptlocalInt(5).Value Then
    sql = sql & " cod_sdep, "
    sql = sql & " cod_local, "
End If
sql = sql & " cod_prod, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cantidad) cant "
sql = sql & "FROM " & funcLocal_Vista("venta_plg", Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion(INTE)
If OptlocalInt(5).Value Then
    sql = sql & " group by grupo_venta,cod_depn,cod_prod "
    sql = sql & " order by cod_prod,cod_depn"
Else
    sql = sql & " group by grupo_venta,cod_depn,cod_sdep,cod_local,cod_prod "
    sql = sql & " order by cod_prod,cod_depn,cod_sdep,cod_local"
End If

If Aplicacion.ObtenerRsDAO(sql, RsDataInt) Then
    
    If Aplicacion.CantReg(RsDataInt) > 0 Then
        L_LlenarGrillaINT
      Else
        'lIMPIO LA GRILLA PARA EL LOCAL QUE NO TUVO VENTAS
      For i = 0 To 1
        sprInt(i).MaxRows = 0
      Next i

    End If
        
    Aplicacion.CerrarDAO RsDataInt
    
End If

ErrGLC:
    frmProd.caption = Aplicacion.SeteoFin
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
            UsoPorInterior = True
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
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorInterior Or UsoPorAero) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub
Private Sub L_SeteosGral()
        
    frdatos.Enabled = False
    botEjecutar(0).Enabled = False
    tabEspigon.Enabled = True


End Sub

Private Function L_TodosLosProd() As String
Dim cl As CLProdMon
Dim prod As String

prod = ""

For Each cl In col_Prod
    prod = prod & ", " & cl.Cod_prod
Next

If prod = "" Then
    L_TodosLosProd = "-1"
Else
    L_TodosLosProd = Right(prod, Len(prod) - 2)
End If


End Function

Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        Select Case Aplicacion.Perfil
            Case "TODOS", "COMUN"
                L_RefrescarEzeA
                L_RefrescarEzeB
                L_RefrescarAep
                L_RefrescarInt
                L_RefrescarResumen
            Case "INTA"
                L_RefrescarEzeA
            Case "INTA"
                L_RefrescarEzeB
            Case "AEP"
                L_RefrescarAep
            Case "INT"
                L_RefrescarInt
        End Select
        L_SeteosGral
        L_RefrescarFrames

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
Select Case TabAep.Tab
    Case 0
        L_TratarExcel sprAEP(0), sprAepPorc(0), "AEROPARQUE ", "Informe por Producto (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Importes", AERO
    Case 1
        L_TratarExcel sprAEP(1), sprAepPorc(1), "AEROPARQUE", "Informe por Producto (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Unidades", AERO
End Select

End Sub

Private Sub botExcelEzeA_Click()
Select Case TabEzeA.Tab
    Case 0
        L_Porcentages sprEzeAPorc(0), sprEzeA(0), "FILAS"
        L_TratarExcel sprEzeA(0), sprEzeAPorc(0), "ESPIGON INTERNACIONAL A", "Informe por Producto (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Importes", EZEA
    Case 1
        L_Porcentages sprEzeAPorc(1), sprEzeA(1), "FILAS"
        L_TratarExcel sprEzeA(1), sprEzeAPorc(0), "ESPIGON INTERNACIONAL A", "Informe por Producto (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Unidades", EZEA
End Select
End Sub

Private Sub botExcelEzeb_Click()
Select Case TabEzeB.Tab
    Case 0
        L_TratarExcel sprEzeB(0), sprEzeAPorc(0), "ESPIGON INTERNACIONAL B", "Informe por Producto (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Importes", EZEB
    Case 1
        L_TratarExcel sprEzeB(1), sprEzeAPorc(0), "ESPIGON INTERNACIONAL B", "Informe por Producto (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Unidades", EZEB
End Select

End Sub

Private Sub L_TratarExcel(spr As control, sprPorc As control, titulo As String, subTit As String, Info As String, Esp As String)
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

frmProd.caption = Aplicacion.SeteoProceso(frmProd.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
'    AppExcel.application.Visible = True
    
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
    rango = Exl_rangos(fila, fila + 3, 1, 1)
    
    Exl_PonerValor AppExcel, fila, Col, "Información sobre :" & Info
        
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, " Local/es :  " & IIf(L_LocalOpt(Esp) = "", "TOTAL", L_LocalOpt(Esp))
    
    Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
    
    fila = fila + 2
    
    Exl_BajarGrillaExel spr, AppExcel, fila, Col, titCol
    fila = fila + spr.MaxRows
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 2
    Exl_BajarGrillaExel sprPorc, AppExcel, fila, Col, titCol
    fila = fila + sprPorc.MaxRows
    rango = Exl_rangos(fila, fila, Col, sprPorc.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If

    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmProd.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub



Private Sub botExcelInt_Click()

Select Case TabAep.Tab
    Case 0
        L_TratarExcel sprInt(0), SprIntPorc(0), "INTERIOR", "Informe por Producto (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Importes", AERO
    Case 1
        L_TratarExcel sprInt(1), SprIntPorc(1), "INTERIOR", "Informe por Producto (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Unidades", AERO
End Select

End Sub

Private Sub botExcelRes_Click()
Select Case stabRes.Tab
    Case 0
        L_Porcentages sprResPorc(0), sprRes(0), "FILAS"
        L_TratarExcel sprRes(0), sprResPorc(0), "Ventas por Espigón", "Informe por Producto (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Importes", "RESUMEN"
    Case 1
        L_Porcentages sprResPorc(1), sprRes(1), "FILAS"
        L_TratarExcel sprRes(1), sprResPorc(0), "Ventas por Espigón", "Informe por Producto (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Unidades", "RESUMEN"
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
            FrmGraficos.CargarGrafico AERO, "Unidades", col_datos

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
            FrmGraficos.CargarGrafico EZEA, "Unidades", col_datos

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
            FrmGraficos.CargarGrafico EZEB, "Unidades", col_datos

    End Select

End Sub

Private Sub botGrTortaInt_Click()

Dim col_datos As Collection

Set col_datos = New Collection
    
Select Case TabInte.Tab
       Case 0
            L_LLenarColeccion col_datos, sprInt(0)
            FrmGraficos.CargarGrafico INTE, "Importes", col_datos
        Case 1
            L_LLenarColeccion col_datos, sprInt(1)
            FrmGraficos.CargarGrafico INTE, "Unidades", col_datos
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
        botPorc(0).Enabled = False
        L_SeteoEjecutar False, EZEA
        frEzeA0.Refresh
        frEzeA1.Refresh

    Case 1
        frEzeA0.Visible = False
        frEzeA1.Visible = False

        L_Porcentages sprEzeAPorc(0), sprEzeA(0), "FILAS"
        L_Porcentages sprEzeAPorc(1), sprEzeA(1), "FILAS"
        botPorc(0).Enabled = True
        frEzeAPorc0.Refresh
        frEzeAPorc1.Refresh

        L_SeteoEjecutar True, EZEA
    Case 2
        frEzeA0.Visible = False
        frEzeA1.Visible = False

        L_Porcentages sprEzeAPorc(0), sprEzeA(0), "COL"
        L_Porcentages sprEzeAPorc(1), sprEzeA(1), "COL"
        
        botPorc(0).Enabled = True
        frEzeAPorc0.Refresh
        frEzeAPorc1.Refresh

        L_SeteoEjecutar True, EZEA
        
End Select

End Sub

Private Sub botPorcaep_Click(Index As Integer)

Select Case Index
    Case 0
        frAep0.Visible = True
        frAep1.Visible = True

        botPorcAep(0).Enabled = False
        L_SeteoEjecutar False, AERO
        frAep0.Refresh
        frAep1.Refresh

    Case 1
        frAep0.Visible = False
        frAep1.Visible = False

        L_Porcentages sprAepPorc(0), sprAEP(0), "FILAS"
        L_Porcentages sprAepPorc(1), sprAEP(1), "FILAS"

        botPorcAep(0).Enabled = True
        frAepPorc0.Refresh
        frAepPorc1.Refresh

        L_SeteoEjecutar True, AERO
    Case 2
        frAep0.Visible = False
        frAep1.Visible = False

        L_Porcentages sprAepPorc(0), sprAEP(0), "COL"
        L_Porcentages sprAepPorc(1), sprAEP(1), "COL"
        
        botPorcAep(0).Enabled = True
        frAepPorc0.Refresh
        frAepPorc1.Refresh

        L_SeteoEjecutar True, AERO
        
End Select


End Sub

Private Sub botPorcB_Click(Index As Integer)

Select Case Index
    Case 0
        frEzeB0.Visible = True
        frEzeB1.Visible = True

        botPorcB(0).Enabled = False
        L_SeteoEjecutar False, EZEB
        frEzeB0.Refresh
        frEzeB1.Refresh

    Case 1
        frEzeB0.Visible = False
        frEzeB1.Visible = False

        L_Porcentages sprEzeBPorc(0), sprEzeB(0), "FILAS"
        L_Porcentages sprEzeBPorc(1), sprEzeB(1), "FILAS"

        botPorcB(0).Enabled = True
        frEzeBPorc0.Refresh
        frEzeBPorc1.Refresh

        L_SeteoEjecutar True, EZEB
    Case 2
        frEzeB0.Visible = False
        frEzeB1.Visible = False

        L_Porcentages sprEzeBPorc(0), sprEzeB(0), "COL"
        L_Porcentages sprEzeBPorc(1), sprEzeB(1), "COL"
        
        botPorcB(0).Enabled = True
        frEzeBPorc0.Refresh
        frEzeBPorc1.Refresh

        L_SeteoEjecutar True, EZEB
        
End Select

End Sub



Private Sub botPorcInt_Click(Index As Integer)
Select Case Index
    Case 0
        FrInt0.Visible = True
        FrInt1.Visible = True

        botPorcint(0).Enabled = False
        L_SeteoEjecutar False, INTE
        FrInt0.Refresh
        FrInt1.Refresh

    Case 1
        FrInt0.Visible = False
        FrInt1.Visible = False

        L_Porcentages SprIntPorc(0), sprInt(0), "FILAS"
        L_Porcentages SprIntPorc(1), sprInt(1), "FILAS"

        botPorcint(0).Enabled = True
        FrIntPorc0.Refresh
        FrIntPorc1.Refresh

        L_SeteoEjecutar True, INTE
    Case 2
        FrInt0.Visible = False
        FrInt1.Visible = False

        L_Porcentages SprIntPorc(0), sprInt(0), "COL"
        L_Porcentages SprIntPorc(1), sprInt(1), "COL"
        
        botPorcint(0).Enabled = True
        FrIntPorc0.Refresh
        FrIntPorc1.Refresh

        L_SeteoEjecutar True, INTE
        
End Select

End Sub

Private Sub botPorcR_Click(Index As Integer)
Dim i

Select Case Index
    Case 0
        frRes0.Visible = True
        frRes1.Visible = True
        botPorcR(0).Enabled = False
        'L_SeteoEjecutar False, EZEA
        frRes0.Refresh
        frRes1.Refresh

    Case 1
        frRes0.Visible = False
        frRes1.Visible = False

        L_Porcentages sprResPorc(0), sprRes(0), "FILAS"
        L_Porcentages sprResPorc(1), sprRes(1), "FILAS"
        botPorcR(0).Enabled = True
        frResPorc0.Refresh
        frResPorc1.Refresh

        'L_SeteoEjecutar True, EZEA
    Case 2
        frRes0.Visible = False
        frRes1.Visible = False

        L_Porcentages sprResPorc(0), sprRes(0), "COL"
        L_Porcentages sprResPorc(1), sprRes(1), "COL"
        
        botPorcR(0).Enabled = True
        frResPorc0.Refresh
        frResPorc1.Refresh

        'L_SeteoEjecutar True, EZEA
        
End Select


End Sub

Private Sub botProd_Click()
    frmProdMon.ProductosMonitoreo col_Prod
End Sub

Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabEspigon
End Sub

Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6000
Me.Width = 9300


mskFDesde.Text = func_Dia1SegunMes_Anio(Month(Date), Year(Date))
If CDate(mskFDesde.FormattedText) = Date Then
    mskFHasta.Text = Format$(Date, FTOFECHA)
Else
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
End If

UsoPorIntA = False
UsoPorIntB = False
UsoPorTotal = False
UsoPorAero = False

L_RefrescarFrames

Set col_Prod = New Collection
L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmProd"

End Sub




Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmProd"
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


Private Sub L_LlenarGrillaINTA()
Dim codProd As String
Dim i As Integer

For i = 0 To 1
    sprEzeA(i).MaxRows = 0
Next

Do While Not RsDataEzeA.EOF
    
    sprEzeA(0).MaxRows = sprEzeA(0).MaxRows + 1
    sprEzeA(1).MaxRows = sprEzeA(1).MaxRows + 1
    
    codProd = RsDataEzeA!Cod_prod
    sprEzeA(0).SetText 1, sprEzeA(0).MaxRows, codProd
    sprEzeA(1).SetText 1, sprEzeA(1).MaxRows, codProd

    Do While codProd = RsDataEzeA!Cod_prod
        Select Case RsDataEzeA.Grupo_Venta
            Case "A"
                sprEzeA(0).SetText 2, sprEzeA(0).MaxRows, str(RsDataEzeA!imp)
                sprEzeA(1).SetText 2, sprEzeA(1).MaxRows, str(RsDataEzeA!cant)
            Case "B"
                sprEzeA(0).SetText 3, sprEzeA(0).MaxRows, str(RsDataEzeA!imp)
                sprEzeA(1).SetText 3, sprEzeA(1).MaxRows, str(RsDataEzeA!cant)
            Case "C"
                sprEzeA(0).SetText 4, sprEzeA(0).MaxRows, str(RsDataEzeA!imp)
                sprEzeA(1).SetText 4, sprEzeA(1).MaxRows, str(RsDataEzeA!cant)
        End Select
        
        RsDataEzeA.MoveNext
        If RsDataEzeA.EOF Then
            Exit Do
        End If
    Loop

Loop
For i = 0 To 1
    Spread_TotalesGrillas sprEzeA(i), sprEzeA(i).MaxCols - 1, 2
Next

End Sub

Private Sub L_LlenarGrillaINTB()
Dim codProd As String
Dim i As Integer

For i = 0 To 1
    sprEzeB(i).MaxRows = 0
Next

Do While Not RsDataEzeB.EOF
    
    sprEzeB(0).MaxRows = sprEzeB(0).MaxRows + 1
    sprEzeB(1).MaxRows = sprEzeB(1).MaxRows + 1
    
    codProd = RsDataEzeB!Cod_prod
    sprEzeB(0).SetText 1, sprEzeB(0).MaxRows, codProd
    sprEzeB(1).SetText 1, sprEzeB(1).MaxRows, codProd

    Do While codProd = RsDataEzeB!Cod_prod
        Select Case RsDataEzeB.Grupo_Venta
            Case "A"
                sprEzeB(0).SetText 2, sprEzeB(0).MaxRows, str(RsDataEzeB!imp)
                sprEzeB(1).SetText 2, sprEzeB(1).MaxRows, str(RsDataEzeB!cant)
            Case "B"
                sprEzeB(0).SetText 3, sprEzeB(0).MaxRows, str(RsDataEzeB!imp)
                sprEzeB(1).SetText 3, sprEzeB(1).MaxRows, str(RsDataEzeB!cant)
            Case "C"
                sprEzeB(0).SetText 4, sprEzeB(0).MaxRows, str(RsDataEzeB!imp)
                sprEzeB(1).SetText 4, sprEzeB(1).MaxRows, str(RsDataEzeB!cant)
        End Select
        
        RsDataEzeB.MoveNext
        If RsDataEzeB.EOF Then
            Exit Do
        End If
    Loop

Loop
For i = 0 To 1
    Spread_TotalesGrillas sprEzeB(i), sprEzeB(i).MaxCols - 1, 2
Next

End Sub
Private Sub L_LlenarGrillaResumen()
Dim codProd As String
Dim i As Integer

For i = 0 To 1
    sprRes(i).MaxRows = 0
Next

Do While Not RsDataRes.EOF
    
    sprRes(0).MaxRows = sprRes(0).MaxRows + 1
    sprRes(1).MaxRows = sprRes(1).MaxRows + 1
    
    codProd = RsDataRes!Cod_prod
    sprRes(0).SetText 1, sprRes(0).MaxRows, codProd
    sprRes(1).SetText 1, sprRes(1).MaxRows, codProd

    Do While codProd = RsDataRes!Cod_prod
        Select Case RsDataRes.Grupo_Venta
            Case "A"
                sprRes(0).SetText 2, sprRes(0).MaxRows, str(RsDataRes!imp)
                sprRes(1).SetText 2, sprRes(1).MaxRows, str(RsDataRes!cant)
            Case "B"
                sprRes(0).SetText 3, sprRes(0).MaxRows, str(RsDataRes!imp)
                sprRes(1).SetText 3, sprRes(1).MaxRows, str(RsDataRes!cant)
            Case "C"
                sprRes(0).SetText 4, sprRes(0).MaxRows, str(RsDataRes!imp)
                sprRes(1).SetText 4, sprRes(1).MaxRows, str(RsDataRes!cant)
        End Select
        
        RsDataRes.MoveNext
        If RsDataRes.EOF Then
            Exit Do
        End If
    Loop

Loop
For i = 0 To 1
    Spread_TotalesGrillas sprRes(i), sprRes(i).MaxCols - 1, 2
Next

End Sub

Private Sub L_LlenarGrillaAEP()
Dim codProd As String
Dim i As Integer

For i = 0 To 1
    sprAEP(i).MaxRows = 0
Next

Do While Not RsDataAep.EOF
    
    sprAEP(0).MaxRows = sprAEP(0).MaxRows + 1
    sprAEP(1).MaxRows = sprAEP(1).MaxRows + 1
    
    codProd = RsDataAep!Cod_prod
    sprAEP(0).SetText 1, sprAEP(0).MaxRows, codProd
    sprAEP(1).SetText 1, sprAEP(1).MaxRows, codProd

    Do While codProd = RsDataAep!Cod_prod
        Select Case RsDataAep.Grupo_Venta
            Case "A"
                sprAEP(0).SetText 2, sprAEP(0).MaxRows, str(RsDataAep!imp)
                sprAEP(1).SetText 2, sprAEP(1).MaxRows, str(RsDataAep!cant)
            Case "B"
                sprAEP(0).SetText 3, sprAEP(0).MaxRows, str(RsDataAep!imp)
                sprAEP(1).SetText 3, sprAEP(1).MaxRows, str(RsDataAep!cant)
            Case "C"
                sprAEP(0).SetText 4, sprAEP(0).MaxRows, str(RsDataAep!imp)
                sprAEP(1).SetText 4, sprAEP(1).MaxRows, str(RsDataAep!cant)
        End Select
        
        RsDataAep.MoveNext
        If RsDataAep.EOF Then
            Exit Do
        End If
    Loop

Loop
For i = 0 To 1
    Spread_TotalesGrillas sprAEP(i), sprAEP(i).MaxCols - 1, 2
Next

End Sub
Private Sub L_LlenarGrillaINT()
Dim codProd As String
Dim i As Integer

For i = 0 To 1
    sprInt(i).MaxRows = 0
Next

Do While Not RsDataInt.EOF
    
    sprInt(0).MaxRows = sprInt(0).MaxRows + 1
    sprInt(1).MaxRows = sprInt(1).MaxRows + 1
    
    codProd = RsDataInt!Cod_prod
    sprInt(0).SetText 1, sprInt(0).MaxRows, codProd
    sprInt(1).SetText 1, sprInt(1).MaxRows, codProd

    Do While codProd = RsDataInt!Cod_prod
        Select Case RsDataInt!Grupo_Venta
            Case "A"
                sprInt(0).SetText 2, sprInt(0).MaxRows, str(RsDataInt!imp)
                sprInt(1).SetText 2, sprInt(1).MaxRows, str(RsDataInt!cant)
            Case "B"
                sprInt(0).SetText 3, sprInt(0).MaxRows, str(RsDataInt!imp)
                sprInt(1).SetText 3, sprInt(1).MaxRows, str(RsDataInt!cant)
            Case "C"
                sprInt(0).SetText 4, sprInt(0).MaxRows, str(RsDataInt!imp)
                sprInt(1).SetText 4, sprInt(1).MaxRows, str(RsDataInt!cant)
        End Select
        
        RsDataInt.MoveNext
        If RsDataInt.EOF Then
            Exit Do
        End If
    Loop

Loop
For i = 0 To 1
    Spread_TotalesGrillas sprInt(i), sprInt(i).MaxCols - 1, 2
Next

End Sub

Private Sub optLocalA_Click(Index As Integer)
    L_RefrescarEzeA
End Sub

Private Sub optLocalAE_Click(Index As Integer)
    L_RefrescarAep
End Sub

Private Sub optLocalB_Click(Index As Integer)
    L_RefrescarEzeB
End Sub

Private Sub Optlocalint_Click(Index As Integer)
    L_RefrescarInt
End Sub

Private Sub optRes_Click(Index As Integer)
    L_RefrescarResumen
End Sub

Private Sub tabAep_Click(PreviousTab As Integer)
frAep0.Refresh
frAep1.Refresh

Select Case TabAep.Tab
    Case 0, 1, 2
        botPorcAep(0).Visible = True
        botPorcAep(1).Visible = True
        botPorcAep(2).Visible = True
        botGrTortaAEP.Visible = True

    Case 3, 4
        botPorcAep(0).Visible = False
        botPorcAep(1).Visible = False
        botPorcAep(2).Visible = False
        botGrTortaAEP.Visible = False

End Select

End Sub

Private Sub tabEzeA_Click(PreviousTab As Integer)
frEzeA0.Refresh
frEzeA1.Refresh


Select Case TabEzeA.Tab
    Case 0, 1, 2
        botPorc(0).Visible = True
        botPorc(1).Visible = True
        botPorc(2).Visible = True
        botGrTortaEA.Visible = True

    Case 3, 4
        botPorc(0).Visible = False
        botPorc(1).Visible = False
        botPorc(2).Visible = False
        botGrTortaEA.Visible = False

End Select


End Sub

Private Sub tabEzeB_Click(PreviousTab As Integer)
frEzeB0.Refresh
frEzeB1.Refresh

Select Case TabEzeB.Tab
    Case 0, 1, 2
        botPorcB(0).Visible = True
        botPorcB(1).Visible = True
        botPorcB(2).Visible = True
        botGrTortaEB.Visible = True

    Case 3, 4
        botPorcB(0).Visible = False
        botPorcB(1).Visible = False
        botPorcB(2).Visible = False
        botGrTortaEB.Visible = False

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

