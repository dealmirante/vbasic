VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.0#0"; "crystl32.ocx"
Begin VB.Form frmVtaRubroNacion 
   Caption         =   "Ventas por Rubro Nacionalidad y Local"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   8205
   Begin VB.CheckBox chkLocal 
      Caption         =   "L22"
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
      Height          =   225
      Index           =   1
      Left            =   1065
      TabIndex        =   50
      Top             =   1590
      Width           =   705
   End
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   135
      TabIndex        =   48
      Top             =   1395
      Width           =   7965
      Begin VB.CheckBox chkLocal 
         Caption         =   "L02"
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
         Height          =   225
         Index           =   8
         Left            =   2340
         TabIndex        =   74
         Top             =   195
         Width           =   645
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "MEND"
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
         Height          =   225
         Index           =   7
         Left            =   6525
         TabIndex        =   65
         Top             =   210
         Width           =   885
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "CORD"
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
         Height          =   225
         Index           =   6
         Left            =   5490
         TabIndex        =   64
         Top             =   210
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         Height          =   405
         Left            =   7425
         Picture         =   "frmVtasRubroNacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   105
         Width           =   510
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "L14"
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
         Height          =   225
         Index           =   5
         Left            =   4665
         TabIndex        =   54
         Top             =   195
         Width           =   705
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "L09"
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
         Height          =   225
         Index           =   4
         Left            =   3825
         TabIndex        =   53
         Top             =   195
         Width           =   705
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "L03"
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
         Height          =   225
         Index           =   3
         Left            =   3045
         TabIndex        =   52
         Top             =   195
         Width           =   705
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "L01"
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
         Height          =   225
         Index           =   2
         Left            =   1665
         TabIndex        =   51
         Top             =   195
         Width           =   645
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "L21"
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
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   49
         Top             =   195
         Width           =   705
      End
   End
   Begin Crystal.CrystalReport rptform 
      Left            =   7875
      Top             =   2310
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327682
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab tabRubro 
      Height          =   4005
      Left            =   345
      TabIndex        =   2
      Top             =   1995
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   7064
      _Version        =   327680
      Tabs            =   8
      TabsPerRow      =   4
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
      TabCaption(0)   =   "ACC"
      TabPicture(0)   =   "frmVtasRubroNacion.frx":0442
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TabSubrubro(0)"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "BEB"
      TabPicture(1)   =   "frmVtasRubroNacion.frx":045E
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TabSubrubro(1)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "CIG"
      TabPicture(2)   =   "frmVtasRubroNacion.frx":047A
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TabSubrubro(2)"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "COM"
      TabPicture(3)   =   "frmVtasRubroNacion.frx":0496
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TabSubrubro(3)"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "COS"
      TabPicture(4)   =   "frmVtasRubroNacion.frx":04B2
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TabSubrubro(4)"
      Tab(4).Control(0).Enabled=   0   'False
      TabCaption(5)   =   "ELE"
      TabPicture(5)   =   "frmVtasRubroNacion.frx":04CE
      Tab(5).ControlCount=   1
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "TabSubrubro(5)"
      Tab(5).Control(0).Enabled=   0   'False
      TabCaption(6)   =   "PER"
      TabPicture(6)   =   "frmVtasRubroNacion.frx":04EA
      Tab(6).ControlCount=   1
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "TabSubrubro(6)"
      Tab(6).Control(0).Enabled=   0   'False
      TabCaption(7)   =   "TAB"
      TabPicture(7)   =   "frmVtasRubroNacion.frx":0506
      Tab(7).ControlCount=   1
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "TabSubrubro(7)"
      Tab(7).Control(0).Enabled=   0   'False
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   1
         Left            =   -74790
         TabIndex        =   25
         Top             =   795
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroNacion.frx":0522
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(1)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroNacion.frx":053E
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroNacion.frx":055A
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(1)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Participación"
         TabPicture(3)   =   "frmVtasRubroNacion.frx":0576
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SprRubroPartic(1)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPartic 
            Height          =   2400
            Index           =   1
            Left            =   -74895
            OleObjectBlob   =   "frmVtasRubroNacion.frx":0592
            TabIndex        =   67
            Top             =   90
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   1
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":09FB
            TabIndex        =   32
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   1
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":0D52
            TabIndex        =   39
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   1
            Left            =   90
            OleObjectBlob   =   "frmVtasRubroNacion.frx":10C5
            TabIndex        =   57
            Top             =   105
            Width           =   6705
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   0
         Left            =   225
         TabIndex        =   22
         Top             =   780
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroNacion.frx":152E
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroNacion.frx":154A
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(0)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroNacion.frx":1566
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(0)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Participación"
         TabPicture(3)   =   "frmVtasRubroNacion.frx":1582
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SprRubroPartic(0)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   0
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":159E
            TabIndex        =   23
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   0
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":18F5
            TabIndex        =   24
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   0
            Left            =   150
            OleObjectBlob   =   "frmVtasRubroNacion.frx":1C68
            TabIndex        =   56
            Top             =   120
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroPartic 
            Height          =   2400
            Index           =   0
            Left            =   -74820
            OleObjectBlob   =   "frmVtasRubroNacion.frx":20D1
            TabIndex        =   66
            Top             =   120
            Width           =   6705
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   2
         Left            =   -74820
         TabIndex        =   26
         Top             =   795
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroNacion.frx":253A
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(2)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroNacion.frx":2556
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(2)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroNacion.frx":2572
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(2)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Participación"
         TabPicture(3)   =   "frmVtasRubroNacion.frx":258E
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SprRubroPartic(2)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPartic 
            Height          =   2400
            Index           =   2
            Left            =   -74910
            OleObjectBlob   =   "frmVtasRubroNacion.frx":25AA
            TabIndex        =   68
            Top             =   75
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   2
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":2A13
            TabIndex        =   33
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   2
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":2D6A
            TabIndex        =   40
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   2
            Left            =   90
            OleObjectBlob   =   "frmVtasRubroNacion.frx":30DD
            TabIndex        =   58
            Top             =   105
            Width           =   6705
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   3
         Left            =   -74790
         TabIndex        =   27
         Top             =   795
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroNacion.frx":3546
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(3)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroNacion.frx":3562
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(3)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroNacion.frx":357E
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(3)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Participación"
         TabPicture(3)   =   "frmVtasRubroNacion.frx":359A
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SprRubroPartic(3)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPartic 
            Height          =   2400
            Index           =   3
            Left            =   -74910
            OleObjectBlob   =   "frmVtasRubroNacion.frx":35B6
            TabIndex        =   69
            Top             =   75
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   3
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":3A1F
            TabIndex        =   34
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   3
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":3D76
            TabIndex        =   41
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   3
            Left            =   90
            OleObjectBlob   =   "frmVtasRubroNacion.frx":40E9
            TabIndex        =   59
            Top             =   90
            Width           =   6705
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   4
         Left            =   -74805
         TabIndex        =   28
         Top             =   795
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroNacion.frx":4552
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(4)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroNacion.frx":456E
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(4)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroNacion.frx":458A
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(4)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Participación"
         TabPicture(3)   =   "frmVtasRubroNacion.frx":45A6
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SprRubroPartic(4)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPartic 
            Height          =   2400
            Index           =   4
            Left            =   -74910
            OleObjectBlob   =   "frmVtasRubroNacion.frx":45C2
            TabIndex        =   70
            Top             =   90
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   4
            Left            =   135
            OleObjectBlob   =   "frmVtasRubroNacion.frx":4A2B
            TabIndex        =   60
            Top             =   90
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   4
            Left            =   -74490
            OleObjectBlob   =   "frmVtasRubroNacion.frx":4E94
            TabIndex        =   35
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   4
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":51EB
            TabIndex        =   42
            Top             =   105
            Width           =   5505
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   5
         Left            =   -74790
         TabIndex        =   29
         Top             =   795
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroNacion.frx":555E
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(5)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroNacion.frx":557A
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(5)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroNacion.frx":5596
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(5)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Participación"
         TabPicture(3)   =   "frmVtasRubroNacion.frx":55B2
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SprRubroPartic(5)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPartic 
            Height          =   2400
            Index           =   5
            Left            =   -74925
            OleObjectBlob   =   "frmVtasRubroNacion.frx":55CE
            TabIndex        =   71
            Top             =   75
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   5
            Left            =   90
            OleObjectBlob   =   "frmVtasRubroNacion.frx":5A37
            TabIndex        =   61
            Top             =   75
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   5
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":5EA0
            TabIndex        =   36
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   5
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":61F7
            TabIndex        =   43
            Top             =   105
            Width           =   5505
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   6
         Left            =   -74805
         TabIndex        =   30
         Top             =   795
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroNacion.frx":656A
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(6)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroNacion.frx":6586
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(6)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroNacion.frx":65A2
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(6)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Participación"
         TabPicture(3)   =   "frmVtasRubroNacion.frx":65BE
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SprRubroPartic(6)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPartic 
            Height          =   2400
            Index           =   6
            Left            =   -74925
            OleObjectBlob   =   "frmVtasRubroNacion.frx":65DA
            TabIndex        =   72
            Top             =   75
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   6
            Left            =   105
            OleObjectBlob   =   "frmVtasRubroNacion.frx":6A43
            TabIndex        =   62
            Top             =   75
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   6
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":6EAC
            TabIndex        =   37
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   6
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":7203
            TabIndex        =   44
            Top             =   105
            Width           =   5505
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   7
         Left            =   -74805
         TabIndex        =   31
         Top             =   795
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroNacion.frx":7576
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(7)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroNacion.frx":7592
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(7)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroNacion.frx":75AE
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(7)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Participación"
         TabPicture(3)   =   "frmVtasRubroNacion.frx":75CA
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SprRubroPartic(7)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPartic 
            Height          =   2400
            Index           =   7
            Left            =   -74925
            OleObjectBlob   =   "frmVtasRubroNacion.frx":75E6
            TabIndex        =   73
            Top             =   75
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   7
            Left            =   90
            OleObjectBlob   =   "frmVtasRubroNacion.frx":7A4F
            TabIndex        =   63
            Top             =   90
            Width           =   6705
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   7
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":7EB8
            TabIndex        =   38
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   7
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroNacion.frx":820F
            TabIndex        =   45
            Top             =   105
            Width           =   5505
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1485
      Left            =   6675
      TabIndex        =   1
      Top             =   -30
      Width           =   1410
      Begin VB.CommandButton CmdReport 
         Height          =   600
         Left            =   60
         Picture         =   "frmVtasRubroNacion.frx":8582
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   825
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   630
         Left            =   60
         Picture         =   "frmVtasRubroNacion.frx":96B8
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   165
         Width           =   600
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
         Index           =   0
         Left            =   705
         Picture         =   "frmVtasRubroNacion.frx":9C4A
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   705
         Picture         =   "frmVtasRubroNacion.frx":9D4C
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   705
         Picture         =   "frmVtasRubroNacion.frx":9E4E
         Style           =   1  'Graphical
         TabIndex        =   11
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
      Width           =   8175
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
         Width           =   6465
         Begin VB.Frame frPerAnt 
            Caption         =   "Período Anterior"
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
            Left            =   3270
            TabIndex        =   14
            Top             =   225
            Width           =   3015
            Begin VB.CommandButton botHelpFDAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmVtasRubroNacion.frx":A670
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmVtasRubroNacion.frx":A7E2
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   600
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesdeAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   17
               Top             =   255
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-##-####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFHastaAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   18
               Top             =   600
               Width           =   1125
               _ExtentX        =   1984
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
               Index           =   2
               Left            =   150
               TabIndex        =   20
               Top             =   255
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
               Index           =   3
               Left            =   135
               TabIndex        =   19
               Top             =   600
               Width           =   1185
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Período "
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
            Left            =   180
            TabIndex        =   4
            Top             =   225
            Width           =   3015
            Begin VB.CommandButton botHelpFH 
               Height          =   345
               Left            =   2535
               Picture         =   "frmVtasRubroNacion.frx":A954
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2550
               Picture         =   "frmVtasRubroNacion.frx":AAC6
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   255
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesde 
               Height          =   285
               Left            =   1380
               TabIndex        =   7
               Top             =   270
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-##-####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFHasta 
               Height          =   285
               Left            =   1395
               TabIndex        =   8
               Top             =   615
               Width           =   1125
               _ExtentX        =   1984
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
               Left            =   135
               TabIndex        =   10
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
               Left            =   150
               TabIndex        =   9
               Top             =   270
               Width           =   1185
            End
         End
      End
   End
   Begin VB.Label labInfo 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   150
      TabIndex        =   21
      Top             =   5820
      Width           =   8730
   End
End
Attribute VB_Name = "frmVtaRubroNacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset
Dim RsDataAnt  As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas



Function L_Locales()
Dim i As Integer
Dim loc  As String

loc = ""
For i = 0 To 7
    If chkLocal(i).Value = 1 Then
       loc = loc & chkLocal(i).caption & "-"
    End If
Next

If loc = "" Then
   loc = "- Total Compañía -"
Else
   loc = "-" & loc
End If

L_Locales = loc
End Function


Private Sub L_PrepararTablaPRN()
Dim sql As String
Dim i As Integer, j As Integer
Dim valUniC As Variant, valImpC As Variant
Dim valUniN As Variant, valImpN As Variant
Dim valRubro As Variant, valLocal As Variant

Aplicacion.EjecutarDAO "TRUNCATE TABLE ESTADIS.PRN_RUBROLOCAL DROP STORAGE"

For i = 0 To 7
    valRubro = tabRubro.TabCaption(i)
    For j = 1 To SprRubroImporte(i).MaxRows
        SprRubroImporte(i).GetText 1, j, valLocal
        If Left(valLocal, 1) = "L" Then
           SprRubroImporte(i).GetText 2, j, valImpC
           SprRubroImporte(i).GetText 3, j, valImpN
           SprRubroUnidades(i).GetText 2, j, valUniC
           SprRubroUnidades(i).GetText 3, j, valUniN
           
           'sql = "INSERT INTO estadis.PRN_RUBROLOCAL ( " _
           '& " cod_depn , cod_sdep , " _
           '& " cod_local , cod_rubro , " _
           '& " act_$ , act_Un , " _
           '& " ant_$ , ant_Un ) VALUES ( " _
           '& "'" & L_DepnSdep(valLocal, "DEPN") & "' , " _
           '& "'" & L_DepnSdep(valLocal, "SDEP") & "' , " _
           '& "'" & valLocal & "' , " _
           '& "'" & valRubro & "' , " _
           '& valImpC & " , " & valUniC & " , " _
           '& valImpN & " , " & valUniN & " ) "
        
           Aplicacion.EjecutarDAO sql
        End If
    Next
Next

End Sub

Private Sub MeImprimir()
Dim i%

On Error GoTo ErrPrint:

rptform.Connect = "ODBC;UID=" & Aplicacion.username & ";PWD=" & Aplicacion.password & ";DATABASE=INTER;"

rptform.ReportFileName = RutaReportes
rptform.ReportFileName = RutaReportes & "RubLoc.rpt"

rptform.ParameterFields(0) = "PrActFechaDesde;" & " DATE(" & Year(mskFDesde.FormattedText) & "," & Month(mskFDesde.FormattedText) & "," & Day(mskFDesde.FormattedText) & ")" & ";true)"
rptform.ParameterFields(1) = "PrActFechaHasta;" & " DATE(" & Year(mskFHasta.FormattedText) & "," & Month(mskFHasta.FormattedText) & "," & Day(mskFHasta.FormattedText) & ")" & ";true)"
rptform.ParameterFields(2) = "PrAntFechaDesde;" & " DATE(" & Year(mskFDesdeAnt.FormattedText) & "," & Month(mskFDesdeAnt.FormattedText) & "," & Day(mskFDesdeAnt.FormattedText) & ")" & ";true)"
rptform.ParameterFields(3) = "PrAntFechaHasta;" & " DATE(" & Year(mskFHastaAnt.FormattedText) & "," & Month(mskFHastaAnt.FormattedText) & "," & Day(mskFHastaAnt.FormattedText) & ")" & ";true)"

rptform.Destination = 0     ' 0 pantalla
                            '1 impresora

rptform.Action = True
 DoEvents


Exit Sub
ErrPrint:
'     frmVtaRubroLocal.caption = Aplicacion.SeteoFin
     Exit Sub


End Sub

Public Function L_DecoRubroAnt()

Select Case RsData!cod_rubr
    
    Case "ACC"
       L_DecoRubroAnt = 0
    Case "BEB"
       L_DecoRubroAnt = 1
    Case "CIG"
       L_DecoRubroAnt = 2
    Case "COM"
       L_DecoRubroAnt = 3
    Case "COS"
       L_DecoRubroAnt = 4
    Case "ELE"
       L_DecoRubroAnt = 5
    Case "PER"
       L_DecoRubroAnt = 6
    Case "TAB"
       L_DecoRubroAnt = 7
    Case Else
       L_DecoRubroAnt = Null
End Select

End Function
Public Function L_DecoRubro()

Select Case RsData!cod_rubr
    
    Case "ACC"
       L_DecoRubro = 0
    Case "BEB"
       L_DecoRubro = 1
    Case "CIG"
       L_DecoRubro = 2
    Case "COM"
       L_DecoRubro = 3
    Case "COS"
       L_DecoRubro = 4
    Case "ELE"
       L_DecoRubro = 5
    Case "PER"
       L_DecoRubro = 6
    Case "TAB"
       L_DecoRubro = 7
    Case Else
       L_DecoRubro = Null
End Select

End Function



Private Function L_Armarcondicion(valor As Integer) As String
Dim Cond, cLocal As String
Dim cant

Dim fechaDesde As String
Dim fechaHasta As String
Dim fechaDesdeAnt As String
Dim fechaHastaAnt As String



fechaDesde = mskFDesde.FormattedText
fechaHasta = mskFHasta.FormattedText
fechaDesdeAnt = mskFDesdeAnt.FormattedText
fechaHastaAnt = mskFHastaAnt.FormattedText

Select Case valor
  Case Is = 1
     Cond = " And fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta) & "  "
  Case Is = 2
     Cond = " and fch_ticket between " & func_ToDate(fechaDesdeAnt) & " And " & func_ToDate(fechaHastaAnt) & "  "
End Select

'Son casi todos los casos de tratamiento especial
cLocal = ""
If chkLocal(0).Value = 1 Then
    cLocal = cLocal & "'L21','L05','L06',"
End If
If chkLocal(1).Value = 1 Then
    cLocal = cLocal & "'L22','L07','L08',"
End If
If chkLocal(2).Value = 1 Then
    cLocal = cLocal & "'L01',"
End If
If chkLocal(3).Value = 1 Then
    cLocal = cLocal & "'L03',"
End If
If chkLocal(8).Value = 1 Then
    cLocal = cLocal & "'L02',"
End If
If chkLocal(4).Value = 1 Then
    cLocal = cLocal & "'L09',"
End If
If chkLocal(5).Value = 1 Then
    cLocal = cLocal & "'L14',"
End If
If chkLocal(6).Value = 1 Then
    cLocal = cLocal & "'L04','L19',"
End If
If chkLocal(7).Value = 1 Then
    cLocal = cLocal & "'L10','L16',"
End If
   
   If cLocal <> "" Then
      Cond = Cond & " And cod_local IN (" & Mid(cLocal, 1, Len(cLocal) - 1) & ")"
   End If

L_Armarcondicion = Cond

End Function

Private Sub L_LimpiarGrillas()
Dim i
labInfo.caption = ""

For i = 0 To 7
  SprRubroImporte(i).MaxRows = 0
  SprRubroUnidades(i).MaxRows = 0
  SprRubroPromedio(i).MaxRows = 0
  SprRubroPartic(i).MaxRows = 0
Next i


End Sub

Public Sub L_LlenarNacion()

Dim sqlL As String
Dim n As Integer
Dim RsDataNac As Recordset
Dim NAC As String
Dim linea As Integer

sqlL = " SELECT distinct decode(cod_pais,0,'Resto del Mundo',9,'Resto del Mundo',descrip) descrip ,"
sqlL = sqlL & " decode(cod_pais,0,8,9,8,cod_pais) cod_pais "
sqlL = sqlL & " FROM ventas.paises "
sqlL = sqlL & " ORDER BY decode(cod_pais,0,'Resto del Mundo',9,'Resto del Mundo',descrip) "

If Aplicacion.ObtenerRsDAO(sqlL, RsDataNac) Then
   If Aplicacion.CantReg(RsData) > 0 Then
      For n = 0 To 7
       SprRubroImporte(n).MaxRows = 0
       SprRubroUnidades(n).MaxRows = 0
       SprRubroPromedio(n).MaxRows = 0
      Next
      NAC = RsDataNac!Descrip
      Do While Not RsDataNac.EOF
         For n = 0 To 7
           SprRubroImporte(n).MaxRows = SprRubroImporte(n).MaxRows + 1
           SprRubroUnidades(n).MaxRows = SprRubroUnidades(n).MaxRows + 1
           SprRubroPromedio(n).MaxRows = SprRubroPromedio(n).MaxRows + 1
           SprRubroPartic(n).MaxRows = SprRubroPartic(n).MaxRows + 1
           
           SprRubroImporte(n).SetText 1, SprRubroImporte(n).MaxRows, Trim(RsDataNac!Descrip)
           SprRubroImporte(n).SetText 7, SprRubroImporte(n).MaxRows, Trim(RsDataNac!cod_pais)
           SprRubroUnidades(n).SetText 1, SprRubroUnidades(n).MaxRows, Trim(RsDataNac!Descrip)
           SprRubroPromedio(n).SetText 1, SprRubroPromedio(n).MaxRows, Trim(RsDataNac!Descrip)
           SprRubroPartic(n).SetText 1, SprRubroPartic(n).MaxRows, Trim(RsDataNac!Descrip)

         Next n
           RsDataNac.MoveNext
      Loop
   
   End If
End If
End Sub

Private Sub L_Participacion(sprSet As control, sprGet As control)
Dim i As Integer
Dim fila As Integer
Dim valor As Variant
Dim Total As Variant
Dim ValorAnt As Variant
Dim TotalANT As Variant
Dim part As Single
   
   sprGet.GetText 2, sprGet.MaxRows, Total
   sprGet.GetText 3, sprGet.MaxRows, TotalANT
   
   
   For fila = 1 To sprGet.MaxRows - 1
       sprGet.GetText 2, fila, valor
       sprGet.GetText 3, fila, ValorAnt
       part = 0
       If Total > 0 Then
           sprSet.SetText 2, fila, str(valor * 100 / Total)
       End If
       If TotalANT > 0 Then
           sprSet.SetText 3, fila, str(ValorAnt * 100 / TotalANT)
       End If
       'If ValorAnt > 0 Then
       '   part = 100 * (1 - ((valor * 100 / TotalGral) / (ValorAnt * 100 / TotalGralant)))
       'End If
       'If part <> 0 Then
       '   SprImpResumen.SetText 10, fila, Format(part, "###0.00")
       'End If
   Next fila

'   SprImpResumen.SetText 4, SprImpResumen.MaxRows, "."
'   SprImpResumen.SetText 7, SprImpResumen.MaxRows, "."
End Sub


Public Function L_Division(divisor As Double, dividendo As Double) As Double

If dividendo = 0 Or divisor = 0 Then
   L_Division = 0
 Else
   L_Division = divisor / dividendo
End If

End Function
Public Sub L_PonerDatos(spr As control, Spr1 As control, spr2 As control, Col As Integer)
Dim rubro As String
Dim SubDepn As String
Dim Fila_Desde As Integer
Dim Imp_Total As Double
Dim Uni_Total As Double
Dim linea As Integer
Dim locales As Variant

rubro = RsData!cod_rubr
linea = 1

Do While Not RsData.EOF And (RsData!cod_rubr = rubro)
   spr.GetText 1, linea, locales
   If Trim(locales) = RsData!Descrip Then
      spr.SetText Col, linea, str(RsData!Importe)
      Imp_Total = Imp_Total + RsData!Importe
      Spr1.SetText Col, linea, str(RsData!unidades)
      Uni_Total = Uni_Total + RsData!unidades
      spr2.SetText Col, linea, str(L_Division(RsData!Importe, RsData!unidades))
      RsData.MoveNext
      If RsData.EOF Then
         Exit Do
      End If
   End If
   
   linea = linea + 1

Loop
  
End Sub

Private Sub L_Sumarizar_(spr As control, Row_D As Integer, Row_H As Integer, Col_D As Integer, Col_H As Integer)
Dim n As Integer
Dim valor1 As Long
Dim Total As Long

For n = Row_D To Row_H
    spr.GetText n, Col_H, valor1
    Total = Total + valor1
Next n

spr.SetText spr.MaxRows + 1, Col_H, Total

End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i As Integer

'On Error GoTo ErrRubLoc:

frmVtaRubroNacion.caption = Aplicacion.SeteoProceso(frmVtaRubroNacion.caption)

'Consulta Actual
sql = " SELECT decode(nacionalidad,0,'Resto del Mundo',9,'Resto del Mundo',descrip) descrip ,"
sql = sql & " COD_RUBR , SUM(IMPORTE) importe, SUM(UNIDADES) unidades"
sql = sql & " FROM ESTADIS.VENTA_RNC R, VENTAS.PAISES P "
sql = sql & " Where COD_PAIS = NACIONALIDAD AND COD_RUBR <> 'REG' "
sql = sql & L_Armarcondicion(1)
sql = sql & " GROUP BY COD_RUBR,decode(nacionalidad,0,'Resto del Mundo',9,'Resto del Mundo',descrip)   "


'Consulta Anterior
sqlX = " SELECT decode(nacionalidad,0,'Resto del Mundo',9,'Resto del Mundo',descrip) descrip, "
sqlX = sqlX & " COD_RUBR , SUM(IMPORTE) importe, SUM(UNIDADES) unidades"
sqlX = sqlX & " FROM ESTADIS.VENTA_RNC R, VENTAS.PAISES P "
sqlX = sqlX & " Where COD_PAIS = NACIONALIDAD AND COD_RUBR <> 'REG' "
sqlX = sqlX & L_Armarcondicion(2)
sqlX = sqlX & " GROUP BY  COD_RUBR,decode(nacionalidad,0,'Resto del Mundo',9,'Resto del Mundo',descrip)  "

If Aplicacion.ObtenerRsDAO(sql, RsData) Then
   
    If Aplicacion.CantReg(RsData) > 0 Then
        L_LlenarNacion
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabRubro.Enabled = True
        
        L_DecoEspigon
        
        If Aplicacion.ObtenerRsDAO(sqlX, RsData) Then
           L_DecoEspigonAnt
        End If
        
        L_Resaltar
            
        For i = 0 To 7
            Spread_TotalesGrillas SprRubroImporte(i), 2, 2
            L_Participacion SprRubroPartic(i), SprRubroImporte(i)
        Next
        
    End If
    
    Aplicacion.CerrarDAO RsData

End If

ErrRubLoc:
    frmVtaRubroNacion.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub




Private Sub L_Resaltar()

Dim fila As Integer
Dim i

For i = 0 To 7
  For fila = 1 To SprRubroImporte(i).MaxRows
     Spread.spread_ResaltarCelda SprRubroImporte(i), 4, fila
    
     Spread.spread_ResaltarCelda SprRubroUnidades(i), 4, fila
    
     Spread.spread_ResaltarCelda SprRubroPromedio(i), 4, fila

  Next fila
Next i

End Sub


Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant
Dim x As Variant
Dim rubro As String

Do While Not RsData.EOF
   
    rubro = RsData!cod_rubr
   
   x = L_DecoRubro
   Do While RsData!cod_rubr = rubro And Not RsData.EOF
        If IsNumeric(x) Then
           L_PonerDatos SprRubroImporte(x), SprRubroUnidades(x), SprRubroPromedio(x), 2
          Else
              RsData.MoveNext
        End If
        If RsData.EOF Then
            Exit Do
        End If
   
   Loop

Loop

End Sub


Private Sub L_DecoEspigonAnt()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant
Dim x As Variant
Dim rubro As String

Do While Not RsData.EOF
   
   rubro = RsData!cod_rubr
   x = L_DecoRubroAnt
   Do While RsData!cod_rubr = rubro
        If IsNumeric(x) Then
           L_PonerDatos SprRubroImporte(x), SprRubroUnidades(x), SprRubroPromedio(x), 3
        Else
           RsDataAnt.MoveNext
        End If
        If RsData.EOF Then
          Exit Do
        End If
   Loop

Loop

End Sub

Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        L_Refrescar
    Case 1

        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        tabRubro.Enabled = False
        L_LimpiarGrillas
        
    Case 2
        Unload Me
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


Private Sub CmdExcel_Click()
FrmRubro.Text1 = 2
FrmRubro.Show 1
End Sub

Private Sub CmdReport_Click()
Dim sql

On Error GoTo ErrPRN:

frmVtaRubroLocal.caption = Aplicacion.SeteoProceso(frmVtaRubroLocal.caption)

Aplicacion.ComienzoTrans
If Aplicacion.EjecutarDAO("lock table estadis.prn_rubrolocal in exclusive mode nowait") Then
    L_PrepararTablaPRN
Else
    MsgBox "POR FAVOR REINTENTAR EN POCOS MINUTOS...", vbOKOnly, "ATENCION"
End If

Aplicacion.TerminarConExitoTrans
MeImprimir

frmVtaRubroLocal.caption = Aplicacion.SeteoFin

Exit Sub
ErrPRN:
    Resume
End Sub

Private Sub Command1_Click()
  L_LimpiarGrillas
  L_Refrescar
  
End Sub

Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabRubro

End Sub

Private Sub Form_Load()
Dim i

Width = Screen.Width * 0.7
Height = Screen.Height * 0.8
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height - 500) / 2

If Day(Date) = 1 Then
    mskFDesde.Text = Format$(Func.func_Dia1SegunMes_Anio(Month(Date - 1), Year(Date - 1)), FTOFECHA)
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
    
    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio - 1, FTOFECHA)
    mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio - 1, FTOFECHA)
    
Else
    mskFDesde.Text = "01-" & Format$(Month(Date), "0#") & "-" & Format$(Year(Date), "####")
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
    
    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
    mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio, FTOFECHA)
    
End If
    

L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmVtaRubroLocal"

End Sub


Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmVtaRubroLocal"
End Sub


Private Sub mskFDesde_LostFocus()

    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
       Else
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
        mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
     Else
    End If
    
    mskFDesdeAnt.Text = Format$(mskFDesdeAnt.FormattedText, FTOFECHA)
    If mskFHastaAnt.Text <> "" Then
        If CDate(mskFHastaAnt.FormattedText) < CDate(mskFDesdeAnt.FormattedText) Or Year(mskFHastaAnt.FormattedText) <> Year(mskFDesdeAnt.FormattedText) Then
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




Private Sub SprRubroImporte_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim sql As String
Dim NAC As Variant, Total As Variant, TotalANT As Variant
Dim DescNac As Variant
Dim d As String

 If Index = 0 Then
    d = "ACC"
 ElseIf Index = 1 Then
    d = "BEB"
 ElseIf Index = 2 Then
    d = "CIG"
 ElseIf Index = 3 Then
    d = "COM"
 ElseIf Index = 4 Then
    d = "COS"
 ElseIf Index = 5 Then
    d = "ELE"
 ElseIf Index = 6 Then
    d = "PER"
 ElseIf Index = 7 Then
    d = "TAB"
 End If

 SprRubroImporte(Index).GetText 7, Row, NAC
 SprRubroImporte(Index).GetText 1, Row, DescNac
 SprRubroImporte(Index).GetText 2, SprRubroImporte(Index).MaxRows, Total
 SprRubroImporte(Index).GetText 3, SprRubroImporte(Index).MaxRows, TotalANT

 If NAC = 8 Then
    NAC = NAC & ",0,9"
 End If
 
If Col = 5 Then
 sql = " Select descr, "
 sql = sql & " sum(importe_act) importe_act ,sum(unidades_act) unidades_act, "
 sql = sql & " sum(importe_ant) importe_ant ,sum(unidades_ant) unidades_ant"
 sql = sql & " From ( "
 sql = sql & " SELECT DESCR,SUM(IMPORTE) IMPORTE_ACT, SUM(UNIDADES) UNIDADES_ACT "
 sql = sql & " , 0 IMPORTE_ANT,0 UNIDADES_ANT "
 sql = sql & " FROM ESTADIS.VENTA_RNC V, BAIRES.SUBRUBRO R "
 sql = sql & " WHERE "
 sql = sql & " V.cod_rubr = '" & d & "' and v.nacionalidad in (" & NAC & ") "
 sql = sql & " and V.COD_SRUB = R.COD_SRUB and v.cod_rubr = r.cod_rubr "
 sql = sql & L_Armarcondicion(1)
 sql = sql & " group by descr "
 sql = sql & " UNION "
 sql = sql & " SELECT DESCR,0 IMPORTE_ACT, 0 UNIDADES_ACT,"
 sql = sql & " SUM(IMPORTE) IMPORTE_ANT,SUM(UNIDADES)  UNIDADES_ANT "
 sql = sql & " FROM ESTADIS.VENTA_RNC V, BAIRES.SUBRUBRO R "
 sql = sql & " WHERE "
 sql = sql & " V.cod_rubr = '" & d & "' and v.nacionalidad in (" & NAC & ") "
 sql = sql & " and V.COD_SRUB = R.COD_SRUB and v.cod_rubr = r.cod_rubr "
 sql = sql & L_Armarcondicion(2)
 sql = sql & " group by descr  )"
 sql = sql & " group by descr "
 sql = sql & " order by descr "


 frmVtaSubRubro.Mostrar "", Trim(d & " - " & DescNac), sql, Val(Total), Val(TotalANT), "Venta por Sub Rubro y Nacionalidad"
ElseIf Col = 6 Then
 sql = " Select descrip descr , "
 sql = sql & " sum(importe_act) importe_act ,sum(unidades_act) unidades_act, "
 sql = sql & " sum(importe_ant) importe_ant ,sum(unidades_ant) unidades_ant"
 sql = sql & " From ( "
 sql = sql & " SELECT DESCRIP,SUM(IMPORTE) IMPORTE_ACT, SUM(UNIDADES) UNIDADES_ACT "
 sql = sql & " , 0 IMPORTE_ANT,0 UNIDADES_ANT "
 sql = sql & " FROM ESTADIS.VENTA_RNC V, BAIRES.proveedor R "
 sql = sql & " WHERE "
 sql = sql & " V.cod_rubr = '" & d & "' and v.nacionalidad in (" & NAC & ") "
 sql = sql & " and V.COD_PROV = R.COD_PROV  "
 sql = sql & L_Armarcondicion(1)
 sql = sql & " group by descrip "
 sql = sql & " UNION "
 sql = sql & " SELECT DESCRip,0 IMPORTE_ACT, 0 UNIDADES_ACT,"
 sql = sql & " SUM(IMPORTE) IMPORTE_ANT,SUM(UNIDADES)  UNIDADES_ANT "
 sql = sql & " FROM ESTADIS.VENTA_RNC V, BAIRES.proveedor R "
 sql = sql & " WHERE "
 sql = sql & " V.cod_rubr = '" & d & "' and v.nacionalidad in (" & NAC & ") "
 sql = sql & " and V.COD_PROV = R.COD_PROV  "
 sql = sql & L_Armarcondicion(2)
 sql = sql & " group by descrip  )"
 sql = sql & " group by descrip "
 sql = sql & " order by descrip "

 frmVtaSubRubro.Mostrar "", Trim(d & " - " & DescNac), sql, Val(Total), Val(TotalANT), "Venta por Proveedor y Nacionalidad"
End If

End Sub


Private Sub SprRubroPartic_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim sql As String
Dim NAC As Variant, Total As Variant, TotalANT As Variant
Dim DescNac As Variant
Dim d As String

 If Index = 0 Then
    d = "ACC"
 ElseIf Index = 1 Then
    d = "BEB"
 ElseIf Index = 2 Then
    d = "CIG"
 ElseIf Index = 3 Then
    d = "COM"
 ElseIf Index = 4 Then
    d = "COS"
 ElseIf Index = 5 Then
    d = "ELE"
 ElseIf Index = 6 Then
    d = "PER"
 ElseIf Index = 7 Then
    d = "TAB"
 End If

 SprRubroImporte(Index).GetText 7, Row, NAC
 SprRubroImporte(Index).GetText 1, Row, DescNac
 SprRubroImporte(Index).GetText 2, SprRubroImporte(Index).MaxRows, Total
 SprRubroImporte(Index).GetText 3, SprRubroImporte(Index).MaxRows, TotalANT

 If NAC = 8 Then
    NAC = NAC & ",0,9"
 End If
 
If Col = 5 Then
 sql = " Select descr, "
 sql = sql & " sum(importe_act) importe_act ,sum(unidades_act) unidades_act, "
 sql = sql & " sum(importe_ant) importe_ant ,sum(unidades_ant) unidades_ant"
 sql = sql & " From ( "
 sql = sql & " SELECT DESCR,SUM(IMPORTE) IMPORTE_ACT, SUM(UNIDADES) UNIDADES_ACT "
 sql = sql & " , 0 IMPORTE_ANT,0 UNIDADES_ANT "
 sql = sql & " FROM ESTADIS.VENTA_RNC V, BAIRES.SUBRUBRO R "
 sql = sql & " WHERE "
 sql = sql & " V.cod_rubr = '" & d & "' and v.nacionalidad in (" & NAC & ") "
 sql = sql & " and V.COD_SRUB = R.COD_SRUB and v.cod_rubr = r.cod_rubr "
 sql = sql & L_Armarcondicion(1)
 sql = sql & " group by descr "
 sql = sql & " UNION "
 sql = sql & " SELECT DESCR,0 IMPORTE_ACT, 0 UNIDADES_ACT,"
 sql = sql & " SUM(IMPORTE) IMPORTE_ANT,SUM(UNIDADES)  UNIDADES_ANT "
 sql = sql & " FROM ESTADIS.VENTA_RNC V, BAIRES.SUBRUBRO R "
 sql = sql & " WHERE "
 sql = sql & " V.cod_rubr = '" & d & "' and v.nacionalidad in (" & NAC & ") "
 sql = sql & " and V.COD_SRUB = R.COD_SRUB and v.cod_rubr = r.cod_rubr "
 sql = sql & L_Armarcondicion(2)
 sql = sql & " group by descr  )"
 sql = sql & " group by descr "
 sql = sql & " order by descr "


 frmVtaSubRubro.Mostrar "", Trim(d & " - " & DescNac), sql, Val(Total), Val(TotalANT), "Venta por Sub Rubro y Nacionalidad"
ElseIf Col = 6 Then
 sql = " Select descrip descr , "
 sql = sql & " sum(importe_act) importe_act ,sum(unidades_act) unidades_act, "
 sql = sql & " sum(importe_ant) importe_ant ,sum(unidades_ant) unidades_ant"
 sql = sql & " From ( "
 sql = sql & " SELECT DESCRIP,SUM(IMPORTE) IMPORTE_ACT, SUM(UNIDADES) UNIDADES_ACT "
 sql = sql & " , 0 IMPORTE_ANT,0 UNIDADES_ANT "
 sql = sql & " FROM ESTADIS.VENTA_RNC V, BAIRES.proveedor R "
 sql = sql & " WHERE "
 sql = sql & " V.cod_rubr = '" & d & "' and v.nacionalidad in (" & NAC & ") "
 sql = sql & " and V.COD_PROV = R.COD_PROV  "
 sql = sql & L_Armarcondicion(1)
 sql = sql & " group by descrip "
 sql = sql & " UNION "
 sql = sql & " SELECT DESCRip,0 IMPORTE_ACT, 0 UNIDADES_ACT,"
 sql = sql & " SUM(IMPORTE) IMPORTE_ANT,SUM(UNIDADES)  UNIDADES_ANT "
 sql = sql & " FROM ESTADIS.VENTA_RNC V, BAIRES.proveedor R "
 sql = sql & " WHERE "
 sql = sql & " V.cod_rubr = '" & d & "' and v.nacionalidad in (" & NAC & ") "
 sql = sql & " and V.COD_PROV = R.COD_PROV  "
 sql = sql & L_Armarcondicion(2)
 sql = sql & " group by descrip  )"
 sql = sql & " group by descrip "
 sql = sql & " order by descrip "

 frmVtaSubRubro.Mostrar "", Trim(d & " - " & DescNac), sql, Val(Total), Val(TotalANT), "Venta por Proveedor y Nacionalidad"
End If

End Sub


