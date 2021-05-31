VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.0#0"; "crystl32.ocx"
Begin VB.Form frmVtaRubroLocal 
   Caption         =   "Ventas por Rubro/Local"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   8205
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
      Top             =   1635
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
      TabPicture(0)   =   "frmVtasRubroLocal.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TabSubrubro(0)"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "BEB"
      TabPicture(1)   =   "frmVtasRubroLocal.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TabSubrubro(1)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "CIG"
      TabPicture(2)   =   "frmVtasRubroLocal.frx":0038
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TabSubrubro(2)"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "COM"
      TabPicture(3)   =   "frmVtasRubroLocal.frx":0054
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TabSubrubro(3)"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "COS"
      TabPicture(4)   =   "frmVtasRubroLocal.frx":0070
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TabSubrubro(4)"
      Tab(4).Control(0).Enabled=   0   'False
      TabCaption(5)   =   "ELE"
      TabPicture(5)   =   "frmVtasRubroLocal.frx":008C
      Tab(5).ControlCount=   1
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "TabSubrubro(5)"
      Tab(5).Control(0).Enabled=   0   'False
      TabCaption(6)   =   "PER"
      TabPicture(6)   =   "frmVtasRubroLocal.frx":00A8
      Tab(6).ControlCount=   1
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "TabSubrubro(6)"
      Tab(6).Control(0).Enabled=   0   'False
      TabCaption(7)   =   "TAB"
      TabPicture(7)   =   "frmVtasRubroLocal.frx":00C4
      Tab(7).ControlCount=   1
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "TabSubrubro(7)"
      Tab(7).Control(0).Enabled=   0   'False
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   1
         Left            =   -74600
         TabIndex        =   26
         Top             =   800
         Width           =   6500
         _ExtentX        =   11456
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroLocal.frx":00E0
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(1)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroLocal.frx":00FC
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroLocal.frx":0118
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(1)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   1
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":0134
            TabIndex        =   47
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   1
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":04A7
            TabIndex        =   40
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   1
            Left            =   495
            OleObjectBlob   =   "frmVtasRubroLocal.frx":07FE
            TabIndex        =   33
            Top             =   105
            Width           =   5505
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   0
         Left            =   400
         TabIndex        =   22
         Top             =   800
         Width           =   6500
         _ExtentX        =   11456
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroLocal.frx":0B71
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroLocal.frx":0B8D
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(0)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroLocal.frx":0BA9
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(0)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   0
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":0BC5
            TabIndex        =   25
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   0
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":0F38
            TabIndex        =   24
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   0
            Left            =   495
            OleObjectBlob   =   "frmVtasRubroLocal.frx":128F
            TabIndex        =   23
            Top             =   105
            Width           =   5505
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   2
         Left            =   -74600
         TabIndex        =   27
         Top             =   800
         Width           =   6500
         _ExtentX        =   11456
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroLocal.frx":1602
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(2)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroLocal.frx":161E
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(2)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroLocal.frx":163A
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   2
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":1656
            TabIndex        =   48
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   2
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":19C9
            TabIndex        =   41
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   2
            Left            =   495
            OleObjectBlob   =   "frmVtasRubroLocal.frx":1D20
            TabIndex        =   34
            Top             =   105
            Width           =   5505
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   3
         Left            =   -74600
         TabIndex        =   28
         Top             =   800
         Width           =   6500
         _ExtentX        =   11456
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroLocal.frx":2093
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(3)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroLocal.frx":20AF
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(3)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroLocal.frx":20CB
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(3)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   3
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":20E7
            TabIndex        =   49
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   3
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":245A
            TabIndex        =   42
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   3
            Left            =   495
            OleObjectBlob   =   "frmVtasRubroLocal.frx":27B1
            TabIndex        =   35
            Top             =   105
            Width           =   5505
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   4
         Left            =   -74600
         TabIndex        =   29
         Top             =   800
         Width           =   6500
         _ExtentX        =   11456
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroLocal.frx":2B24
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(4)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroLocal.frx":2B40
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(4)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroLocal.frx":2B5C
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(4)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   4
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":2B78
            TabIndex        =   50
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   4
            Left            =   -74490
            OleObjectBlob   =   "frmVtasRubroLocal.frx":2EEB
            TabIndex        =   43
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   4
            Left            =   495
            OleObjectBlob   =   "frmVtasRubroLocal.frx":3242
            TabIndex        =   36
            Top             =   105
            Width           =   5505
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   5
         Left            =   -74600
         TabIndex        =   30
         Top             =   800
         Width           =   6500
         _ExtentX        =   11456
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroLocal.frx":35B5
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(5)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroLocal.frx":35D1
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(5)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroLocal.frx":35ED
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(5)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   5
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":3609
            TabIndex        =   51
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   5
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":397C
            TabIndex        =   44
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   5
            Left            =   495
            OleObjectBlob   =   "frmVtasRubroLocal.frx":3CD3
            TabIndex        =   37
            Top             =   105
            Width           =   5505
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   6
         Left            =   -74595
         TabIndex        =   31
         Top             =   800
         Width           =   6500
         _ExtentX        =   11456
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroLocal.frx":4046
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprRubroImporte(6)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroLocal.frx":4062
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprRubroUnidades(6)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroLocal.frx":407E
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(6)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   6
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":409A
            TabIndex        =   52
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   6
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":440D
            TabIndex        =   45
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   6
            Left            =   495
            OleObjectBlob   =   "frmVtasRubroLocal.frx":4764
            TabIndex        =   38
            Top             =   105
            Width           =   5505
         End
      End
      Begin TabDlg.SSTab TabSubrubro 
         Height          =   3000
         Index           =   7
         Left            =   -74600
         TabIndex        =   32
         Top             =   800
         Width           =   6500
         _ExtentX        =   11456
         _ExtentY        =   5292
         _Version        =   327680
         TabOrientation  =   1
         Tab             =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroLocal.frx":4AD7
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "SprRubroImporte(7)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroLocal.frx":4AF3
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "SprRubroUnidades(7)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroLocal.frx":4B0F
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(7)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   7
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":4B2B
            TabIndex        =   53
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   2400
            Index           =   7
            Left            =   495
            OleObjectBlob   =   "frmVtasRubroLocal.frx":4E9E
            TabIndex        =   46
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   2400
            Index           =   7
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroLocal.frx":51F5
            TabIndex        =   39
            Top             =   105
            Width           =   5505
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1500
      Left            =   6675
      TabIndex        =   1
      Top             =   -15
      Width           =   1455
      Begin VB.CommandButton CmdReport 
         Height          =   615
         Left            =   60
         Picture         =   "frmVtasRubroLocal.frx":5568
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   825
         Width           =   600
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   630
         Left            =   60
         Picture         =   "frmVtasRubroLocal.frx":669E
         Style           =   1  'Graphical
         TabIndex        =   54
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
         Picture         =   "frmVtasRubroLocal.frx":6C30
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
         Picture         =   "frmVtasRubroLocal.frx":6D32
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
         Picture         =   "frmVtasRubroLocal.frx":6E34
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
               Picture         =   "frmVtasRubroLocal.frx":7656
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmVtasRubroLocal.frx":77C8
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
               Picture         =   "frmVtasRubroLocal.frx":793A
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2550
               Picture         =   "frmVtasRubroLocal.frx":7AAC
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
      Left            =   135
      TabIndex        =   21
      Top             =   5445
      Width           =   8730
   End
End
Attribute VB_Name = "frmVtaRubroLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset
Dim RsDataAnt  As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas



Private Function L_DepnSdep(loc As Variant, tipo As String)
If tipo = "DEPN" Then
    Select Case loc
        Case "L01", "L02", "L03", "L07", "L05", "L06", "L08", "L22"
            L_DepnSdep = "EZE"
        Case "L09", "L14"
            L_DepnSdep = "AEP"
        Case "L10", "L16", "L15", "L04", "L18", "L11", "L19"
            L_DepnSdep = "INT"
    End Select
Else
    Select Case loc
        Case "L01", "L02", "L03"
            L_DepnSdep = "INTB"
        Case "L07", "L05", "L06", "L08", "L21", "L22"
            L_DepnSdep = "INTA"
        Case "L09", "L14"
            L_DepnSdep = "AEP"
        Case "L10", "L16"
            L_DepnSdep = "MEND"
        Case "L15"
            L_DepnSdep = "MDPL"
        Case "L04", "L19"
            L_DepnSdep = "CORD"
        Case "L18"
            L_DepnSdep = "BARI"
        Case "L11"
            L_DepnSdep = "IGUA"
    End Select
End If

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
           
           sql = "INSERT INTO estadis.PRN_RUBROLOCAL ( " _
           & " cod_depn , cod_sdep , " _
           & " cod_local , cod_rubro , " _
           & " act_$ , act_Un , " _
           & " ant_$ , ant_Un ) VALUES ( " _
           & "'" & L_DepnSdep(valLocal, "DEPN") & "' , " _
           & "'" & L_DepnSdep(valLocal, "SDEP") & "' , " _
           & "'" & valLocal & "' , " _
           & "'" & valRubro & "' , " _
           & valImpC & " , " & valUniC & " , " _
           & valImpN & " , " & valUniN & " ) "
        
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

Select Case RsDataAnt!cod_rubr
    
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
Dim Cond
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
     Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta) & " And Cod_rubr <> 'REG' "
  Case Is = 2
     Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesdeAnt) & " And " & func_ToDate(fechaHastaAnt) & " And Cod_rubr <> 'REG' "
End Select

L_Armarcondicion = Cond

End Function

Private Sub L_LimpiarGrillas()
Dim i
labInfo.caption = ""

For i = 0 To 7
  SprRubroImporte(i).MaxRows = 0
  SprRubroUnidades(i).MaxRows = 0
  SprRubroPromedio(i).MaxRows = 0
Next i


End Sub

Public Sub L_LlenarLocales()

Dim sqlL As String
Dim n As Integer
Dim RsDataLoc As Recordset
Dim Sdep As String
Dim depn As String
Dim linea As Integer

sqlL = " SELECT distinct cod_depn,cod_sdep,cod_loc "
sqlL = sqlL & " FROM ventas.local_sublocal Where cod_loc NOT IN  ('L21','L07','L08') "
sqlL = sqlL & " ORDER BY cod_depn,cod_sdep,cod_loc "

If Aplicacion.ObtenerRsDAO(sqlL, RsDataLoc) Then
   If Aplicacion.CantReg(RsData) > 0 Then
      For n = 0 To 7
       SprRubroImporte(n).MaxRows = 0
       SprRubroUnidades(n).MaxRows = 0
       SprRubroPromedio(n).MaxRows = 0
      Next
      depn = RsDataLoc!cod_depn & " " & RsDataLoc!Cod_Sdep
      Do While Not RsDataLoc.EOF
          Select Case RsDataLoc!cod_depn
                        
            Case DSLoc(1).Dep
                For n = 0 To 7
                 SprRubroImporte(n).MaxRows = SprRubroImporte(n).MaxRows + 1
                 SprRubroUnidades(n).MaxRows = SprRubroUnidades(n).MaxRows + 1
                 SprRubroPromedio(n).MaxRows = SprRubroPromedio(n).MaxRows + 1
                 SprRubroImporte(n).SetText 1, SprRubroImporte(n).MaxRows, Trim(RsDataLoc!cod_loc)
                 SprRubroUnidades(n).SetText 1, SprRubroUnidades(n).MaxRows, Trim(RsDataLoc!cod_loc)
                 SprRubroPromedio(n).SetText 1, SprRubroPromedio(n).MaxRows, Trim(RsDataLoc!cod_loc)
                Next n
            
            Case DSLoc(2).Dep
               Select Case RsDataLoc!Cod_Sdep
                  Case DSLoc(2).Sdep
                     For n = 0 To 7
                        SprRubroImporte(n).MaxRows = SprRubroImporte(n).MaxRows + 1
                        SprRubroUnidades(n).MaxRows = SprRubroUnidades(n).MaxRows + 1
                        SprRubroPromedio(n).MaxRows = SprRubroPromedio(n).MaxRows + 1
                        SprRubroImporte(n).SetText 1, SprRubroImporte(n).MaxRows, Trim(RsDataLoc!cod_loc)
                        SprRubroUnidades(n).SetText 1, SprRubroUnidades(n).MaxRows, Trim(RsDataLoc!cod_loc)
                        SprRubroPromedio(n).SetText 1, SprRubroPromedio(n).MaxRows, Trim(RsDataLoc!cod_loc)
                     Next n
                  Case DSLoc(3).Sdep
                
                     For n = 0 To 7
                        SprRubroImporte(n).MaxRows = SprRubroImporte(n).MaxRows + 1
                        SprRubroUnidades(n).MaxRows = SprRubroUnidades(n).MaxRows + 1
                        SprRubroPromedio(n).MaxRows = SprRubroPromedio(n).MaxRows + 1
                        SprRubroImporte(n).SetText 1, SprRubroImporte(n).MaxRows, Trim(RsDataLoc!cod_loc)
                        SprRubroUnidades(n).SetText 1, SprRubroUnidades(n).MaxRows, Trim(RsDataLoc!cod_loc)
                        SprRubroPromedio(n).SetText 1, SprRubroPromedio(n).MaxRows, Trim(RsDataLoc!cod_loc)
                     Next n
               End Select
               
            Case DSLoc(4).Dep
         
                For n = 0 To 7
                    SprRubroImporte(n).MaxRows = SprRubroImporte(n).MaxRows + 1
                    SprRubroUnidades(n).MaxRows = SprRubroUnidades(n).MaxRows + 1
                    SprRubroPromedio(n).MaxRows = SprRubroPromedio(n).MaxRows + 1
                    SprRubroImporte(n).SetText 1, SprRubroImporte(n).MaxRows, Trim(RsDataLoc!cod_loc)
                    SprRubroUnidades(n).SetText 1, SprRubroUnidades(n).MaxRows, Trim(RsDataLoc!cod_loc)
                    SprRubroPromedio(n).SetText 1, SprRubroPromedio(n).MaxRows, Trim(RsDataLoc!cod_loc)
                Next n
          
          End Select
          
          RsDataLoc.MoveNext
          
          If RsDataLoc.EOF Then
            Exit Do
          End If
          
          If (depn <> RsDataLoc!cod_depn & " " & RsDataLoc!Cod_Sdep And RsDataLoc!cod_depn <> "INT") Or _
             (depn <> RsDataLoc!cod_depn And RsDataLoc!cod_depn = "INT") Then
            For n = 0 To 7
              SprRubroImporte(n).MaxRows = SprRubroImporte(n).MaxRows + 1
              SprRubroUnidades(n).MaxRows = SprRubroUnidades(n).MaxRows + 1
              SprRubroPromedio(n).MaxRows = SprRubroPromedio(n).MaxRows + 1
              SprRubroImporte(n).SetText 1, SprRubroImporte(n).MaxRows, Right(depn, 4)
              SprRubroUnidades(n).SetText 1, SprRubroUnidades(n).MaxRows, Right(depn, 4)
              SprRubroPromedio(n).SetText 1, SprRubroPromedio(n).MaxRows, Right(depn, 4)
              Spread_PintaLinea SprRubroImporte(n), SprRubroImporte(n).MaxRows, 1, SprRubroImporte(n).MaxRows, 4
              Spread_PintaLinea SprRubroUnidades(n), SprRubroUnidades(n).MaxRows, 1, SprRubroUnidades(n).MaxRows, 4
              Spread_PintaLinea SprRubroPromedio(n), SprRubroPromedio(n).MaxRows, 1, SprRubroPromedio(n).MaxRows, 4
            Next
            If RsDataLoc!cod_depn = DSLoc(4).Dep Then
                depn = RsDataLoc!cod_depn '& "    "
            Else
                depn = RsDataLoc!cod_depn & " " & RsDataLoc!Cod_Sdep
            End If
          
        End If
          
          
   Loop
   For n = 0 To 7
     SprRubroImporte(n).MaxRows = SprRubroImporte(n).MaxRows + 1
     SprRubroUnidades(n).MaxRows = SprRubroUnidades(n).MaxRows + 1
     SprRubroPromedio(n).MaxRows = SprRubroPromedio(n).MaxRows + 1
     SprRubroImporte(n).SetText 1, SprRubroImporte(n).MaxRows, " INT"
     SprRubroUnidades(n).SetText 1, SprRubroUnidades(n).MaxRows, " INT"
     SprRubroPromedio(n).SetText 1, SprRubroPromedio(n).MaxRows, " INT"
     Spread_PintaLinea SprRubroImporte(n), SprRubroImporte(n).MaxRows, 1, SprRubroImporte(n).MaxRows, 4
     Spread_PintaLinea SprRubroUnidades(n), SprRubroUnidades(n).MaxRows, 1, SprRubroUnidades(n).MaxRows, 4
     Spread_PintaLinea SprRubroPromedio(n), SprRubroPromedio(n).MaxRows, 1, SprRubroPromedio(n).MaxRows, 4
   Next n
   
   End If
End If

End Sub

Public Sub L_PonerDatosAnterior(spr As control, Spr1 As control, spr2 As control)

Dim rubro As String
Dim SubDepn As String
Dim Fila_Desde As Integer
Dim Imp_Total As Double
Dim Uni_Total As Double
Dim linea As Integer
Dim locales As Variant

rubro = RsDataAnt!cod_rubr
SubDepn = RsDataAnt!Cod_Sdep
linea = 1

Do While Not RsDataAnt.EOF And (RsDataAnt!cod_rubr = rubro)
   spr.GetText 1, linea, locales
   If Trim(locales) = RsDataAnt!cod_local Then
      spr.SetText 3, linea, str(RsDataAnt!Importe)
      Imp_Total = Imp_Total + RsDataAnt!Importe
      Spr1.SetText 3, linea, str(RsDataAnt!unidades)
      Uni_Total = Uni_Total + RsDataAnt!unidades
      spr2.SetText 3, linea, str(L_Division(RsDataAnt!Importe, RsDataAnt!unidades))
      RsDataAnt.MoveNext
      If RsDataAnt.EOF Then
        Exit Do
      End If
    Else
        Select Case Trim(locales)
          Case DSLoc(1).Dep  'Aeroparque
            spr.SetText 3, linea, str(Imp_Total)
            Spr1.SetText 3, linea, str(Uni_Total)
            spr2.SetText 3, linea, str(L_Division(Imp_Total, Uni_Total))
            Imp_Total = 0
            Uni_Total = 0
          Case DSLoc(4).Dep  'Interior
            spr.SetText 3, linea, str(Imp_Total)
            Spr1.SetText 3, linea, str(Uni_Total)
            spr2.SetText 3, linea, str(L_Division(Imp_Total, Uni_Total))
            Imp_Total = 0
            Uni_Total = 0
          Case Else  'Entra por Ezeiza
             Select Case Trim(locales)
                Case DSLoc(2).Sdep
                    spr.SetText 3, linea, str(Imp_Total)
                    Spr1.SetText 3, linea, str(Uni_Total)
                    spr2.SetText 3, linea, str(L_Division(Imp_Total, Uni_Total))
                    Imp_Total = 0
                    Uni_Total = 0
                Case DSLoc(3).Sdep
                    spr.SetText 3, linea, str(Imp_Total)
                    Spr1.SetText 3, linea, str(Uni_Total)
                    spr2.SetText 3, linea, str(L_Division(Imp_Total, Uni_Total))
                    Imp_Total = 0
                    Uni_Total = 0
             End Select
        End Select
   End If
   
   linea = linea + 1

Loop

'Para que ponga el total de la última Dependencia
spr.SetText 3, spr.MaxRows, str(Imp_Total)
Spr1.SetText 3, Spr1.MaxRows, str(Uni_Total)
spr2.SetText 3, spr2.MaxRows, str(L_Division(Imp_Total, Uni_Total))
 
End Sub

Public Function L_Division(divisor As Double, dividendo As Double) As Double

If dividendo = 0 Or divisor = 0 Then
   L_Division = 0
 Else
   L_Division = divisor / dividendo
End If

End Function
Public Sub L_PonerDatosActual(spr As control, Spr1 As control, spr2 As control)
Dim rubro As String
Dim SubDepn As String
Dim Fila_Desde As Integer
Dim Imp_Total As Double
Dim Uni_Total As Double
Dim linea As Integer
Dim locales As Variant

rubro = RsData!cod_rubr
SubDepn = RsData!Cod_Sdep
linea = 1

Do While Not RsData.EOF And (RsData!cod_rubr = rubro)
   spr.GetText 1, linea, locales
   If Trim(locales) = RsData!cod_local Then
      spr.SetText 2, linea, str(RsData!Importe)
      Imp_Total = Imp_Total + RsData!Importe
      Spr1.SetText 2, linea, str(RsData!unidades)
      Uni_Total = Uni_Total + RsData!unidades
      spr2.SetText 2, linea, str(L_Division(RsData!Importe, RsData!unidades))
      RsData.MoveNext
      If RsData.EOF Then
         Exit Do
      End If
   Else
        Select Case Trim(locales)
          Case DSLoc(1).Dep  'Aeroparque
            spr.SetText 2, linea, str(Imp_Total)
            Spr1.SetText 2, linea, str(Uni_Total)
            spr2.SetText 2, linea, str(L_Division(Imp_Total, Uni_Total))
            Imp_Total = 0
            Uni_Total = 0
          Case DSLoc(4).Dep  'Interior
            spr.SetText 2, linea, str(Imp_Total)
            Spr1.SetText 2, linea, str(Uni_Total)
            spr2.SetText 2, linea, str(L_Division(Imp_Total, Uni_Total))
            Imp_Total = 0
            Uni_Total = 0
          Case Else  'Entra por Ezeiza
             Select Case Trim(locales)
                Case DSLoc(2).Sdep
                    spr.SetText 2, linea, str(Imp_Total)
                    Spr1.SetText 2, linea, str(Uni_Total)
                    spr2.SetText 2, linea, str(L_Division(Imp_Total, Uni_Total))
                    Imp_Total = 0
                    Uni_Total = 0
                Case DSLoc(3).Sdep
                    spr.SetText 2, linea, str(Imp_Total)
                    Spr1.SetText 2, linea, str(Uni_Total)
                    spr2.SetText 2, linea, str(L_Division(Imp_Total, Uni_Total))
                    Imp_Total = 0
                    Uni_Total = 0
             End Select
        End Select
   End If
   
   linea = linea + 1

Loop
 
'Para que ponga el total de la última Dependencia
spr.SetText 2, spr.MaxRows, str(Imp_Total)
Spr1.SetText 2, Spr1.MaxRows, str(Uni_Total)
spr2.SetText 2, spr2.MaxRows, str(L_Division(Imp_Total, Uni_Total))
 
End Sub

Private Sub L_Sumarizar(spr As control, Row_D As Integer, Row_H As Integer, Col_D As Integer, Col_H As Integer)

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

'On Error GoTo ErrRubLoc:

frmVtaRubroLocal.caption = Aplicacion.SeteoProceso(frmVtaRubroLocal.caption)

'Consulta Actual
sql = " Select cod_rubr,cod_depn,cod_sdep, decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local) cod_local,"
sql = sql & " sum(importe) importe,sum(unidades)unidades "
sql = sql & " From estadis.view_venta_rubro_local "
sql = sql & L_Armarcondicion(1)
sql = sql & " group by cod_rubr,cod_depn,cod_sdep, decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local)  "
sql = sql & " order by cod_rubr,cod_depn,cod_sdep,cod_local "


'Consulta Anterior
sqlX = " Select cod_rubr,cod_depn,cod_sdep, decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local) cod_local ,"
sqlX = sqlX & " sum(importe) importe,sum(unidades)unidades "
sqlX = sqlX & " From estadis.view_venta_rubro_local "
sqlX = sqlX & L_Armarcondicion(2)
sqlX = sqlX & " group by cod_rubr,cod_depn,cod_sdep, decode(cod_local,'L21', decode(cod_sloc,0,'L05','L06'),'L07','L22',cod_local)  "
sqlX = sqlX & " order by cod_rubr,cod_depn,cod_sdep,cod_local "

If Aplicacion.ObtenerRsDAO(sql, RsData) Then
   
    If Aplicacion.CantReg(RsData) > 0 Then
        L_LlenarLocales
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabRubro.Enabled = True
        
        L_DecoEspigon
        
        If Aplicacion.ObtenerRsDAO(sqlX, RsDataAnt) Then
           L_DecoEspigonAnt
        End If
        
        L_Resaltar
        
        Aplicacion.CerrarDAO RsDataAnt
    
    End If
    
    Aplicacion.CerrarDAO RsData

End If

ErrRubLoc:
    frmVtaRubroLocal.caption = Aplicacion.SeteoFin
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
           L_PonerDatosActual SprRubroImporte(x), SprRubroUnidades(x), SprRubroPromedio(x)
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

Do While Not RsDataAnt.EOF
   
   rubro = RsDataAnt!cod_rubr
   x = L_DecoRubroAnt
   Do While RsDataAnt!cod_rubr = rubro
        If IsNumeric(x) Then
           L_PonerDatosAnterior SprRubroImporte(x), SprRubroUnidades(x), SprRubroPromedio(x)
        Else
           RsDataAnt.MoveNext
        End If
        If RsDataAnt.EOF Then
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

Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabRubro

End Sub

Private Sub Form_Load()
Dim i

Width = Screen.Width * 0.7
Height = Screen.Height * 0.7
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




