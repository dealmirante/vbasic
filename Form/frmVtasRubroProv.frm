VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.0#0"; "crystl32.ocx"
Begin VB.Form frmVtaRubroProv 
   Caption         =   "Ventas por Rubro/Local"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   10965
   Begin VB.Frame Frame2 
      Height          =   1125
      Left            =   8295
      TabIndex        =   37
      Top             =   -90
      Width           =   2130
      Begin VB.OptionButton optSort 
         Caption         =   "Por Imp. Año"
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
         Index           =   2
         Left            =   165
         TabIndex        =   40
         Top             =   645
         Width           =   1530
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Por Nombre"
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
         Index           =   0
         Left            =   165
         TabIndex        =   39
         Top             =   165
         Value           =   -1  'True
         Width           =   1290
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Por Imp. Mes"
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
         Index           =   1
         Left            =   165
         TabIndex        =   38
         Top             =   405
         Width           =   1530
      End
   End
   Begin Crystal.CrystalReport rptform 
      Left            =   9210
      Top             =   3255
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327682
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab tabRubro 
      Height          =   5295
      Left            =   135
      TabIndex        =   2
      Top             =   1065
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   9340
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
      TabPicture(0)   =   "frmVtasRubroProv.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tabsubrubro(0)"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "BEB"
      TabPicture(1)   =   "frmVtasRubroProv.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabsubrubro(1)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "CIG"
      TabPicture(2)   =   "frmVtasRubroProv.frx":0038
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabsubrubro(2)"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "COM"
      TabPicture(3)   =   "frmVtasRubroProv.frx":0054
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tabsubrubro(3)"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "COS"
      TabPicture(4)   =   "frmVtasRubroProv.frx":0070
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "tabsubrubro(4)"
      Tab(4).Control(0).Enabled=   0   'False
      TabCaption(5)   =   "ELE"
      TabPicture(5)   =   "frmVtasRubroProv.frx":008C
      Tab(5).ControlCount=   1
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "tabsubrubro(5)"
      Tab(5).Control(0).Enabled=   0   'False
      TabCaption(6)   =   "PER"
      TabPicture(6)   =   "frmVtasRubroProv.frx":00A8
      Tab(6).ControlCount=   1
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "tabsubrubro(6)"
      Tab(6).Control(0).Enabled=   0   'False
      TabCaption(7)   =   "TAB"
      TabPicture(7)   =   "frmVtasRubroProv.frx":00C4
      Tab(7).ControlCount=   1
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "tabsubrubro(7)"
      Tab(7).Control(0).Enabled=   0   'False
      Begin TabDlg.SSTab tabsubrubro 
         Height          =   4485
         Index           =   4
         Left            =   -74880
         TabIndex        =   35
         Top             =   660
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   7911
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroProv.frx":00E0
         Tab(0).ControlCount=   3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(6)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(7)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SprRubroImporte(4)"
         Tab(0).Control(2).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroProv.frx":00FC
         Tab(1).ControlCount=   3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2(24)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2(25)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "SprRubroUnidades(4)"
         Tab(1).Control(2).Enabled=   0   'False
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "frmVtasRubroProv.frx":0118
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(4)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   3495
            Index           =   4
            Left            =   -74865
            OleObjectBlob   =   "frmVtasRubroProv.frx":0134
            TabIndex        =   74
            Top             =   465
            Width           =   10110
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   3495
            Index           =   4
            Left            =   105
            OleObjectBlob   =   "frmVtasRubroProv.frx":05FD
            TabIndex        =   53
            Top             =   465
            Width           =   10170
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   4
            Left            =   -74670
            OleObjectBlob   =   "frmVtasRubroProv.frx":0AC4
            TabIndex        =   36
            Top             =   180
            Width           =   5505
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   25
            Left            =   -67935
            TabIndex        =   76
            Top             =   0
            Width           =   3045
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   24
            Left            =   -72180
            TabIndex        =   75
            Top             =   0
            Width           =   4260
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   7
            Left            =   2895
            TabIndex        =   55
            Top             =   0
            Width           =   4245
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   6
            Left            =   7125
            TabIndex        =   54
            Top             =   0
            Width           =   2985
         End
      End
      Begin TabDlg.SSTab tabsubrubro 
         Height          =   4500
         Index           =   3
         Left            =   -74895
         TabIndex        =   33
         Top             =   675
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   7938
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         BackColor       =   16777215
         ForeColor       =   16711680
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmVtasRubroProv.frx":0E37
         Tab(0).ControlCount=   3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(14)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(15)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SprRubroImporte(3)"
         Tab(0).Control(2).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroProv.frx":0E53
         Tab(1).ControlCount=   3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2(22)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2(23)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "SprRubroUnidades(3)"
         Tab(1).Control(2).Enabled=   0   'False
         TabCaption(2)   =   "Tab 2"
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(3)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   3495
            Index           =   3
            Left            =   -74895
            OleObjectBlob   =   "frmVtasRubroProv.frx":0E6F
            TabIndex        =   71
            Top             =   465
            Width           =   10155
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   3495
            Index           =   3
            Left            =   120
            OleObjectBlob   =   "frmVtasRubroProv.frx":1338
            TabIndex        =   50
            Top             =   465
            Width           =   10155
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   3
            Left            =   -74190
            OleObjectBlob   =   "frmVtasRubroProv.frx":17FF
            TabIndex        =   34
            Top             =   270
            Width           =   5505
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   23
            Left            =   -67950
            TabIndex        =   73
            Top             =   0
            Width           =   3045
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   22
            Left            =   -72195
            TabIndex        =   72
            Top             =   0
            Width           =   4260
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   15
            Left            =   2895
            TabIndex        =   52
            Top             =   0
            Width           =   4245
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   14
            Left            =   7125
            TabIndex        =   51
            Top             =   0
            Width           =   2985
         End
      End
      Begin TabDlg.SSTab tabsubrubro 
         Height          =   4500
         Index           =   1
         Left            =   -74895
         TabIndex        =   10
         Top             =   675
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   7938
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroProv.frx":1B72
         Tab(0).ControlCount=   3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(2)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(3)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SprRubroImporte(1)"
         Tab(0).Control(2).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroProv.frx":1B8E
         Tab(1).ControlCount=   3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2(18)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2(19)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "SprRubroUnidades(1)"
         Tab(1).Control(2).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroProv.frx":1BAA
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(1)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   3495
            Index           =   1
            Left            =   -74895
            OleObjectBlob   =   "frmVtasRubroProv.frx":1BC6
            TabIndex        =   65
            Top             =   465
            Width           =   10170
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   3495
            Index           =   1
            Left            =   90
            OleObjectBlob   =   "frmVtasRubroProv.frx":208F
            TabIndex        =   44
            Top             =   465
            Width           =   10170
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   1
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroProv.frx":2556
            TabIndex        =   15
            Top             =   105
            Width           =   5505
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   19
            Left            =   -67950
            TabIndex        =   67
            Top             =   0
            Width           =   3045
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   18
            Left            =   -72195
            TabIndex        =   66
            Top             =   0
            Width           =   4260
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   3
            Left            =   2880
            TabIndex        =   46
            Top             =   0
            Width           =   4245
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   2
            Left            =   7110
            TabIndex        =   45
            Top             =   0
            Width           =   2985
         End
      End
      Begin TabDlg.SSTab tabsubrubro 
         Height          =   4500
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   675
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   7938
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroProv.frx":28C9
         Tab(0).ControlCount=   3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SprRubroImporte(0)"
         Tab(0).Control(2).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroProv.frx":28E5
         Tab(1).ControlCount=   3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2(16)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2(17)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "SprRubroUnidades(0)"
         Tab(1).Control(2).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroProv.frx":2901
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(0)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   0
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroProv.frx":291D
            TabIndex        =   9
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   3495
            Index           =   0
            Left            =   90
            OleObjectBlob   =   "frmVtasRubroProv.frx":2C90
            TabIndex        =   8
            Top             =   465
            Width           =   10125
         End
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   3495
            Index           =   0
            Left            =   -74910
            OleObjectBlob   =   "frmVtasRubroProv.frx":3157
            TabIndex        =   41
            Top             =   465
            Width           =   10125
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   17
            Left            =   -72210
            TabIndex        =   43
            Top             =   0
            Width           =   4260
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   16
            Left            =   -67965
            TabIndex        =   42
            Top             =   0
            Width           =   3045
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   1
            Left            =   7110
            TabIndex        =   32
            Top             =   0
            Width           =   2985
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   0
            Left            =   2880
            TabIndex        =   31
            Top             =   0
            Width           =   4245
         End
      End
      Begin TabDlg.SSTab tabsubrubro 
         Height          =   4500
         Index           =   2
         Left            =   -74895
         TabIndex        =   11
         Top             =   675
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   7938
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroProv.frx":3620
         Tab(0).ControlCount=   3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(4)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(5)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SprRubroImporte(2)"
         Tab(0).Control(2).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroProv.frx":363C
         Tab(1).ControlCount=   3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2(20)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2(21)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "SprRubroUnidades(2)"
         Tab(1).Control(2).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroProv.frx":3658
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   3495
            Index           =   2
            Left            =   -74895
            OleObjectBlob   =   "frmVtasRubroProv.frx":3674
            TabIndex        =   68
            Top             =   465
            Width           =   10200
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   3495
            Index           =   2
            Left            =   120
            OleObjectBlob   =   "frmVtasRubroProv.frx":3B3D
            TabIndex        =   47
            Top             =   465
            Width           =   10170
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   2
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroProv.frx":4004
            TabIndex        =   16
            Top             =   105
            Width           =   5505
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   21
            Left            =   -67950
            TabIndex        =   70
            Top             =   0
            Width           =   3045
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   20
            Left            =   -72195
            TabIndex        =   69
            Top             =   0
            Width           =   4260
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   5
            Left            =   2895
            TabIndex        =   49
            Top             =   0
            Width           =   4245
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   4
            Left            =   7125
            TabIndex        =   48
            Top             =   0
            Width           =   2985
         End
      End
      Begin TabDlg.SSTab tabsubrubro 
         Height          =   4500
         Index           =   5
         Left            =   -74895
         TabIndex        =   12
         Top             =   675
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   7938
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroProv.frx":4377
         Tab(0).ControlCount=   3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(12)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(13)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SprRubroImporte(5)"
         Tab(0).Control(2).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroProv.frx":4393
         Tab(1).ControlCount=   3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2(26)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2(27)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "SprRubroUnidades(5)"
         Tab(1).Control(2).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroProv.frx":43AF
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(5)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   3495
            Index           =   5
            Left            =   -74880
            OleObjectBlob   =   "frmVtasRubroProv.frx":43CB
            TabIndex        =   77
            Top             =   465
            Width           =   10170
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   3495
            Index           =   5
            Left            =   120
            OleObjectBlob   =   "frmVtasRubroProv.frx":4894
            TabIndex        =   56
            Top             =   465
            Width           =   10200
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   5
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroProv.frx":4D5B
            TabIndex        =   17
            Top             =   105
            Width           =   5505
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   27
            Left            =   -67950
            TabIndex        =   79
            Top             =   0
            Width           =   3045
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   26
            Left            =   -72195
            TabIndex        =   78
            Top             =   0
            Width           =   4260
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   13
            Left            =   2910
            TabIndex        =   58
            Top             =   0
            Width           =   4245
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   12
            Left            =   7140
            TabIndex        =   57
            Top             =   0
            Width           =   2985
         End
      End
      Begin TabDlg.SSTab tabsubrubro 
         Height          =   4500
         Index           =   6
         Left            =   -74895
         TabIndex        =   13
         Top             =   675
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   7938
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroProv.frx":50CE
         Tab(0).ControlCount=   3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(10)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(11)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SprRubroImporte(6)"
         Tab(0).Control(2).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroProv.frx":50EA
         Tab(1).ControlCount=   3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2(28)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2(29)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "SprRubroUnidades(6)"
         Tab(1).Control(2).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroProv.frx":5106
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(6)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   3495
            Index           =   6
            Left            =   -74880
            OleObjectBlob   =   "frmVtasRubroProv.frx":5122
            TabIndex        =   80
            Top             =   465
            Width           =   10125
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   3495
            Index           =   6
            Left            =   105
            OleObjectBlob   =   "frmVtasRubroProv.frx":55EB
            TabIndex        =   59
            Top             =   465
            Width           =   10140
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   6
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroProv.frx":5AB2
            TabIndex        =   18
            Top             =   105
            Width           =   5505
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   29
            Left            =   -67950
            TabIndex        =   82
            Top             =   0
            Width           =   3045
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   28
            Left            =   -72195
            TabIndex        =   81
            Top             =   0
            Width           =   4260
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   11
            Left            =   2895
            TabIndex        =   61
            Top             =   0
            Width           =   4245
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   10
            Left            =   7125
            TabIndex        =   60
            Top             =   0
            Width           =   2985
         End
      End
      Begin TabDlg.SSTab tabsubrubro 
         Height          =   4500
         Index           =   7
         Left            =   -74895
         TabIndex        =   14
         Top             =   675
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   7938
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasRubroProv.frx":5E25
         Tab(0).ControlCount=   3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(8)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(9)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SprRubroImporte(7)"
         Tab(0).Control(2).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasRubroProv.frx":5E41
         Tab(1).ControlCount=   3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2(30)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2(31)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "SprRubroUnidades(7)"
         Tab(1).Control(2).Enabled=   0   'False
         TabCaption(2)   =   "Importe / Unidades"
         TabPicture(2)   =   "frmVtasRubroProv.frx":5E5D
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprRubroPromedio(7)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprRubroUnidades 
            Height          =   3495
            Index           =   7
            Left            =   -74850
            OleObjectBlob   =   "frmVtasRubroProv.frx":5E79
            TabIndex        =   83
            Top             =   465
            Width           =   10140
         End
         Begin FPSpread.vaSpread SprRubroImporte 
            Height          =   3495
            Index           =   7
            Left            =   120
            OleObjectBlob   =   "frmVtasRubroProv.frx":6342
            TabIndex        =   62
            Top             =   465
            Width           =   10185
         End
         Begin FPSpread.vaSpread SprRubroPromedio 
            Height          =   2400
            Index           =   7
            Left            =   -74505
            OleObjectBlob   =   "frmVtasRubroProv.frx":6809
            TabIndex        =   19
            Top             =   105
            Width           =   5505
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   31
            Left            =   -67905
            TabIndex        =   85
            Top             =   0
            Width           =   3045
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   30
            Left            =   -72150
            TabIndex        =   84
            Top             =   0
            Width           =   4260
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   9
            Left            =   2895
            TabIndex        =   64
            Top             =   0
            Width           =   4245
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Información Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Index           =   8
            Left            =   7125
            TabIndex        =   63
            Top             =   0
            Width           =   2985
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   885
      Left            =   5520
      TabIndex        =   1
      Top             =   60
      Width           =   2520
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   540
         Left            =   60
         Picture         =   "frmVtasRubroProv.frx":6B7C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   225
         Width           =   585
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
         Height          =   555
         Index           =   0
         Left            =   705
         Picture         =   "frmVtasRubroProv.frx":710E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
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
         Height          =   555
         Index           =   1
         Left            =   1305
         Picture         =   "frmVtasRubroProv.frx":7210
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   195
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
         Height          =   570
         Index           =   2
         Left            =   1890
         Picture         =   "frmVtasRubroProv.frx":7312
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   570
      End
   End
   Begin VB.Frame frdatos 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   8250
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
         Height          =   885
         Left            =   210
         TabIndex        =   21
         Top             =   105
         Width           =   5235
         Begin VB.CommandButton botHelpFD 
            Height          =   345
            Left            =   5070
            Picture         =   "frmVtasRubroProv.frx":7B34
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   -75
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton botHelpFH 
            Height          =   345
            Left            =   5055
            Picture         =   "frmVtasRubroProv.frx":7CA6
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   60
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAnioMes 
            Height          =   315
            Left            =   3915
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   705
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            ItemData        =   "frmVtasRubroProv.frx":7E18
            Left            =   1530
            List            =   "frmVtasRubroProv.frx":7E40
            TabIndex        =   23
            Text            =   "cboMes"
            Top             =   345
            Width           =   930
         End
         Begin VB.ComboBox cboAnio 
            Height          =   315
            ItemData        =   "frmVtasRubroProv.frx":7E74
            Left            =   4080
            List            =   "frmVtasRubroProv.frx":7E8D
            TabIndex        =   22
            Text            =   "Combo1"
            Top             =   360
            Width           =   900
         End
         Begin MSMask.MaskEdBox mskFDesde 
            Height          =   285
            Left            =   3900
            TabIndex        =   27
            Top             =   -60
            Visible         =   0   'False
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
            Left            =   3915
            TabIndex        =   28
            Top             =   60
            Visible         =   0   'False
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
            Index           =   0
            Left            =   255
            TabIndex        =   30
            Top             =   360
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
            Left            =   2730
            TabIndex        =   29
            Top             =   345
            Width           =   1185
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
      TabIndex        =   6
      Top             =   5700
      Width           =   8730
   End
End
Attribute VB_Name = "frmVtaRubroProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset

Dim TotCia_MES As Single
Dim ProyCia_MES As Single
Dim TotCia_ANO As Single

Dim CantTotCia_MES As Single
Dim CantTotCia_ANO As Single

Dim cadenaRubro As String

Private Sub L_ObtenerTotales()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset

sql = " select 'A' tipo,sum(importe) Imp,sum(cantidad) Cant "
sql = sql & " , round(sum(importe)/to_char(sysdate-1,'dd')*to_char(last_day(sysdate-1),'dd')) PY"
sql = sql & " , round(sum(cantidad)/to_char(sysdate-1,'dd')*to_char(last_day(sysdate-1),'dd')) PY_un"
sql = sql & " From baires.venta_mes v "
sql = sql & " Where aniomes = " & txtAnioMes.Text
sql = sql & " Union"
sql = sql & " select 'T' tipo,sum(importe) Imp,sum(cantidad) Cant , 0, 0 "
sql = sql & " From baires.venta_mes v "
sql = sql & " where aniomes between " & cboAnio.Text & "01 And " & txtAnioMes.Text

TotCia_MES = 0
TotCia_ANO = 0
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        If rs!tipo = "A" Then
            TotCia_MES = rs!imp
            ProyCia_MES = rs!py
            CantTotCia_MES = rs!cant
        Else
            TotCia_ANO = rs!imp
            CantTotCia_ANO = rs!cant
        End If
    rs.MoveNext
    Loop
    rs.Close
End If
End Sub

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

rptform.Destination = 0     ' 0 pantalla
                            '1 impresora

rptform.Action = True
 DoEvents


Exit Sub
ErrPrint:
'     frmVtaRubroLocal.caption = Aplicacion.SeteoFin
     Exit Sub


End Sub

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



Private Sub L_LimpiarGrillas()
Dim i
labInfo.caption = ""

For i = 0 To 7
  SprRubroImporte(i).MaxRows = 0
  SprRubroUnidades(i).MaxRows = 0
  SprRubroPromedio(i).MaxRows = 0
Next i


End Sub

Public Sub L_PonerDatosActual(spr As control, Spr1 As control, spr2 As control)
Dim rubro As String
Dim Imp_TotalMES As Single, Imp_TotalANO As Single
Dim Cant_TotalMES As Single, Cant_TotalANO As Single
Dim Py_totalMes As Single, PyCant_totalMes As Single
Dim i As Integer
Dim valorMes As Variant, valorAno As Variant


rubro = RsData!cod_rubr


Do While Not RsData.EOF And (RsData!cod_rubr = rubro)
      spr.MaxRows = spr.MaxRows + 1
      spr.SetText 2, spr.MaxRows, Trim(RsData!Descrip)
      
      spr.SetText 3, spr.MaxRows, str(RsData!impmes)
      Imp_TotalMES = Imp_TotalMES + RsData!impmes

      spr.SetText 6, spr.MaxRows, str(RsData!py)
      Py_totalMes = Py_totalMes + RsData!py

      spr.SetText 7, spr.MaxRows, str(RsData!impanual)
      Imp_TotalANO = Imp_TotalANO + RsData!impanual

      Spr1.MaxRows = Spr1.MaxRows + 1
      Spr1.SetText 2, Spr1.MaxRows, Trim(RsData!Descrip)
      
      Spr1.SetText 3, Spr1.MaxRows, str(RsData!cantmes)
      Cant_TotalMES = Cant_TotalMES + RsData!cantmes

      Spr1.SetText 6, Spr1.MaxRows, str(RsData!py_un)
      PyCant_totalMes = PyCant_totalMes + RsData!py_un

      Spr1.SetText 7, Spr1.MaxRows, str(RsData!cantanual)
      Cant_TotalANO = Cant_TotalANO + RsData!cantanual

      RsData.MoveNext
      If RsData.EOF Then
         Exit Do
      End If
Loop
 
'Para que ponga el total de la última Dependencia
spr.MaxRows = spr.MaxRows + 1
Spread.Spread_TotalesLinea spr
spr.SetText 3, spr.MaxRows, str(Imp_TotalMES)
spr.SetText 7, spr.MaxRows, str(Imp_TotalANO)
spr.SetText 6, spr.MaxRows, str(Py_totalMes)
 
Spr1.MaxRows = Spr1.MaxRows + 1
Spread.Spread_TotalesLinea Spr1
Spr1.SetText 3, Spr1.MaxRows, str(Cant_TotalMES)
Spr1.SetText 6, Spr1.MaxRows, str(PyCant_totalMes)
Spr1.SetText 7, Spr1.MaxRows, str(Cant_TotalANO)
'Participacion

For i = 1 To spr.MaxRows
    spr.GetText 3, i, valorMes
    spr.GetText 7, i, valorAno
    
    spr.SetText 4, i, Format(100 * valorMes / Imp_TotalMES, "###.00")
    spr.SetText 8, i, Format(100 * valorAno / Imp_TotalANO, "###.00")
    
    spr.SetText 5, i, Format(100 * valorMes / TotCia_MES, "###.00")
    spr.SetText 9, i, Format(100 * valorAno / TotCia_ANO, "###.00")
    
Next

For i = 1 To spr.MaxRows
    Spr1.GetText 3, i, valorMes
    Spr1.GetText 7, i, valorAno
    
    Spr1.SetText 4, i, Format(100 * valorMes / Cant_TotalMES, "###.00")
    Spr1.SetText 8, i, Format(100 * valorAno / Cant_TotalANO, "###.00")
    
    Spr1.SetText 5, i, Format(100 * valorMes / CantTotCia_MES, "###.00")
    Spr1.SetText 9, i, Format(100 * valorAno / CantTotCia_ANO, "###.00")
    
Next

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

frmVtaRubroProv.caption = Aplicacion.SeteoProceso(frmVtaRubroProv.caption)

'Consulta Actual
sql = " select cod_rubr,cod_prov,descrip"
sql = sql & " ,sum(ImpMes) ImpMes, sum(ImpAnual) ImpAnual"
sql = sql & " ,sum(CantMes) CantMes, sum(CantAnual) CantAnual"
sql = sql & " ,round(sum(ImpMes)/to_char(sysdate-1,'dd')*to_char(last_day(sysdate-1),'dd')) PY"
sql = sql & " ,round(sum(CantMes)/to_char(sysdate-1,'dd')*to_char(last_day(sysdate-1),'dd')) PY_un"
sql = sql & " From ("
sql = sql & " select  cod_rubr,pr.cod_prov,pr.descrip, sum(importe) ImpMes, 0 ImpAnual, sum(cantidad) CantMes, 0  CantAnual "
sql = sql & " from baires.producto p,baires.proveedor pr,baires.venta_mes v "
sql = sql & " Where cod_rubr <> 'REG' and aniomes = " & txtAnioMes.Text
sql = sql & " and p.cod_prod = v.cod_prod and pr.cod_prov = p.cod_prov "
sql = sql & " group by  cod_rubr,pr.cod_prov,pr.descrip "
sql = sql & " Union "
sql = sql & " select  cod_rubr,pr.cod_prov,pr.descrip, 0 ImpMes, sum(importe) ImpAnual, 0 CantMes, sum(cantidad)  CantAnual "
sql = sql & " from baires.producto p,baires.proveedor pr,baires.venta_mes v "
sql = sql & " where cod_rubr <> 'REG' and  aniomes between " & cboAnio.Text & "01 And " & txtAnioMes.Text
sql = sql & " and p.cod_prod = v.cod_prod and pr.cod_prov = p.cod_prov "
sql = sql & " group by  cod_rubr,pr.cod_prov,pr.descrip "
sql = sql & " ) "
sql = sql & " group by  cod_rubr,cod_prov,descrip "
sql = sql & " "


If Aplicacion.ObtenerRsDAO(sql, RsData) Then
   
    If Aplicacion.CantReg(RsData) > 0 Then

        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabRubro.Enabled = True
        
        L_ObtenerTotales
        
        L_DecoEspigon
                
        L_Resaltar
        
    
    End If
    
    Aplicacion.CerrarDAO RsData

End If

ErrRubLoc:
    frmVtaRubroProv.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub




Private Sub L_Resaltar()

Dim fila As Integer
Dim i

For i = 0 To 7
  For fila = 1 To SprRubroImporte(i).MaxRows
     Spread.spread_ResaltarCelda SprRubroImporte(i), 4, fila
     Spread.spread_ResaltarCelda SprRubroImporte(i), 5, fila

     Spread.spread_ResaltarCelda SprRubroImporte(i), 8, fila
     Spread.spread_ResaltarCelda SprRubroImporte(i), 9, fila
  
     Spread.spread_ResaltarCelda SprRubroUnidades(i), 4, fila
     Spread.spread_ResaltarCelda SprRubroUnidades(i), 5, fila

     Spread.spread_ResaltarCelda SprRubroUnidades(i), 8, fila
     Spread.spread_ResaltarCelda SprRubroUnidades(i), 9, fila
  
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


Public Sub Set_cadenaRubro(valor As String)
cadenaRubro = valor
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

Private Sub cboAnio_Click()
If cboAnio.Text = Format(Now, "YYYY") Then
    txtAnioMes.Text = cboAnio.Text & cboMes.Text
Else
    txtAnioMes.Text = cboAnio.Text & "12"
End If
End Sub

Private Sub cboMes_Click()
txtAnioMes.Text = cboAnio.Text & cboMes.Text
End Sub


Private Sub CmdExcel_Click()
Dim AppExcel As Excel.Application
Dim libroExcel As Excel.Workbook
Dim hojaExcel As Excel.Worksheet
Dim titulo As String, subtitulo As String
Dim NOMBRE As String
Dim ind As Integer

cadenaRubro = ""
FrmRubro.Text1 = 3
FrmRubro.Show 1

titulo = "Informe de ventas por Proveedor - Total Cía. "
If cadenaRubro <> "" Then
    
    NOMBRE = frmDir.NombreArchivo()
    DoEvents
    If NOMBRE <> "" Then
        frmVtaRubroProv.caption = Aplicacion.SeteoProceso(frmVtaRubroProv.caption)
        ind = 1
        Set AppExcel = CreateObject("Excel.application")
    
        Set libroExcel = AppExcel.Workbooks.Add
    
        'AppExcel.Visible = True
        
        If InStr(1, cadenaRubro, "ACC") <> 0 Then
            subtitulo = "ACCESORIOS " & cboAnio.Text & " - " & cboMes.Text
            Set hojaExcel = libroExcel.Worksheets(ind)
            hojaExcel.Activate
            hojaExcel.Name = "ACC"
            L_TratarExcel hojaExcel, SprRubroImporte(0), titulo, subtitulo
            ind = ind + 1
        End If
        If InStr(1, cadenaRubro, "BEB") <> 0 Then
            subtitulo = "BEBIDAS " & cboAnio.Text & " - " & cboMes.Text
            Set hojaExcel = libroExcel.Worksheets(ind)
            hojaExcel.Activate
            hojaExcel.Name = "BEB"
            L_TratarExcel hojaExcel, SprRubroImporte(1), titulo, subtitulo
            ind = ind + 1
        End If
        If InStr(1, cadenaRubro, "CIG") <> 0 Then
            subtitulo = "CIGARRILLOS " & cboAnio.Text & " - " & cboMes.Text
            Set hojaExcel = libroExcel.Worksheets(ind)
            hojaExcel.Activate
            hojaExcel.Name = "CIG"
            L_TratarExcel hojaExcel, SprRubroImporte(2), titulo, subtitulo
            ind = ind + 1
        End If
        If InStr(1, cadenaRubro, "COM") <> 0 Then
            subtitulo = "COMESTIBLES " & cboAnio.Text & " - " & cboMes.Text
            If ind > 3 Then
                libroExcel.Sheets.Add
                Set hojaExcel = libroExcel.Worksheets(ind - 1)
                hojaExcel.Move after:=libroExcel.Worksheets(ind)
            Else
                Set hojaExcel = libroExcel.Worksheets(ind)
            End If
            hojaExcel.Activate
            hojaExcel.Name = "COM"
            L_TratarExcel hojaExcel, SprRubroImporte(3), titulo, subtitulo
            ind = ind + 1
        End If
        If InStr(1, cadenaRubro, "COS") <> 0 Then
            subtitulo = "COSMETICOS " & cboAnio.Text & " - " & cboMes.Text
            If ind > 3 Then
                libroExcel.Sheets.Add
                Set hojaExcel = libroExcel.Worksheets(ind - 1)
                hojaExcel.Move after:=libroExcel.Worksheets(ind)
            Else
                Set hojaExcel = libroExcel.Worksheets(ind)
            End If
            hojaExcel.Activate
            hojaExcel.Name = "COS"
            L_TratarExcel hojaExcel, SprRubroImporte(4), titulo, subtitulo
            ind = ind + 1
        End If
        If InStr(1, cadenaRubro, "ELE") <> 0 Then
            subtitulo = "ELECTRONICA " & cboAnio.Text & " - " & cboMes.Text
            If ind > 3 Then
                libroExcel.Sheets.Add
                Set hojaExcel = libroExcel.Worksheets(ind - 1)
                hojaExcel.Move after:=libroExcel.Worksheets(ind)
            Else
                Set hojaExcel = libroExcel.Worksheets(ind)
            End If
            hojaExcel.Activate
            hojaExcel.Name = "ELE"
            L_TratarExcel hojaExcel, SprRubroImporte(5), titulo, subtitulo
            ind = ind + 1
        End If
        If InStr(1, cadenaRubro, "PER") <> 0 Then
            subtitulo = "PERFUMERIA " & cboAnio.Text & " - " & cboMes.Text
            If ind > 3 Then
                libroExcel.Sheets.Add
                Set hojaExcel = libroExcel.Worksheets(ind - 1)
                hojaExcel.Move after:=libroExcel.Worksheets(ind)
            Else
                Set hojaExcel = libroExcel.Worksheets(ind)
            End If
            hojaExcel.Activate
            hojaExcel.Name = "PER"
            L_TratarExcel hojaExcel, SprRubroImporte(6), titulo, subtitulo
            ind = ind + 1
        End If
        If InStr(1, cadenaRubro, "TAB") <> 0 Then
            subtitulo = "TABACOS " & cboAnio.Text & " - " & cboMes.Text
            If ind > 3 Then
                libroExcel.Sheets.Add
                Set hojaExcel = libroExcel.Worksheets(ind - 1)
                hojaExcel.Move after:=libroExcel.Worksheets(ind)
            Else
                Set hojaExcel = libroExcel.Worksheets(ind)
            End If
            hojaExcel.Activate
            hojaExcel.Name = "TAB"
            L_TratarExcel hojaExcel, SprRubroImporte(7), titulo, subtitulo
            ind = ind + 1
        End If
        frmVtaRubroProv.caption = Aplicacion.SeteoFin
        If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
          libroExcel.PrintOut
        End If
        
        libroExcel.SaveAs (NOMBRE & ".xls")
        AppExcel.Quit
        Set AppExcel = Nothing
    
    End If
End If


End Sub


Private Sub L_TratarExcel(hojaExcel As Object, spr As control, titulo As String, subTit As String)

Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim i As Integer
Dim tit As Variant


On Error GoTo ErrorExl:

    
    
    ReDim titCol(spr.MaxCols)
    Col = 1
    fila = 4
    
    For i = 1 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor hojaExcel, 2, 1, titulo
    rango = Exl_rangos(2, 2, 1, spr.MaxCols)
    Exl_Letra hojaExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion hojaExcel, rango, Exl_Centro, Exl_CentroVert, False
    hojaExcel.Application.Range(rango).Merge
    Exl_Lineas hojaExcel, rango, Exl_Linsimple
    Exl_ColorInt hojaExcel, rango, Exl_Gris
    
    Exl_PonerValor hojaExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra hojaExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion hojaExcel, rango, Exl_Izq, Exl_CentroVert, False
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra hojaExcel, rango, NEGRITA, 12, "Ms Serif"
    
    hojaExcel.Application.Range(rango).Merge
    Exl_Lineas hojaExcel, rango, Exl_Linsimple
            
                
        fila = fila + 1
        Exl_BajarGrillaExel spr, hojaExcel, fila, Col, titCol
        rango = Exl_rangos(fila + 1, fila + spr.MaxRows, Col + 1, spr.MaxCols)
        Exl_Format hojaExcel, rango
        
        Exl.Exl_AnchoCol hojaExcel, 1, 1, 20
        Exl.Exl_AnchoCol hojaExcel, 3, 4, 9
        Exl.Exl_AnchoCol hojaExcel, 6, 7, 9
        
        'Elimina las columnas que sobran
        rango = Exl_rangos(1, 1000, spr.MaxCols, spr.MaxCols)
        hojaExcel.Application.Range(rango).Delete

        'rango = Exl_rangos(fila + 3, fila + 4, 3, 3)
        'Exl_ColorInt AppExcel, rango, Exl_Gris
        'rango = Exl_rangos(fila + 3, fila + 4, 5, 5)
        'Exl_ColorInt AppExcel, rango, Exl_Gris
        'rango = Exl_rangos(fila + 4, fila + 4, 4, 4)
        'Exl_ColorInt AppExcel, rango, Exl_Gris
        'rango = Exl_rangos(fila + 4, fila + 4, 6, 6)
        'Exl_ColorInt AppExcel, rango, Exl_Gris
        
        'fila = fila + sprGral(0).MaxRows + 1
        'rango = Exl_rangos(fila, fila, col, sprGral(0).MaxCols)
        'Exl_ColorInt AppExcel, rango, Exl_Gris

    
    'AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    'AppExcel.application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    
    'If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
    '    AppExcel.PrintOut
    'End If
    'AppExcel.SaveAs NOMBRE & ".xls"
    'Set AppExcel = Nothing


ErrorExl:

    frmVtaRubroProv.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub



Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabRubro

End Sub

Private Sub Form_Load()
Dim i

Width = Screen.Width * 0.95
Height = Screen.Height * 0.8
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height - 1000) / 2

    mskFDesde.Text = Format$(Func.func_Dia1SegunMes_Anio(Month(Date - 1), Year(Date - 1)), FTOFECHA)
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
        

cboMes.Text = Format(Now - 1, "MM")
cboAnio.Text = Format(Now - 1, "YYYY")
txtAnioMes.Text = Format(Now - 1, "YYYYMM")
L_LimpiarGrillas

For i = 0 To 7
    'TabSubrubro(i).TabVisible(1) = False
    TabSubrubro(i).TabVisible(2) = False
Next

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

    
    If mskFHasta.Text <> "" Then
        If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
            mskFHasta.Text = mskFDesde.Text
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


End Sub





Private Sub optSort_Click(Index As Integer)
Dim i As Integer

Select Case Index
    Case 0
        For i = 0 To 7
        Spread.spread_OrdenarGrilla SprRubroImporte(i), 2, SS_SORT_ORDER_ASCENDING
        Spread.spread_OrdenarGrilla SprRubroUnidades(i), 2, SS_SORT_ORDER_ASCENDING
        Next
    Case 1
        For i = 0 To 7
        Spread.spread_OrdenarGrilla SprRubroImporte(i), 3, SS_SORT_ORDER_DESCENDING
        Spread.spread_OrdenarGrilla SprRubroUnidades(i), 3, SS_SORT_ORDER_DESCENDING
        Next
    Case 2
        For i = 0 To 7
        Spread.spread_OrdenarGrilla SprRubroImporte(i), 7, SS_SORT_ORDER_DESCENDING
        Spread.spread_OrdenarGrilla SprRubroUnidades(i), 7, SS_SORT_ORDER_DESCENDING
        Next
End Select

End Sub
