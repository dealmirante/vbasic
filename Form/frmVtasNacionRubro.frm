VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.0#0"; "crystl32.ocx"
Begin VB.Form frmVtaNacionRubro 
   Caption         =   "Ventas por Nacionalidad Rubro y Local"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   10965
   Begin VB.ListBox lstNacionSolapa 
      Height          =   2010
      Left            =   10035
      TabIndex        =   35
      Top             =   2760
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.ListBox lstNacion 
      Height          =   1635
      ItemData        =   "frmVtasNacionRubro.frx":0000
      Left            =   8250
      List            =   "frmVtasNacionRubro.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   34
      Top             =   105
      Width           =   2505
   End
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
      TabIndex        =   26
      Top             =   1590
      Width           =   705
   End
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   135
      TabIndex        =   24
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
         Left            =   2370
         TabIndex        =   156
         Top             =   195
         Width           =   705
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
         Left            =   6450
         TabIndex        =   33
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
         Left            =   5475
         TabIndex        =   32
         Top             =   210
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         Height          =   405
         Left            =   7425
         Picture         =   "frmVtasNacionRubro.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Left            =   4695
         TabIndex        =   30
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
         Left            =   3900
         TabIndex        =   29
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
         Left            =   3120
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   195
         Width           =   705
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
         TabIndex        =   25
         Top             =   195
         Width           =   705
      End
   End
   Begin Crystal.CrystalReport rptform 
      Left            =   8280
      Top             =   285
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327682
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab tabRubro 
      Height          =   4065
      Left            =   60
      TabIndex        =   2
      Top             =   1980
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   7170
      _Version        =   327680
      Tabs            =   20
      TabsPerRow      =   8
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
      TabCaption(0)   =   "Argentino"
      TabPicture(0)   =   "frmVtasNacionRubro.frx":0446
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6(0)"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Uruguayo"
      TabPicture(1)   =   "frmVtasNacionRubro.frx":0462
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6(1)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Brasileño"
      TabPicture(2)   =   "frmVtasNacionRubro.frx":047E
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6(2)"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "Chileno"
      TabPicture(3)   =   "frmVtasNacionRubro.frx":049A
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6(3)"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "N. Americano"
      TabPicture(4)   =   "frmVtasNacionRubro.frx":04B6
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6(4)"
      Tab(4).Control(0).Enabled=   0   'False
      TabCaption(5)   =   "O. Americanos"
      TabPicture(5)   =   "frmVtasNacionRubro.frx":04D2
      Tab(5).ControlCount=   1
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6(5)"
      Tab(5).Control(0).Enabled=   0   'False
      TabCaption(6)   =   "O.Europeos"
      TabPicture(6)   =   "frmVtasNacionRubro.frx":04EE
      Tab(6).ControlCount=   1
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame6(6)"
      Tab(6).Control(0).Enabled=   0   'False
      TabCaption(7)   =   "Resto Mundo"
      TabPicture(7)   =   "frmVtasNacionRubro.frx":050A
      Tab(7).ControlCount=   1
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame6(7)"
      Tab(7).Control(0).Enabled=   0   'False
      TabCaption(8)   =   "Aleman"
      TabPicture(8)   =   "frmVtasNacionRubro.frx":0526
      Tab(8).ControlCount=   1
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame6(8)"
      Tab(8).Control(0).Enabled=   0   'False
      TabCaption(9)   =   "Italiano"
      TabPicture(9)   =   "frmVtasNacionRubro.frx":0542
      Tab(9).ControlCount=   1
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame6(9)"
      Tab(9).Control(0).Enabled=   0   'False
      TabCaption(10)  =   "Español"
      TabPicture(10)  =   "frmVtasNacionRubro.frx":055E
      Tab(10).ControlCount=   1
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame6(10)"
      Tab(10).Control(0).Enabled=   0   'False
      TabCaption(11)  =   "Frances"
      TabPicture(11)  =   "frmVtasNacionRubro.frx":057A
      Tab(11).ControlCount=   1
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Frame6(11)"
      Tab(11).Control(0).Enabled=   0   'False
      TabCaption(12)  =   "Ingles"
      TabPicture(12)  =   "frmVtasNacionRubro.frx":0596
      Tab(12).ControlCount=   1
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Frame6(12)"
      Tab(12).Control(0).Enabled=   0   'False
      TabCaption(13)  =   "Mexicano"
      TabPicture(13)  =   "frmVtasNacionRubro.frx":05B2
      Tab(13).ControlCount=   1
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "Frame6(13)"
      Tab(13).Control(0).Enabled=   0   'False
      TabCaption(14)  =   "Colombiano"
      TabPicture(14)  =   "frmVtasNacionRubro.frx":05CE
      Tab(14).ControlCount=   1
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "Frame6(14)"
      Tab(14).Control(0).Enabled=   0   'False
      TabCaption(15)  =   "Venezolano"
      TabPicture(15)  =   "frmVtasNacionRubro.frx":05EA
      Tab(15).ControlCount=   1
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "Frame6(15)"
      Tab(15).Control(0).Enabled=   0   'False
      TabCaption(16)  =   "Peruano"
      TabPicture(16)  =   "frmVtasNacionRubro.frx":0606
      Tab(16).ControlCount=   1
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "Frame6(16)"
      Tab(16).Control(0).Enabled=   0   'False
      TabCaption(17)  =   "Ecuatoriano"
      TabPicture(17)  =   "frmVtasNacionRubro.frx":0622
      Tab(17).ControlCount=   1
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "Frame6(17)"
      Tab(17).Control(0).Enabled=   0   'False
      TabCaption(18)  =   "Paraguayo"
      TabPicture(18)  =   "frmVtasNacionRubro.frx":063E
      Tab(18).ControlCount=   1
      Tab(18).ControlEnabled=   0   'False
      Tab(18).Control(0)=   "Frame6(18)"
      Tab(18).Control(0).Enabled=   0   'False
      TabCaption(19)  =   "Boliviano"
      TabPicture(19)  =   "frmVtasNacionRubro.frx":065A
      Tab(19).ControlCount=   1
      Tab(19).ControlEnabled=   0   'False
      Tab(19).Control(0)=   "Frame6(19)"
      Tab(19).Control(0).Enabled=   0   'False
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   19
         Left            =   -74925
         TabIndex        =   95
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   19
            Left            =   135
            TabIndex        =   107
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":0676
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(19)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":0692
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(19)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":06AE
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(19)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":06CA
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(19)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   19
               Left            =   180
               OleObjectBlob   =   "frmVtasNacionRubro.frx":06E6
               TabIndex        =   119
               Top             =   120
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   19
               Left            =   -74295
               OleObjectBlob   =   "frmVtasNacionRubro.frx":0B4F
               TabIndex        =   131
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   19
               Left            =   -74355
               OleObjectBlob   =   "frmVtasNacionRubro.frx":0EA6
               TabIndex        =   143
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   19
               Left            =   -74820
               OleObjectBlob   =   "frmVtasNacionRubro.frx":1219
               TabIndex        =   155
               Top             =   90
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   18
         Left            =   -74910
         TabIndex        =   94
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   18
            Left            =   135
            TabIndex        =   106
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":1682
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(18)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":169E
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(18)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":16BA
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(18)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":16D6
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(18)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   18
               Left            =   165
               OleObjectBlob   =   "frmVtasNacionRubro.frx":16F2
               TabIndex        =   118
               Top             =   120
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   18
               Left            =   -74280
               OleObjectBlob   =   "frmVtasNacionRubro.frx":1B5B
               TabIndex        =   130
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   18
               Left            =   -74295
               OleObjectBlob   =   "frmVtasNacionRubro.frx":1EB2
               TabIndex        =   142
               Top             =   90
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   18
               Left            =   -74835
               OleObjectBlob   =   "frmVtasNacionRubro.frx":2225
               TabIndex        =   154
               Top             =   90
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   17
         Left            =   -74910
         TabIndex        =   93
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   17
            Left            =   150
            TabIndex        =   105
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":268E
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(17)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":26AA
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(17)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":26C6
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(17)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":26E2
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(17)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   17
               Left            =   165
               OleObjectBlob   =   "frmVtasNacionRubro.frx":26FE
               TabIndex        =   117
               Top             =   105
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   17
               Left            =   -74265
               OleObjectBlob   =   "frmVtasNacionRubro.frx":2B67
               TabIndex        =   129
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   17
               Left            =   -74265
               OleObjectBlob   =   "frmVtasNacionRubro.frx":2EBE
               TabIndex        =   141
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   17
               Left            =   -74820
               OleObjectBlob   =   "frmVtasNacionRubro.frx":3231
               TabIndex        =   153
               Top             =   90
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   16
         Left            =   -74940
         TabIndex        =   92
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   16
            Left            =   150
            TabIndex        =   104
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":369A
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(16)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":36B6
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(16)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":36D2
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(16)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":36EE
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(16)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   16
               Left            =   165
               OleObjectBlob   =   "frmVtasNacionRubro.frx":370A
               TabIndex        =   116
               Top             =   105
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   16
               Left            =   -74310
               OleObjectBlob   =   "frmVtasNacionRubro.frx":3B73
               TabIndex        =   128
               Top             =   120
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   16
               Left            =   -74310
               OleObjectBlob   =   "frmVtasNacionRubro.frx":3ECA
               TabIndex        =   140
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   16
               Left            =   -74820
               OleObjectBlob   =   "frmVtasNacionRubro.frx":423D
               TabIndex        =   152
               Top             =   90
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   15
         Left            =   -74925
         TabIndex        =   91
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   15
            Left            =   150
            TabIndex        =   103
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":46A6
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(15)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":46C2
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(15)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":46DE
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(15)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":46FA
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(15)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   15
               Left            =   165
               OleObjectBlob   =   "frmVtasNacionRubro.frx":4716
               TabIndex        =   115
               Top             =   105
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   15
               Left            =   -74295
               OleObjectBlob   =   "frmVtasNacionRubro.frx":4B7F
               TabIndex        =   127
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   15
               Left            =   -74295
               OleObjectBlob   =   "frmVtasNacionRubro.frx":4ED6
               TabIndex        =   139
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   15
               Left            =   -74820
               OleObjectBlob   =   "frmVtasNacionRubro.frx":5249
               TabIndex        =   151
               Top             =   90
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   14
         Left            =   -74910
         TabIndex        =   90
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   14
            Left            =   135
            TabIndex        =   102
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":56B2
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(14)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":56CE
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(14)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":56EA
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(14)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":5706
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(14)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   14
               Left            =   180
               OleObjectBlob   =   "frmVtasNacionRubro.frx":5722
               TabIndex        =   114
               Top             =   90
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   14
               Left            =   -74265
               OleObjectBlob   =   "frmVtasNacionRubro.frx":5B8B
               TabIndex        =   126
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   14
               Left            =   -74280
               OleObjectBlob   =   "frmVtasNacionRubro.frx":5EE2
               TabIndex        =   138
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   14
               Left            =   -74835
               OleObjectBlob   =   "frmVtasNacionRubro.frx":6255
               TabIndex        =   150
               Top             =   75
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   13
         Left            =   -74820
         TabIndex        =   89
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   13
            Left            =   135
            TabIndex        =   101
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":66BE
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(13)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":66DA
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(13)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":66F6
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(13)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":6712
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(13)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   13
               Left            =   165
               OleObjectBlob   =   "frmVtasNacionRubro.frx":672E
               TabIndex        =   113
               Top             =   105
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   13
               Left            =   -74280
               OleObjectBlob   =   "frmVtasNacionRubro.frx":6B97
               TabIndex        =   125
               Top             =   120
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   13
               Left            =   -74280
               OleObjectBlob   =   "frmVtasNacionRubro.frx":6EEE
               TabIndex        =   137
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   13
               Left            =   -74835
               OleObjectBlob   =   "frmVtasNacionRubro.frx":7261
               TabIndex        =   149
               Top             =   90
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   12
         Left            =   -74940
         TabIndex        =   88
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   12
            Left            =   135
            TabIndex        =   100
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":76CA
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(12)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":76E6
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(12)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":7702
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(12)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":771E
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(12)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   12
               Left            =   150
               OleObjectBlob   =   "frmVtasNacionRubro.frx":773A
               TabIndex        =   112
               Top             =   105
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   12
               Left            =   -74295
               OleObjectBlob   =   "frmVtasNacionRubro.frx":7BA3
               TabIndex        =   124
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   12
               Left            =   -74250
               OleObjectBlob   =   "frmVtasNacionRubro.frx":7EFA
               TabIndex        =   136
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   12
               Left            =   -74835
               OleObjectBlob   =   "frmVtasNacionRubro.frx":826D
               TabIndex        =   148
               Top             =   90
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   11
         Left            =   -74925
         TabIndex        =   87
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   11
            Left            =   135
            TabIndex        =   99
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":86D6
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(11)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":86F2
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(11)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":870E
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(11)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":872A
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(11)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   11
               Left            =   165
               OleObjectBlob   =   "frmVtasNacionRubro.frx":8746
               TabIndex        =   111
               Top             =   105
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   11
               Left            =   -74280
               OleObjectBlob   =   "frmVtasNacionRubro.frx":8BAF
               TabIndex        =   123
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   11
               Left            =   -74310
               OleObjectBlob   =   "frmVtasNacionRubro.frx":8F06
               TabIndex        =   135
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   11
               Left            =   -74850
               OleObjectBlob   =   "frmVtasNacionRubro.frx":9279
               TabIndex        =   147
               Top             =   90
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   10
         Left            =   -74820
         TabIndex        =   86
         Top             =   795
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   10
            Left            =   135
            TabIndex        =   98
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":96E2
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(10)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":96FE
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(10)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":971A
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(10)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":9736
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(10)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   10
               Left            =   165
               OleObjectBlob   =   "frmVtasNacionRubro.frx":9752
               TabIndex        =   110
               Top             =   105
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   10
               Left            =   -74295
               OleObjectBlob   =   "frmVtasNacionRubro.frx":9BBB
               TabIndex        =   122
               Top             =   120
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   10
               Left            =   -74280
               OleObjectBlob   =   "frmVtasNacionRubro.frx":9F12
               TabIndex        =   134
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   10
               Left            =   -74865
               OleObjectBlob   =   "frmVtasNacionRubro.frx":A285
               TabIndex        =   146
               Top             =   75
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   9
         Left            =   -74835
         TabIndex        =   85
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   9
            Left            =   135
            TabIndex        =   97
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":A6EE
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(9)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":A70A
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(9)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":A726
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(9)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":A742
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(9)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   9
               Left            =   150
               OleObjectBlob   =   "frmVtasNacionRubro.frx":A75E
               TabIndex        =   109
               Top             =   105
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   9
               Left            =   -74295
               OleObjectBlob   =   "frmVtasNacionRubro.frx":ABC7
               TabIndex        =   121
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   9
               Left            =   -74295
               OleObjectBlob   =   "frmVtasNacionRubro.frx":AF1E
               TabIndex        =   133
               Top             =   90
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   9
               Left            =   -74835
               OleObjectBlob   =   "frmVtasNacionRubro.frx":B291
               TabIndex        =   145
               Top             =   75
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   8
         Left            =   -74805
         TabIndex        =   84
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2910
            Index           =   8
            Left            =   150
            TabIndex        =   96
            Top             =   180
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   5133
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importes"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":B6FA
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(8)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":B716
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(8)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe/Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":B732
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(8)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":B74E
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(8)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   8
               Left            =   150
               OleObjectBlob   =   "frmVtasNacionRubro.frx":B76A
               TabIndex        =   108
               Top             =   105
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   8
               Left            =   -74295
               OleObjectBlob   =   "frmVtasNacionRubro.frx":BBD3
               TabIndex        =   120
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   8
               Left            =   -74310
               OleObjectBlob   =   "frmVtasNacionRubro.frx":BF2A
               TabIndex        =   132
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   8
               Left            =   -74835
               OleObjectBlob   =   "frmVtasNacionRubro.frx":C29D
               TabIndex        =   144
               Top             =   90
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   7
         Left            =   -74880
         TabIndex        =   48
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   3000
            Index           =   7
            Left            =   195
            TabIndex        =   79
            Top             =   135
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
            TabPicture(0)   =   "frmVtasNacionRubro.frx":C706
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(7)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":C722
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(7)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe / Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":C73E
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(7)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":C75A
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(7)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   7
               Left            =   -74925
               OleObjectBlob   =   "frmVtasNacionRubro.frx":C776
               TabIndex        =   80
               Top             =   75
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   7
               Left            =   90
               OleObjectBlob   =   "frmVtasNacionRubro.frx":CBDF
               TabIndex        =   81
               Top             =   90
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   7
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":D048
               TabIndex        =   82
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   7
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":D39F
               TabIndex        =   83
               Top             =   105
               Width           =   5505
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   6
         Left            =   -74865
         TabIndex        =   47
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   3000
            Index           =   6
            Left            =   195
            TabIndex        =   74
            Top             =   135
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
            TabPicture(0)   =   "frmVtasNacionRubro.frx":D712
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(6)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":D72E
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(6)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe / Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":D74A
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(6)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":D766
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(6)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   6
               Left            =   -74925
               OleObjectBlob   =   "frmVtasNacionRubro.frx":D782
               TabIndex        =   78
               Top             =   75
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   6
               Left            =   105
               OleObjectBlob   =   "frmVtasNacionRubro.frx":DBEB
               TabIndex        =   77
               Top             =   75
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   6
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":E054
               TabIndex        =   76
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   6
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":E3AB
               TabIndex        =   75
               Top             =   105
               Width           =   5505
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   5
         Left            =   -74895
         TabIndex        =   46
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   3000
            Index           =   5
            Left            =   180
            TabIndex        =   69
            Top             =   135
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
            TabPicture(0)   =   "frmVtasNacionRubro.frx":E71E
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(5)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":E73A
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(5)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe / Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":E756
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(5)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":E772
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(5)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   5
               Left            =   -74925
               OleObjectBlob   =   "frmVtasNacionRubro.frx":E78E
               TabIndex        =   73
               Top             =   75
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   5
               Left            =   90
               OleObjectBlob   =   "frmVtasNacionRubro.frx":EBF7
               TabIndex        =   72
               Top             =   75
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   5
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":F060
               TabIndex        =   71
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   5
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":F3B7
               TabIndex        =   70
               Top             =   105
               Width           =   5505
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   4
         Left            =   -74850
         TabIndex        =   45
         Top             =   795
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   3000
            Index           =   4
            Left            =   180
            TabIndex        =   64
            Top             =   120
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
            TabPicture(0)   =   "frmVtasNacionRubro.frx":F72A
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(4)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":F746
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(4)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe / Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":F762
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(4)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":F77E
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(4)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   4
               Left            =   -74910
               OleObjectBlob   =   "frmVtasNacionRubro.frx":F79A
               TabIndex        =   68
               Top             =   90
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   4
               Left            =   135
               OleObjectBlob   =   "frmVtasNacionRubro.frx":FC03
               TabIndex        =   67
               Top             =   90
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   4
               Left            =   -74490
               OleObjectBlob   =   "frmVtasNacionRubro.frx":1006C
               TabIndex        =   66
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   4
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":103C3
               TabIndex        =   65
               Top             =   105
               Width           =   5505
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   3
         Left            =   -74955
         TabIndex        =   44
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   3000
            Index           =   3
            Left            =   165
            TabIndex        =   59
            Top             =   135
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
            TabPicture(0)   =   "frmVtasNacionRubro.frx":10736
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(3)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":10752
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(3)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe / Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":1076E
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(3)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":1078A
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(3)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   3
               Left            =   -74910
               OleObjectBlob   =   "frmVtasNacionRubro.frx":107A6
               TabIndex        =   63
               Top             =   75
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   3
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":10C0F
               TabIndex        =   62
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   3
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":10F66
               TabIndex        =   61
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   3
               Left            =   90
               OleObjectBlob   =   "frmVtasNacionRubro.frx":112D9
               TabIndex        =   60
               Top             =   90
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   2
         Left            =   -74895
         TabIndex        =   43
         Top             =   795
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   3000
            Index           =   2
            Left            =   150
            TabIndex        =   54
            Top             =   135
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
            TabPicture(0)   =   "frmVtasNacionRubro.frx":11742
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(2)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":1175E
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(2)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe / Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":1177A
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(2)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":11796
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(2)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   2
               Left            =   -74910
               OleObjectBlob   =   "frmVtasNacionRubro.frx":117B2
               TabIndex        =   58
               Top             =   75
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   2
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":11C1B
               TabIndex        =   57
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   2
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":11F72
               TabIndex        =   56
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmVtasNacionRubro.frx":122E5
               TabIndex        =   55
               Top             =   105
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3180
         Index           =   0
         Left            =   405
         TabIndex        =   42
         Top             =   810
         Width           =   7380
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   3000
            Index           =   0
            Left            =   180
            TabIndex        =   49
            Top             =   150
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
            TabPicture(0)   =   "frmVtasNacionRubro.frx":1274E
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":1276A
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(0)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe / Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":12786
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(0)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":127A2
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(0)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   0
               Left            =   -74265
               OleObjectBlob   =   "frmVtasNacionRubro.frx":127BE
               TabIndex        =   53
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   0
               Left            =   -74295
               OleObjectBlob   =   "frmVtasNacionRubro.frx":12B15
               TabIndex        =   52
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   0
               Left            =   165
               OleObjectBlob   =   "frmVtasNacionRubro.frx":12E88
               TabIndex        =   51
               Top             =   120
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   0
               Left            =   -74820
               OleObjectBlob   =   "frmVtasNacionRubro.frx":132F1
               TabIndex        =   50
               Top             =   120
               Width           =   6705
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3165
         Index           =   1
         Left            =   -73980
         TabIndex        =   36
         Top             =   810
         Width           =   7275
         Begin TabDlg.SSTab tabSubRubro 
            Height          =   2970
            Index           =   1
            Left            =   120
            TabIndex        =   37
            Top             =   135
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   5239
            _Version        =   327680
            TabOrientation  =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Importe"
            TabPicture(0)   =   "frmVtasNacionRubro.frx":1375A
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprRubroImporte(1)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "frmVtasNacionRubro.frx":13776
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprRubroUnidades(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Importe / Unidades"
            TabPicture(2)   =   "frmVtasNacionRubro.frx":13792
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprRubroPromedio(1)"
            Tab(2).Control(0).Enabled=   0   'False
            TabCaption(3)   =   "Participación"
            TabPicture(3)   =   "frmVtasNacionRubro.frx":137AE
            Tab(3).ControlCount=   1
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SprRubroPartic(1)"
            Tab(3).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprRubroPartic 
               Height          =   2400
               Index           =   1
               Left            =   -74895
               OleObjectBlob   =   "frmVtasNacionRubro.frx":137CA
               TabIndex        =   41
               Top             =   90
               Width           =   6705
            End
            Begin FPSpread.vaSpread SprRubroUnidades 
               Height          =   2400
               Index           =   1
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":13C33
               TabIndex        =   40
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroPromedio 
               Height          =   2400
               Index           =   1
               Left            =   -74505
               OleObjectBlob   =   "frmVtasNacionRubro.frx":13F8A
               TabIndex        =   39
               Top             =   105
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprRubroImporte 
               Height          =   2400
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmVtasNacionRubro.frx":142FD
               TabIndex        =   38
               Top             =   105
               Width           =   6705
            End
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
         Picture         =   "frmVtasNacionRubro.frx":14766
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   825
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   630
         Left            =   60
         Picture         =   "frmVtasNacionRubro.frx":1589C
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   720
         Picture         =   "frmVtasNacionRubro.frx":15E2E
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   705
         Picture         =   "frmVtasNacionRubro.frx":15F30
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
         Picture         =   "frmVtasNacionRubro.frx":16032
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
               Picture         =   "frmVtasNacionRubro.frx":16854
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmVtasNacionRubro.frx":169C6
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
               Picture         =   "frmVtasNacionRubro.frx":16B38
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2550
               Picture         =   "frmVtasNacionRubro.frx":16CAA
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
Attribute VB_Name = "frmVtaNacionRubro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset
Dim RsDataAnt  As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consulta

Dim CantTab As Integer
Private Sub L_CargarLista()
Dim sql As String
Dim rs As Recordset
Dim ind As Integer

sql = " Select descrip,cod_pais From ventas.paises where cod_pais not in (0,9) Order by cod_pais "

CantTab = 0
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    ind = 0
    Do While Not rs.EOF
        lstNacion.AddItem rs!Descrip
        lstNacionSolapa.AddItem str(rs!cod_pais)
        
        ind = ind + 1
        
        rs.MoveNext
    Loop
    Aplicacion.CerrarDAO rs
End If

For ind = 0 To 19
    If ind < 7 Then
    lstNacion.Selected(ind) = True
    Else
    lstNacion.Selected(ind) = False
    End If
 Next

CantTab = 8

End Sub


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

Private Sub L_SetearSolapasNacion()
Dim i

For i = 0 To 19
    tabRubro.TabVisible(i) = False
Next

End Sub

Private Sub l_SetearTamanos()
Dim i

For i = 0 To 19
    Frame6(i).Visible = False
Next
    
If lstNacion.SelCount < 6 Then
   tabRubro.TabsPerRow = lstNacion.SelCount
   tabRubro.Width = 8000
ElseIf lstNacion.SelCount < 7 Then
   tabRubro.TabsPerRow = 6
   tabRubro.Width = 9800
Else
   tabRubro.TabsPerRow = 8
   tabRubro.Width = 10800
End If

For i = 0 To 19 'lstNacion.SelCount
    If lstNacion.Selected(i) = True Then
        tabRubro.Tab = i
        Frame6(i).Left = (tabRubro.Width / 2) - (Frame6(0).Width / 2)
        Frame6(i).Top = 500 + (Int((lstNacion.SelCount - 1) / 8) * 150)
        Frame6(i).Visible = True
        'Frame6(Val(lstNacionSolapa.List(i))).Refresh
    End If
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

Public Function L_DecoNac()
Dim i As Integer
Dim cod As Integer

    
    cod = -1
    For i = 0 To 20
        If Val(lstNacionSolapa.List(i)) = RsData!NAC Then
           'If lstNacion.Selected(i) Then
                cod = i
           'End If
           Exit For
        End If
    Next
    
    L_DecoNac = cod
    
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

For i = 0 To lstNacion.ListCount - 1
  SprRubroImporte(i).MaxRows = 0
  SprRubroUnidades(i).MaxRows = 0
  SprRubroPromedio(i).MaxRows = 0
  SprRubroPartic(i).MaxRows = 0
Next i
Frame6(tabRubro.Tab).Refresh


End Sub

Public Sub L_LlenarRubro()

Dim sqlL As String
Dim n As Integer
Dim RsDataRub As Recordset
Dim rub As String
Dim linea As Integer

sqlL = " SELECT "
sqlL = sqlL & " cod_rubr descrip "
sqlL = sqlL & " FROM baires.rubro Where cod_rubr <> 'REG' "
sqlL = sqlL & " ORDER BY cod_rubr  "

If Aplicacion.ObtenerRsDAO(sqlL, RsDataRub) Then
   If Aplicacion.CantReg(RsDataRub) > 0 Then
      For n = 0 To lstNacion.ListCount - 1
       SprRubroImporte(n).MaxRows = 0
       SprRubroUnidades(n).MaxRows = 0
       SprRubroPromedio(n).MaxRows = 0
      Next
      rub = RsDataRub!Descrip
      Do While Not RsDataRub.EOF
         For n = 0 To lstNacion.ListCount - 1
           SprRubroImporte(n).MaxRows = SprRubroImporte(n).MaxRows + 1
           SprRubroUnidades(n).MaxRows = SprRubroUnidades(n).MaxRows + 1
           SprRubroPromedio(n).MaxRows = SprRubroPromedio(n).MaxRows + 1
           SprRubroPartic(n).MaxRows = SprRubroPartic(n).MaxRows + 1
           
           SprRubroImporte(n).SetText 1, SprRubroImporte(n).MaxRows, Trim(RsDataRub!Descrip)
           SprRubroImporte(n).SetText 7, SprRubroImporte(n).MaxRows, Trim(RsDataRub!Descrip)
           SprRubroUnidades(n).SetText 1, SprRubroUnidades(n).MaxRows, Trim(RsDataRub!Descrip)
           SprRubroPromedio(n).SetText 1, SprRubroPromedio(n).MaxRows, Trim(RsDataRub!Descrip)
           SprRubroPartic(n).SetText 1, SprRubroPartic(n).MaxRows, Trim(RsDataRub!Descrip)

         Next n
           RsDataRub.MoveNext
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
Dim nacion As String
Dim Fila_Desde As Integer
Dim Imp_Total As Double
Dim Uni_Total As Double
Dim linea As Integer
Dim rub As Variant

nacion = RsData!Descrip
linea = 1

Do While Not RsData.EOF And (RsData!Descrip = nacion)
   spr.GetText 1, linea, rub
   If Trim(rub) = RsData!cod_rubr Then
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

Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i As Integer

'On Error GoTo ErrRubLoc:

frmVtaNacionRubro.caption = Aplicacion.SeteoProceso(frmVtaNacionRubro.caption)

'Consulta Actual
sql = " SELECT decode(nacionalidad,0,'Resto del Mundo',9,'Resto del Mundo',descrip) descrip ,"
sql = sql & " decode(nacionalidad,0,8,9,8,nacionalidad) nac ,"
sql = sql & " COD_RUBR , SUM(IMPORTE) importe, SUM(UNIDADES) unidades"
sql = sql & " FROM ESTADIS.VENTA_RNC R, VENTAS.PAISES P "
sql = sql & " Where COD_PAIS = NACIONALIDAD AND COD_RUBR <> 'REG' "
sql = sql & L_Armarcondicion(1)
sql = sql & " GROUP BY decode(nacionalidad,0,8,9,8,nacionalidad),decode(nacionalidad,0,'Resto del Mundo',9,'Resto del Mundo',descrip),COD_RUBR   "


'Consulta Anterior
sqlX = " SELECT decode(nacionalidad,0,'Resto del Mundo',9,'Resto del Mundo',descrip) descrip, "
sqlX = sqlX & " decode(nacionalidad,0,8,9,8,nacionalidad) nac ,"
sqlX = sqlX & " COD_RUBR , SUM(IMPORTE) importe, SUM(UNIDADES) unidades"
sqlX = sqlX & " FROM ESTADIS.VENTA_RNC R, VENTAS.PAISES P "
sqlX = sqlX & " Where COD_PAIS = NACIONALIDAD AND COD_RUBR <> 'REG' "
sqlX = sqlX & L_Armarcondicion(2)
sqlX = sqlX & " GROUP BY decode(nacionalidad,0,8,9,8,nacionalidad), decode(nacionalidad,0,'Resto del Mundo',9,'Resto del Mundo',descrip),COD_RUBR  "

If Aplicacion.ObtenerRsDAO(sql, RsData) Then
   
    If Aplicacion.CantReg(RsData) > 0 Then
        L_LlenarRubro
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabRubro.Enabled = True
        
        L_DecoEspigon
        
        If Aplicacion.ObtenerRsDAO(sqlX, RsData) Then
           L_DecoEspigonAnt
        End If
        
        L_Resaltar
            
        For i = 0 To lstNacion.ListCount - 1
            Spread_TotalesGrillas SprRubroImporte(i), 2, 2
            L_Participacion SprRubroPartic(i), SprRubroImporte(i)
        Next
        
    End If
    
    Aplicacion.CerrarDAO RsData
    Frame6(tabRubro.Tab).Refresh
End If

ErrRubLoc:
    frmVtaNacionRubro.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub




Private Sub L_Resaltar()

Dim fila As Integer
Dim i

For i = 0 To lstNacion.ListCount - 1
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
Dim NAC As String

Do While Not RsData.EOF
   
   NAC = RsData!NAC
   
   x = L_DecoNac
   If x <> -1 Then
   Do While RsData!NAC = NAC And Not RsData.EOF
        If IsNumeric(x) Then
           L_PonerDatos SprRubroImporte(x), SprRubroUnidades(x), SprRubroPromedio(x), 2
          Else
              RsData.MoveNext
        End If
        If RsData.EOF Then
            Exit Do
        End If
   
   Loop
   Else
    Do While RsData!NAC = NAC And Not RsData.EOF
        RsData.MoveNext
        If RsData.EOF Then
            Exit Do
        End If
    Loop
   End If
Loop

End Sub

Private Sub L_DecoEspigonAnt()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant
Dim x As Variant
Dim NAC As String

Do While Not RsData.EOF
   
   NAC = RsData!NAC
   x = L_DecoNac
   Do While RsData!NAC = NAC
        If IsNumeric(x) Then
           L_PonerDatos SprRubroImporte(x), SprRubroUnidades(x), SprRubroPromedio(x), 3
        Else
           RsData.MoveNext
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
FrmNacionalidad.Text1 = 2
FrmNacionalidad.Show 1
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

'L_SetearSolapasNacion
'L_CargarLista
End Sub

Private Sub Form_Load()
Dim i

Width = Screen.Width * 0.93
Height = Screen.Height * 0.77
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
L_SetearSolapasNacion
L_CargarLista


'frmPrincipal.lstForms.AddItem "frmVtaRubroLocal"


End Sub


Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmVtaRubroLocal"
End Sub


Private Sub lstNacion_ItemCheck(item As Integer)
Dim a

tabRubro.TabVisible(item) = Not tabRubro.TabVisible(item)
'Frame6(lstNacionSolapa.List(item)).Refresh
'TabSubrubro(4).Visible = True

l_SetearTamanos

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
Dim NAC As Integer, Total As Variant, TotalANT As Variant
Dim DescNac As Variant
Dim d As Variant

 If Index = 0 Then
    NAC = 1
 ElseIf Index = 1 Then
    NAC = 2
 ElseIf Index = 2 Then
    NAC = 3
 ElseIf Index = 3 Then
    NAC = 4
 ElseIf Index = 4 Then
    NAC = 5
 ElseIf Index = 5 Then
    NAC = 6
 ElseIf Index = 6 Then
    NAC = 7
 ElseIf Index = 7 Then
    NAC = 8
 ElseIf Index = 8 Then
    NAC = 10
 ElseIf Index = 9 Then
    NAC = 11
 ElseIf Index = 10 Then
    NAC = 12
 ElseIf Index = 11 Then
    NAC = 13
 ElseIf Index = 12 Then
    NAC = 14
 ElseIf Index = 13 Then
    NAC = 20
 ElseIf Index = 14 Then
    NAC = 21
 ElseIf Index = 15 Then
    NAC = 22
 ElseIf Index = 16 Then
    NAC = 23
 ElseIf Index = 17 Then
    NAC = 24
 ElseIf Index = 18 Then
    NAC = 25
 ElseIf Index = 19 Then
    NAC = 26
 
 End If

 SprRubroImporte(Index).GetText 7, Row, d
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


 frmVtaSubRubro.Mostrar "", Trim(d & " - " & tabRubro.TabCaption(Index)), sql, Val(Total), Val(TotalANT), "Venta por Sub Rubro y Nacionalidad"
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


