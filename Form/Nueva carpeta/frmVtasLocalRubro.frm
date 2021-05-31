VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.0#0"; "CRYSTL32.OCX"
Begin VB.Form frmVtaLocalRubro 
   Caption         =   "Ventas por Rubro/Local"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9705
   Begin Crystal.CrystalReport rptform 
      Left            =   6375
      Top             =   5820
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327682
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   1500
      Left            =   7815
      TabIndex        =   1
      Top             =   -15
      Width           =   1455
      Begin VB.CommandButton CmdReport 
         Height          =   615
         Left            =   60
         Picture         =   "frmVtasLocalRubro.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   825
         Width           =   600
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   630
         Left            =   60
         Picture         =   "frmVtasLocalRubro.frx":1136
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Picture         =   "frmVtasLocalRubro.frx":16C8
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmVtasLocalRubro.frx":17CA
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmVtasLocalRubro.frx":18CC
         Style           =   1  'Graphical
         TabIndex        =   10
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
         TabIndex        =   2
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
            TabIndex        =   13
            Top             =   225
            Width           =   3015
            Begin VB.CommandButton botHelpFDAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmVtasLocalRubro.frx":20EE
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmVtasLocalRubro.frx":2260
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   600
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesdeAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   16
               Top             =   255
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               _Version        =   327680
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-##-####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFHastaAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   17
               Top             =   600
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               _Version        =   327680
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
               TabIndex        =   19
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
               TabIndex        =   18
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
            TabIndex        =   3
            Top             =   225
            Width           =   3015
            Begin VB.CommandButton botHelpFH 
               Height          =   345
               Left            =   2535
               Picture         =   "frmVtasLocalRubro.frx":23D2
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2550
               Picture         =   "frmVtasLocalRubro.frx":2544
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   255
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesde 
               Height          =   285
               Left            =   1380
               TabIndex        =   6
               Top             =   270
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
               Left            =   1395
               TabIndex        =   7
               Top             =   615
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               _Version        =   327680
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
               TabIndex        =   9
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
               TabIndex        =   8
               Top             =   270
               Width           =   1185
            End
         End
      End
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   4590
      Left            =   60
      TabIndex        =   23
      Top             =   1590
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   8096
      _Version        =   327680
      Tabs            =   5
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
      TabCaption(0)   =   "TOTAL CIA"
      TabPicture(0)   =   "frmVtasLocalRubro.frx":26B6
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tabTotal"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "EZE - A"
      TabPicture(1)   =   "frmVtasLocalRubro.frx":26D2
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TabEzeA"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "EZE - B"
      TabPicture(2)   =   "frmVtasLocalRubro.frx":26EE
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TabEzeB"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "AEP"
      TabPicture(3)   =   "frmVtasLocalRubro.frx":270A
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TabAep"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "INTERIOR"
      TabPicture(4)   =   "frmVtasLocalRubro.frx":2726
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TabInterior"
      Tab(4).Control(0).Enabled=   0   'False
      Begin TabDlg.SSTab TabEzeA 
         Height          =   4000
         Left            =   -74910
         TabIndex        =   24
         Top             =   400
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   7038
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   5
         TabsPerRow      =   6
         TabHeight       =   520
         ForeColor       =   16711680
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
         TabPicture(0)   =   "frmVtasLocalRubro.frx":2742
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSTab1"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "L05"
         TabPicture(1)   =   "frmVtasLocalRubro.frx":275E
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSTab2"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "L06"
         TabPicture(2)   =   "frmVtasLocalRubro.frx":277A
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SSTab3"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "L07"
         TabPicture(3)   =   "frmVtasLocalRubro.frx":2796
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SSTab4"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "L08"
         TabPicture(4)   =   "frmVtasLocalRubro.frx":27B2
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "SSTab5"
         Tab(4).Control(0).Enabled=   0   'False
         Begin TabDlg.SSTab SSTab1 
            Height          =   3375
            Left            =   105
            TabIndex        =   34
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5953
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":27CE
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImpResumen(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNI"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":27EA
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnidResumen(0)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprUnidResumen 
               Height          =   3105
               Index           =   0
               Left            =   -74835
               OleObjectBlob   =   "frmVtasLocalRubro.frx":2806
               TabIndex        =   59
               Top             =   135
               Width           =   8310
            End
            Begin FPSpread.vaSpread SprImpResumen 
               Height          =   3060
               Index           =   0
               Left            =   165
               OleObjectBlob   =   "frmVtasLocalRubro.frx":2CFE
               TabIndex        =   55
               Top             =   150
               Width           =   8370
            End
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   3360
            Left            =   -74895
            TabIndex        =   35
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5927
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":3218
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":3234
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(0)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3075
               Index           =   0
               Left            =   150
               OleObjectBlob   =   "frmVtasLocalRubro.frx":3250
               TabIndex        =   63
               Top             =   135
               Width           =   8385
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3060
               Index           =   0
               Left            =   -74850
               OleObjectBlob   =   "frmVtasLocalRubro.frx":376A
               TabIndex        =   77
               Top             =   150
               Width           =   8370
            End
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   3390
            Left            =   -74895
            TabIndex        =   36
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5980
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":3C74
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(1)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":3C90
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(1)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3120
               Index           =   1
               Left            =   150
               OleObjectBlob   =   "frmVtasLocalRubro.frx":3CAC
               TabIndex        =   64
               Top             =   135
               Width           =   8385
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3105
               Index           =   1
               Left            =   -74880
               OleObjectBlob   =   "frmVtasLocalRubro.frx":41C6
               TabIndex        =   78
               Top             =   135
               Width           =   8400
            End
         End
         Begin TabDlg.SSTab SSTab4 
            Height          =   3375
            Left            =   -74895
            TabIndex        =   37
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5953
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":46BE
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(2)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":46DA
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(2)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3075
               Index           =   2
               Left            =   150
               OleObjectBlob   =   "frmVtasLocalRubro.frx":46F6
               TabIndex        =   65
               Top             =   120
               Width           =   8385
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3105
               Index           =   2
               Left            =   -74850
               OleObjectBlob   =   "frmVtasLocalRubro.frx":4C10
               TabIndex        =   79
               Top             =   105
               Width           =   8370
            End
         End
         Begin TabDlg.SSTab SSTab5 
            Height          =   3390
            Left            =   -74895
            TabIndex        =   38
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5980
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":5108
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(3)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":5124
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(3)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3120
               Index           =   3
               Left            =   135
               OleObjectBlob   =   "frmVtasLocalRubro.frx":5140
               TabIndex        =   66
               Top             =   120
               Width           =   8400
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3105
               Index           =   3
               Left            =   -74865
               OleObjectBlob   =   "frmVtasLocalRubro.frx":565A
               TabIndex        =   80
               Top             =   120
               Width           =   8400
            End
         End
      End
      Begin TabDlg.SSTab TabEzeB 
         Height          =   4000
         Left            =   -74910
         TabIndex        =   25
         Top             =   400
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   7038
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   4
         TabsPerRow      =   6
         TabHeight       =   520
         ForeColor       =   16711680
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
         TabPicture(0)   =   "frmVtasLocalRubro.frx":5B52
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSTab6"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "L01"
         TabPicture(1)   =   "frmVtasLocalRubro.frx":5B6E
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSTab19"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "L02"
         TabPicture(2)   =   "frmVtasLocalRubro.frx":5B8A
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SSTab8"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "L03"
         TabPicture(3)   =   "frmVtasLocalRubro.frx":5BA6
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SSTab9"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SSTab7 
            Height          =   2400
            Index           =   4
            Left            =   -74505
            OleObjectBlob   =   "frmVtasLocalRubro.frx":5BC2
            TabIndex        =   26
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SSTab14 
            Height          =   2400
            Index           =   5
            Left            =   -74505
            OleObjectBlob   =   "frmVtasLocalRubro.frx":5F29
            TabIndex        =   27
            Top             =   105
            Width           =   5505
         End
         Begin TabDlg.SSTab SSTab6 
            Height          =   3300
            Left            =   100
            TabIndex        =   39
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5821
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":6274
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImpResumen(1)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":6290
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnidResumen(1)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprUnidResumen 
               Height          =   2850
               Index           =   1
               Left            =   -74835
               OleObjectBlob   =   "frmVtasLocalRubro.frx":62AC
               TabIndex        =   60
               Top             =   195
               Width           =   8370
            End
            Begin FPSpread.vaSpread SprImpResumen 
               Height          =   2955
               Index           =   1
               Left            =   135
               OleObjectBlob   =   "frmVtasLocalRubro.frx":67A4
               TabIndex        =   56
               Top             =   135
               Width           =   8370
            End
         End
         Begin TabDlg.SSTab SSTab8 
            Height          =   3375
            Left            =   -74895
            TabIndex        =   40
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5953
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":6CBE
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(5)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":6CDA
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(5)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3090
               Index           =   5
               Left            =   135
               OleObjectBlob   =   "frmVtasLocalRubro.frx":6CF6
               TabIndex        =   68
               Top             =   120
               Width           =   8385
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3090
               Index           =   5
               Left            =   -74865
               OleObjectBlob   =   "frmVtasLocalRubro.frx":7210
               TabIndex        =   82
               Top             =   135
               Width           =   8370
            End
         End
         Begin TabDlg.SSTab SSTab9 
            Height          =   3375
            Left            =   -74895
            TabIndex        =   41
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5953
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":7708
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(6)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":7724
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(6)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3075
               Index           =   6
               Left            =   135
               OleObjectBlob   =   "frmVtasLocalRubro.frx":7740
               TabIndex        =   69
               Top             =   120
               Width           =   8430
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3090
               Index           =   6
               Left            =   -74865
               OleObjectBlob   =   "frmVtasLocalRubro.frx":7C5A
               TabIndex        =   83
               Top             =   135
               Width           =   8385
            End
         End
         Begin TabDlg.SSTab SSTab19 
            Height          =   3375
            Left            =   -74895
            TabIndex        =   53
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5953
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":8152
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(4)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":816E
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(4)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3105
               Index           =   4
               Left            =   120
               OleObjectBlob   =   "frmVtasLocalRubro.frx":818A
               TabIndex        =   67
               Top             =   120
               Width           =   8430
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3090
               Index           =   4
               Left            =   -74880
               OleObjectBlob   =   "frmVtasLocalRubro.frx":86A4
               TabIndex        =   81
               Top             =   120
               Width           =   8415
            End
         End
      End
      Begin TabDlg.SSTab TabAep 
         Height          =   4000
         Left            =   -74910
         TabIndex        =   28
         Top             =   400
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   7038
         _Version        =   327680
         TabOrientation  =   1
         TabsPerRow      =   6
         TabHeight       =   520
         ForeColor       =   16711680
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
         TabPicture(0)   =   "frmVtasLocalRubro.frx":8B9C
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSTab10"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "L09"
         TabPicture(1)   =   "frmVtasLocalRubro.frx":8BB8
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSTab11"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "L14"
         TabPicture(2)   =   "frmVtasLocalRubro.frx":8BD4
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SSTab12"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SSTab7 
            Height          =   2400
            Index           =   7
            Left            =   -74505
            OleObjectBlob   =   "frmVtasLocalRubro.frx":8BF0
            TabIndex        =   29
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SSTab14 
            Height          =   2400
            Index           =   8
            Left            =   -74505
            OleObjectBlob   =   "frmVtasLocalRubro.frx":8F57
            TabIndex        =   30
            Top             =   105
            Width           =   5505
         End
         Begin TabDlg.SSTab SSTab10 
            Height          =   3345
            Left            =   105
            TabIndex        =   42
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5900
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":92A2
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImpResumen(2)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":92BE
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnidResumen(2)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprUnidResumen 
               Height          =   3075
               Index           =   2
               Left            =   -74850
               OleObjectBlob   =   "frmVtasLocalRubro.frx":92DA
               TabIndex        =   61
               Top             =   120
               Width           =   8385
            End
            Begin FPSpread.vaSpread SprImpResumen 
               Height          =   3090
               Index           =   2
               Left            =   150
               OleObjectBlob   =   "frmVtasLocalRubro.frx":97D2
               TabIndex        =   57
               Top             =   120
               Width           =   8370
            End
         End
         Begin TabDlg.SSTab SSTab11 
            Height          =   3390
            Left            =   -74895
            TabIndex        =   43
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5980
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":9CEC
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(7)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":9D08
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(7)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3120
               Index           =   7
               Left            =   135
               OleObjectBlob   =   "frmVtasLocalRubro.frx":9D24
               TabIndex        =   70
               Top             =   120
               Width           =   8400
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3120
               Index           =   7
               Left            =   -74865
               OleObjectBlob   =   "frmVtasLocalRubro.frx":A23E
               TabIndex        =   84
               Top             =   120
               Width           =   8385
            End
         End
         Begin TabDlg.SSTab SSTab12 
            Height          =   3405
            Left            =   -74895
            TabIndex        =   44
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   6006
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":A736
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(8)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":A752
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(8)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3090
               Index           =   8
               Left            =   150
               OleObjectBlob   =   "frmVtasLocalRubro.frx":A76E
               TabIndex        =   71
               Top             =   150
               Width           =   8400
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3120
               Index           =   8
               Left            =   -74895
               OleObjectBlob   =   "frmVtasLocalRubro.frx":AC88
               TabIndex        =   85
               Top             =   135
               Width           =   8400
            End
         End
      End
      Begin TabDlg.SSTab TabInterior 
         Height          =   4000
         Left            =   -74910
         TabIndex        =   31
         Top             =   400
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   7038
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
         ForeColor       =   16711680
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
         TabPicture(0)   =   "frmVtasLocalRubro.frx":B180
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSTab13"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "L04"
         TabPicture(1)   =   "frmVtasLocalRubro.frx":B19C
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSTab20"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "L10"
         TabPicture(2)   =   "frmVtasLocalRubro.frx":B1B8
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SSTab15"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "L15"
         TabPicture(3)   =   "frmVtasLocalRubro.frx":B1D4
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SSTab16"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "L16"
         TabPicture(4)   =   "frmVtasLocalRubro.frx":B1F0
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "SSTab17"
         Tab(4).Control(0).Enabled=   0   'False
         TabCaption(5)   =   "L18"
         TabPicture(5)   =   "frmVtasLocalRubro.frx":B20C
         Tab(5).ControlCount=   1
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "SSTab18"
         Tab(5).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SSTab7 
            Height          =   2400
            Index           =   10
            Left            =   -74505
            OleObjectBlob   =   "frmVtasLocalRubro.frx":B228
            TabIndex        =   32
            Top             =   105
            Width           =   5505
         End
         Begin FPSpread.vaSpread SSTab14 
            Height          =   2400
            Index           =   11
            Left            =   -74490
            OleObjectBlob   =   "frmVtasLocalRubro.frx":B58F
            TabIndex        =   33
            Top             =   105
            Width           =   5505
         End
         Begin TabDlg.SSTab SSTab13 
            Height          =   3405
            Left            =   105
            TabIndex        =   45
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   6006
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":B8DA
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImpResumen(3)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":B8F6
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnidResumen(3)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprUnidResumen 
               Height          =   3120
               Index           =   3
               Left            =   -74865
               OleObjectBlob   =   "frmVtasLocalRubro.frx":B912
               TabIndex        =   62
               Top             =   135
               Width           =   8385
            End
            Begin FPSpread.vaSpread SprImpResumen 
               Height          =   3105
               Index           =   3
               Left            =   120
               OleObjectBlob   =   "frmVtasLocalRubro.frx":BE0A
               TabIndex        =   58
               Top             =   120
               Width           =   8430
            End
         End
         Begin TabDlg.SSTab SSTab15 
            Height          =   3375
            Left            =   -74895
            TabIndex        =   47
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5953
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":C324
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(10)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":C340
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(10)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3105
               Index           =   10
               Left            =   105
               OleObjectBlob   =   "frmVtasLocalRubro.frx":C35C
               TabIndex        =   73
               Top             =   105
               Width           =   8430
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3105
               Index           =   10
               Left            =   -74880
               OleObjectBlob   =   "frmVtasLocalRubro.frx":C876
               TabIndex        =   87
               Top             =   120
               Width           =   8430
            End
         End
         Begin TabDlg.SSTab SSTab16 
            Height          =   3405
            Left            =   -74895
            TabIndex        =   48
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   6006
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":CD6E
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(11)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":CD8A
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(11)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3165
               Index           =   11
               Left            =   120
               OleObjectBlob   =   "frmVtasLocalRubro.frx":CDA6
               TabIndex        =   74
               Top             =   105
               Width           =   8415
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3150
               Index           =   11
               Left            =   -74880
               OleObjectBlob   =   "frmVtasLocalRubro.frx":D2C0
               TabIndex        =   88
               Top             =   120
               Width           =   8385
            End
         End
         Begin TabDlg.SSTab SSTab17 
            Height          =   3420
            Left            =   -74895
            TabIndex        =   49
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   6033
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":D7B8
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(12)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":D7D4
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(12)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3165
               Index           =   12
               Left            =   120
               OleObjectBlob   =   "frmVtasLocalRubro.frx":D7F0
               TabIndex        =   75
               Top             =   105
               Width           =   8430
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3165
               Index           =   12
               Left            =   -74895
               OleObjectBlob   =   "frmVtasLocalRubro.frx":DD0A
               TabIndex        =   89
               Top             =   120
               Width           =   8430
            End
         End
         Begin TabDlg.SSTab SSTab18 
            Height          =   3420
            Left            =   -74895
            TabIndex        =   50
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   6033
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":E202
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(13)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":E21E
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(13)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3165
               Index           =   13
               Left            =   135
               OleObjectBlob   =   "frmVtasLocalRubro.frx":E23A
               TabIndex        =   76
               Top             =   120
               Width           =   8400
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3150
               Index           =   13
               Left            =   -74880
               OleObjectBlob   =   "frmVtasLocalRubro.frx":E754
               TabIndex        =   90
               Top             =   120
               Width           =   8430
            End
         End
         Begin TabDlg.SSTab SSTab20 
            Height          =   3390
            Left            =   -74895
            TabIndex        =   54
            Top             =   150
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5980
            _Version        =   327680
            TabOrientation  =   3
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "IMP"
            TabPicture(0)   =   "frmVtasLocalRubro.frx":EC4C
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprImp(9)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "UNID"
            TabPicture(1)   =   "frmVtasLocalRubro.frx":EC68
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprUnid(9)"
            Tab(1).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprImp 
               Height          =   3135
               Index           =   9
               Left            =   105
               OleObjectBlob   =   "frmVtasLocalRubro.frx":EC84
               TabIndex        =   72
               Top             =   105
               Width           =   8430
            End
            Begin FPSpread.vaSpread SprUnid 
               Height          =   3135
               Index           =   9
               Left            =   -74895
               OleObjectBlob   =   "frmVtasLocalRubro.frx":F19E
               TabIndex        =   86
               Top             =   120
               Width           =   8445
            End
         End
      End
      Begin TabDlg.SSTab tabTotal 
         Height          =   3735
         Left            =   435
         TabIndex        =   51
         Top             =   480
         Width           =   8640
         _ExtentX        =   15240
         _ExtentY        =   6588
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importe"
         TabPicture(0)   =   "frmVtasLocalRubro.frx":F642
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprImpTotal"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmVtasLocalRubro.frx":F65E
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprUnidTotal"
         Tab(1).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprUnidTotal 
            Height          =   3180
            Left            =   -74880
            OleObjectBlob   =   "frmVtasLocalRubro.frx":F67A
            TabIndex        =   52
            Top             =   120
            Width           =   8355
         End
         Begin FPSpread.vaSpread SprImpTotal 
            Height          =   3150
            Left            =   120
            OleObjectBlob   =   "frmVtasLocalRubro.frx":FB72
            TabIndex        =   46
            Top             =   150
            Width           =   8340
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
      TabIndex        =   20
      Top             =   5445
      Width           =   8730
   End
End
Attribute VB_Name = "frmVtaLocalRubro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset
Dim RsDataAnt  As Recordset
Dim rs As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas

Dim locales(0 To 13) As String

Private Sub L_PrepararTablaPRN()
Dim i As Integer, j As Integer
Dim valUniC As Variant, valImpC As Variant
Dim valUniN As Variant, valImpN As Variant
Dim valRubro As Variant
Dim SQL As String, depn(0 To 3) As String

depn(0) = "EZEINTAR  "
depn(1) = "EZEINTBS  "
depn(2) = "AEPAEP T  "
depn(3) = "INTINT U  "

Aplicacion.EjecutarDAO "TRUNCATE TABLE ESTADIS.PRN_RUBROLOCAL DROP STORAGE"

For i = 0 To 13
    For j = 1 To SprImp(i).MaxRows
        If j = SprImp(i).MaxRows Then
            valRubro = "TOT"
            SprImp(i).GetText 2, j, valImpC
            SprImp(i).GetText 3, j, valImpN
            SprUnid(i).GetText 2, j, valUniC
            SprUnid(i).GetText 3, j, valUniN
            
            SQL = "UPDATE estadis.PRN_RUBROLOCAL SET " _
            & " TOT_act_$ = " & valImpC _
            & ", TOT_ant_$ = " & valImpN _
            & " WHERE cod_depn = '" & Mid(locales(i), 1, 3) & "' And " _
            & " cod_sdep = '" & Trim(Mid(locales(i), 4, 4)) & "' And " _
            & " cod_local = '" & Mid(locales(i), 8, 3) & "'  "
        
        Else
            SprImp(i).GetText 1, j, valRubro
            SprImp(i).GetText 2, j, valImpC
            SprImp(i).GetText 3, j, valImpN
            SprUnid(i).GetText 2, j, valUniC
            SprUnid(i).GetText 3, j, valUniN
            
            SQL = "INSERT INTO estadis.PRN_RUBROLOCAL ( " _
            & " cod_depn , cod_sdep , " _
            & " cod_local , cod_rubro , " _
            & " act_$ , act_Un , " _
            & " ant_$ , ant_Un ) VALUES ( " _
            & "'" & Mid(locales(i), 1, 3) & "' , " _
            & "'" & Trim(Mid(locales(i), 4, 4)) & "' , " _
            & "'" & Mid(locales(i), 8, 3) & "' , " _
            & "'" & valRubro & "' , " _
            & valImpC & " , " & valUniC & " , " _
            & valImpN & " , " & valUniN & " ) "
        
        End If
            
        Aplicacion.EjecutarDAO SQL
        
    Next
Next
For i = 0 To 3
    For j = 1 To SprImpResumen(i).MaxRows
        If j = SprImpResumen(i).MaxRows Then
            valRubro = "TOT"
            SprImpResumen(i).GetText 2, j, valImpC
            SprImpResumen(i).GetText 3, j, valImpN
            SprUnidResumen(i).GetText 2, j, valUniC
            SprUnidResumen(i).GetText 3, j, valUniN
            
            SQL = "UPDATE estadis.PRN_RUBROLOCAL SET " _
            & " TOT_act_$ = " & valImpC _
            & ", TOT_ant_$ = " & valImpN _
            & " WHERE cod_depn = '" & Mid(depn(i), 1, 3) & "' And " _
            & " cod_sdep = '" & Trim(Mid(depn(i), 4, 4)) & "' And " _
            & " cod_local = '" & Mid(depn(i), 8, 3) & "'  "
        
        Else
            SprImpResumen(i).GetText 1, j, valRubro
            SprImpResumen(i).GetText 2, j, valImpC
            SprImpResumen(i).GetText 3, j, valImpN
            SprUnidResumen(i).GetText 2, j, valUniC
            SprUnidResumen(i).GetText 3, j, valUniN
            
            SQL = "INSERT INTO estadis.PRN_RUBROLOCAL ( " _
            & " cod_depn , cod_sdep , " _
            & " cod_local , cod_rubro , " _
            & " act_$ , act_Un , " _
            & " ant_$ , ant_Un ) VALUES ( " _
            & "'" & Mid(depn(i), 1, 3) & "' , " _
            & "'" & Trim(Mid(depn(i), 4, 4)) & "' , " _
            & "'" & Mid(depn(i), 8, 3) & "' , " _
            & "'" & valRubro & "' , " _
            & valImpC & " , " & valUniC & " , " _
            & valImpN & " , " & valUniN & " ) "
        
        End If
        Aplicacion.EjecutarDAO SQL
    Next
Next
End Sub

Private Sub L_TratarExcel(Titulo As String, subTit As String)

Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim fila As Integer, filaant As Integer
Dim i, n As Integer
Dim tit As Variant
Dim nombre As String
Dim color As Integer
Dim PeriodoActual As String
Dim PeriodoAnterior As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

frmVtaLocalRubro.caption = Aplicacion.SeteoProceso(frmVtaLocalRubro.caption)

If nombre <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    ''AppExcel.application.Visible = True

    ReDim titCol(SprImpTotal.MaxCols)
    col = 1
    fila = 1
    color = Exl_Gris
    For i = 1 To frmVtaLocalRubro.SprImpTotal.MaxCols
        SprImpTotal.GetText i, 0, tit
        titCol(i) = tit
    Next
       
       PeriodoActual = "Período Actual : " & frmVtaLocalRubro.mskFDesde.FormattedText & " - " & frmVtaLocalRubro.mskFHasta.FormattedText
       PeriodoAnterior = "Período Anterior : " & frmVtaLocalRubro.mskFDesdeAnt.FormattedText & " - " & frmVtaLocalRubro.mskFHastaAnt.FormattedText
    
       Exl_PonerValor AppExcel, fila, col, PeriodoActual
       fila = fila + 1
       Exl_PonerValor AppExcel, fila, col, PeriodoAnterior
       fila = fila + 1
'----------------------------------------------
        Exl_PonerValor AppExcel, fila, col, "TOTAL COMPAÑIA"
        rango = Exl_rangos(fila, fila, 1, 7)
        Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        Exl_ColorInt AppExcel, rango, color
           
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
            
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "IMPORTE"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge
        
        fila = fila + 2
        
        Exl_BajarGrillaExel SprImpTotal, AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprImpTotal.MaxRows, 2, 7)
        Exl_Format AppExcel, rango
          
        rango = Exl_rangos((fila + SprImpTotal.MaxRows), (fila + SprImpTotal.MaxRows), 2, 7)
        Exl_ColorInt AppExcel, rango, color
        
        '-------------------------
        fila = fila + SprImpResumen(n).MaxRows
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "UNIDADES"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 2
                        
        Exl_BajarGrillaExel SprUnidTotal, AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprUnidTotal.MaxRows, 2, 7)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila + SprUnidTotal.MaxRows, fila + SprUnidTotal.MaxRows, 2, 7)
        Exl_ColorInt AppExcel, rango, color
  
        fila = fila + SprUnidTotal.MaxRows + 3
  
'--------------------------------------------------
  
  
  
  For n = 0 To 3
        Exl_PonerValor AppExcel, fila, col, TabEspigon.TabCaption(n + 1)
        rango = Exl_rangos(fila, fila, 1, 7)
        Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        Exl_ColorInt AppExcel, rango, color
           
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
            
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "IMPORTE"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge
        
        fila = fila + 2
        
        Exl_BajarGrillaExel SprImpResumen(n), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprImpResumen(n).MaxRows, 2, 7)
        Exl_Format AppExcel, rango
          
        rango = Exl_rangos((fila + SprImpResumen(n).MaxRows), (fila + SprImpResumen(n).MaxRows), 2, 7)
        Exl_ColorInt AppExcel, rango, color
        
        '-------------------------
        fila = fila + SprImpResumen(n).MaxRows
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "UNIDADES"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 2
                        
        Exl_BajarGrillaExel SprUnidResumen(n), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprUnidResumen(n).MaxRows, 2, 7)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila + SprUnidResumen(n).MaxRows, fila + SprUnidResumen(n).MaxRows, 2, 7)
        Exl_ColorInt AppExcel, rango, color
  
        fila = fila + SprUnidResumen(n).MaxRows + 3
  
  Next n
'--------------------------------------------------
  For n = 1 To 3
        Exl_PonerValor AppExcel, fila, col, TabEZEB.TabCaption(n)
        rango = Exl_rangos(fila, fila, 1, 7)
        Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        Exl_ColorInt AppExcel, rango, color
           
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
            
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "IMPORTE"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge
        
        fila = fila + 2
        
        Exl_BajarGrillaExel SprImp(n + 3), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprImp(n + 3).MaxRows, 2, 7)
        Exl_Format AppExcel, rango
          
        rango = Exl_rangos((fila + SprImp(n).MaxRows), (fila + SprImp(n).MaxRows), 2, 7)
        Exl_ColorInt AppExcel, rango, color
        
        '-------------------------
        fila = fila + SprImp(n + 3).MaxRows
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "UNIDADES"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 2
                        
        Exl_BajarGrillaExel SprUnid(n + 3), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprUnid(n + 3).MaxRows, 2, 7)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila + SprUnid(n + 3).MaxRows, fila + SprUnid(n + 3).MaxRows, 2, 7)
        Exl_ColorInt AppExcel, rango, color
  
        fila = fila + SprUnid(n + 3).MaxRows + 3
  
  Next n

'--------------------------------------
  For n = 1 To 4
        Exl_PonerValor AppExcel, fila, col, tabEZEA.TabCaption(n)
        rango = Exl_rangos(fila, fila, 1, 7)
        Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        Exl_ColorInt AppExcel, rango, color
           
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
            
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "IMPORTE"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge
        
        fila = fila + 2
        
        Exl_BajarGrillaExel SprImp(n), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprImp(n).MaxRows, 2, 7)
        Exl_Format AppExcel, rango
          
        rango = Exl_rangos((fila + SprImp(n).MaxRows), (fila + SprImp(n).MaxRows), 2, 7)
        Exl_ColorInt AppExcel, rango, color
        
        '-------------------------
        fila = fila + SprImp(n).MaxRows
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "UNIDADES"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 2
                        
        Exl_BajarGrillaExel SprUnid(n), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprUnid(n).MaxRows, 2, 7)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila + SprUnid(n).MaxRows, fila + SprUnid(n).MaxRows, 2, 7)
        Exl_ColorInt AppExcel, rango, color
  
        fila = fila + SprUnid(n).MaxRows + 3
  
  Next n

'--------------------------------------
  For n = 1 To 2
        Exl_PonerValor AppExcel, fila, col, TabAEP.TabCaption(n)
        rango = Exl_rangos(fila, fila, 1, 7)
        Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        Exl_ColorInt AppExcel, rango, color
           
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
            
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "IMPORTE"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge
        
        fila = fila + 2
        
        Exl_BajarGrillaExel SprImp(n + 6), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprImp(n + 6).MaxRows, 2, 7)
        Exl_Format AppExcel, rango
          
        rango = Exl_rangos((fila + SprImp(n).MaxRows), (fila + SprImp(n).MaxRows), 2, 7)
        Exl_ColorInt AppExcel, rango, color
        
        '-------------------------
        fila = fila + SprImp(n + 6).MaxRows
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "UNIDADES"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 2
                        
        Exl_BajarGrillaExel SprUnid(n + 6), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprUnid(n + 6).MaxRows, 2, 7)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila + SprUnid(n + 6).MaxRows, fila + SprUnid(n + 6).MaxRows, 2, 7)
        Exl_ColorInt AppExcel, rango, color
  
        fila = fila + SprUnid(n + 6).MaxRows + 3
  
  Next n

'--------------------------------------
  For n = 1 To 5
        Exl_PonerValor AppExcel, fila, col, TabInterior.TabCaption(n)
        rango = Exl_rangos(fila, fila, 1, 7)
        Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        Exl_ColorInt AppExcel, rango, color
           
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
            
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "IMPORTE"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge
        
        fila = fila + 2
        
        Exl_BajarGrillaExel SprImp(n + 8), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprImp(n + 8).MaxRows, 2, 7)
        Exl_Format AppExcel, rango
          
        rango = Exl_rangos((fila + SprImp(n).MaxRows), (fila + SprImp(n).MaxRows), 2, 7)
        Exl_ColorInt AppExcel, rango, color
        
        '-------------------------
        fila = fila + SprImp(n + 8).MaxRows
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, "UNIDADES"
        rango = Exl_rangos(fila, fila, 1, 2)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 2
                        
        Exl_BajarGrillaExel SprUnid(n + 8), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + SprUnid(n + 8).MaxRows, 2, 7)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila + SprUnid(n + 8).MaxRows, fila + SprUnid(n + 8).MaxRows, 2, 7)
        Exl_ColorInt AppExcel, rango, color
  
        fila = fila + SprUnid(n + 8).MaxRows + 3
  
  Next n

'--------------------------------------


 
   ' Exl.Exl_AnchoCol AppExcel, frmVtaRubroLocal.SprRubroImporte(0).MaxCols, frmVtaRubroLocal.SprRubroImporte(0).MaxCols, 1
    Exl.Exl_AnchoCol AppExcel, 1, 1, 8
    Exl.Exl_AnchoCol AppExcel, 2, 2, 12
    Exl.Exl_AnchoCol AppExcel, 3, 3, 12
    Exl.Exl_AnchoCol AppExcel, 4, 4, 10
    
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    AppExcel.Application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    'If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
    '    AppExcel.PrintOut
    'End If
    
    AppExcel.SaveAs nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmVtaLocalRubro.caption = Aplicacion.SeteoFin
    Exit Sub


End Sub

Private Sub L_LlenarResumen()
        L_LlenarResumen_Aep
        L_LlenarResumen_Interior
        L_LlenarResumen_EzeA
        L_LlenarResumen_EzeB
        L_LlenarResumen_Cia
End Sub

Private Sub MeImprimir()
Dim i%

On Error GoTo ErrPrint:


'rptform.Connect = "ODBC;UID=" & Aplicacion.username & ";PWD=" & Aplicacion.password & ";DATABASE=INTER;"
rptform.Connect = "ODBC;UID=ESTADIS;PWD=ESTADISSYS;DATABASE=INTER;"
rptform.ReportFileName = RutaReportes
rptform.ReportFileName = RutaReportes & "LocRub.rpt"

rptform.ParameterFields(0) = "PrActFechaDesde;" & " DATE(" & Year(mskFDesde.FormattedText) & "," & Month(mskFDesde.FormattedText) & "," & Day(mskFDesde.FormattedText) & ")" & ";TRUE)"
rptform.ParameterFields(1) = "PrActFechaHasta;" & " DATE(" & Year(mskFHasta.FormattedText) & "," & Month(mskFHasta.FormattedText) & "," & Day(mskFHasta.FormattedText) & ")" & ";TRUE)"
rptform.ParameterFields(2) = "PrAntFechaDesde;" & " DATE(" & Year(mskFDesdeAnt.FormattedText) & "," & Month(mskFDesdeAnt.FormattedText) & "," & Day(mskFDesdeAnt.FormattedText) & ")" & ";TRUE)"
rptform.ParameterFields(3) = "PrAntFechaHasta;" & " DATE(" & Year(mskFHastaAnt.FormattedText) & "," & Month(mskFHastaAnt.FormattedText) & "," & Day(mskFHastaAnt.FormattedText) & ")" & ";TRUE)"

rptform.Destination = 0     ' 0 pantalla
                            '1 impresora

rptform.Action = True
DoEvents


Exit Sub
ErrPrint:
     Exit Sub

End Sub


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
     Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)
  Case Is = 2
     Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesdeAnt) & " And " & func_ToDate(fechaHastaAnt)
End Select

     Cond = Cond & " And cod_local <> 'L11' "

L_Armarcondicion = Cond

End Function

Private Sub L_LimpiarGrillas()
Dim i
labInfo.caption = ""

For i = 0 To 13
  SprImp(i).MaxRows = 0
  SprUnid(i).MaxRows = 0
Next i

For i = 0 To 3
 SprImpResumen(i).MaxRows = 0
 SprUnidResumen(i).MaxRows = 0
Next i

SprImpTotal.MaxRows = 0
SprUnidTotal.MaxRows = 0

End Sub

Public Sub L_LlenarLocales()

Dim sqlL As String
Dim n As Integer
Dim RsDataLoc As Recordset

sqlL = " SELECT distinct cod_rubr "
sqlL = sqlL & " FROM baires.rubro "
sqlL = sqlL & " ORDER BY cod_rubr "

If Aplicacion.ObtenerRsDAO(sqlL, RsDataLoc) Then
   If Aplicacion.CantReg(RsData) > 0 Then
      For n = 0 To 13
       SprImp(n).MaxRows = 0
       SprUnid(n).MaxRows = 0
      Next
      
      For n = 0 To 3
       SprImpResumen(n).MaxRows = 0
       SprUnidResumen(n).MaxRows = 0
      Next n
      
      SprImpTotal.MaxRows = 0
      SprUnidTotal.MaxRows = 0
      
      Do While Not RsDataLoc.EOF
                      
            For n = 0 To 13
             SprImp(n).MaxRows = SprImp(n).MaxRows + 1
             SprUnid(n).MaxRows = SprUnid(n).MaxRows + 1
             SprImp(n).SetText 1, SprImp(n).MaxRows + 1, Trim(RsDataLoc!cod_rubr)
             SprUnid(n).SetText 1, SprUnid(n).MaxRows + 1, Trim(RsDataLoc!cod_rubr)
            Next
           
            For n = 0 To 3
             SprImpResumen(n).MaxRows = SprImpResumen(n).MaxRows + 1
             SprUnidResumen(n).MaxRows = SprUnidResumen(n).MaxRows + 1
             SprImpResumen(n).SetText 1, SprImpResumen(n).MaxRows, Trim(RsDataLoc!cod_rubr)
             SprUnidResumen(n).SetText 1, SprUnidResumen(n).MaxRows, Trim(RsDataLoc!cod_rubr)
            Next n
            
            SprImpTotal.MaxRows = SprImpTotal.MaxRows + 1
            SprUnidTotal.MaxRows = SprUnidTotal.MaxRows + 1
            SprImpTotal.SetText 1, SprImpTotal.MaxRows, Trim(RsDataLoc!cod_rubr)
            SprUnidTotal.SetText 1, SprUnidTotal.MaxRows, Trim(RsDataLoc!cod_rubr)
           
            RsDataLoc.MoveNext
          
          If RsDataLoc.EOF Then
            Exit Do
          End If
          
   Loop

   
   End If
End If

End Sub








Private Sub L_Refrescar()
Dim SQL As String
Dim sqlX As String


On Error GoTo ErrLocRub:

frmVtaLocalRubro.caption = Aplicacion.SeteoProceso(frmVtaLocalRubro.caption)

'Consulta Actual
SQL = " Select cod_depn,cod_sdep,cod_local,cod_rubr,"
SQL = SQL & " NVL(sum(importe),0) importe,NVL(sum(unidades),0) unidades "
SQL = SQL & " From estadis.view_venta_rubro_local "
SQL = SQL & L_Armarcondicion(1)
SQL = SQL & " group by cod_depn,cod_sdep,cod_local,cod_rubr "
SQL = SQL & " order by cod_depn,cod_sdep,cod_local,cod_rubr "


'Consulta Anterior
sqlX = " Select cod_depn,cod_sdep,cod_local ,cod_rubr "
sqlX = sqlX & " ,NVL(sum(importe),0) importe,NVL(sum(unidades),0) unidades "
sqlX = sqlX & " From estadis.view_venta_rubro_local "
sqlX = sqlX & L_Armarcondicion(2)
sqlX = sqlX & " group by cod_depn,cod_sdep,cod_local,cod_rubr "
sqlX = sqlX & " order by cod_depn,cod_sdep,cod_local,cod_rubr "

If Aplicacion.ObtenerRsDAO(SQL, RsData) Then
   
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        
        L_LlenarLocales
        L_DecoEspigon
        
        If Aplicacion.ObtenerRsDAO(sqlX, RsDataAnt) Then
           L_DecoEspigonAnt
        End If
        
        L_LlenarTotales
        L_LlenarResumen
        L_Participacion
        L_Resaltar
        
        Aplicacion.CerrarDAO RsDataAnt
    
    End If
    
    Aplicacion.CerrarDAO RsData

End If

ErrLocRub:
    frmVtaLocalRubro.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub




Private Sub L_LlenarTotales()
Dim fila As Integer
Dim i


For i = 0 To 13
    Spread.Spread_TotalesGrillas SprImp(i), SprImp(i).MaxCols, 2
    Spread.Spread_TotalesGrillas SprUnid(i), SprUnid(i).MaxCols, 2
Next i

For i = 0 To 3
    Spread.Spread_TotalesGrillas SprImpResumen(i), SprImpResumen(i).MaxCols, 2
    Spread.Spread_TotalesGrillas SprUnidResumen(i), SprUnidResumen(i).MaxCols, 2
Next i

    Spread.Spread_TotalesGrillas SprImpTotal, SprImpTotal.MaxCols, 2
    Spread.Spread_TotalesGrillas SprUnidTotal, SprUnidTotal.MaxCols, 2


End Sub

Private Sub L_LlenarResumen_EzeA()

Dim i As Integer
Dim X As Integer
Dim Rubro As Variant
Dim ImpActual As Variant
Dim ImpAnterior As Variant
Dim ImpActualResumen As Double
Dim ImpAnteriorResumen As Double
Dim UniActual As Variant
Dim UniAnterior As Variant
Dim UniActualResumen As Long
Dim UniAnteriorResumen As Long
  
    For i = 1 To SprImpResumen(0).MaxRows - 1
      For X = 0 To 3
        SprImp(X).GetText 2, i, ImpActual
        SprImp(X).GetText 3, i, ImpAnterior
        SprUnid(X).GetText 2, i, UniActual
        SprUnid(X).GetText 3, i, UniAnterior
        'Acumulo los valores
        ImpActualResumen = ImpActualResumen + ImpActual
        ImpAnteriorResumen = ImpAnteriorResumen + ImpAnterior
        UniActualResumen = UniActualResumen + UniActual
        UniAnteriorResumen = UniAnteriorResumen + UniAnterior
      Next X
        'Lleno resumen
        SprImpResumen(0).SetText 2, i, str(ImpActualResumen)
        SprImpResumen(0).SetText 3, i, str(ImpAnteriorResumen)
        SprUnidResumen(0).SetText 2, i, str(UniActualResumen)
        SprUnidResumen(0).SetText 3, i, str(UniAnteriorResumen)
        If UniActualResumen = 0 Or ImpActualResumen = 0 Then
             SprUnidResumen(0).SetText 5, i, "0"
          Else
             SprUnidResumen(0).SetText 5, i, str(ImpActualResumen / UniActualResumen)
        End If
        If UniAnteriorResumen = 0 Or ImpAnteriorResumen = 0 Then
             SprUnidResumen(0).SetText 6, i, "0"
          Else
             SprUnidResumen(0).SetText 6, i, str(ImpAnteriorResumen / UniAnteriorResumen)
        End If
        ImpActualResumen = 0
        ImpAnteriorResumen = 0
        UniActualResumen = 0
        UniAnteriorResumen = 0
        
  Next i
    
End Sub

Private Sub L_LlenarResumen_EzeB()

Dim i As Integer
Dim X As Integer
Dim Rubro As Variant
Dim ImpActual As Variant
Dim ImpAnterior As Variant
Dim ImpActualResumen As Single
Dim ImpAnteriorResumen As Single
Dim UniActual As Variant
Dim UniAnterior As Variant
Dim UniActualResumen As Long
Dim UniAnteriorResumen As Long
  
    For i = 1 To SprImpResumen(1).MaxRows - 1
      For X = 4 To 6
        SprImp(X).GetText 2, i, ImpActual
        SprImp(X).GetText 3, i, ImpAnterior
        SprUnid(X).GetText 2, i, UniActual
        SprUnid(X).GetText 3, i, UniAnterior
        'Acumulo los valores
        ImpActualResumen = ImpActualResumen + ImpActual
        ImpAnteriorResumen = ImpAnteriorResumen + ImpAnterior
        UniActualResumen = UniActualResumen + UniActual
        UniAnteriorResumen = UniAnteriorResumen + UniAnterior
      Next X
        'Lleno resumen
        SprImpResumen(1).SetText 2, i, str(ImpActualResumen)
        SprImpResumen(1).SetText 3, i, str(ImpAnteriorResumen)
        SprUnidResumen(1).SetText 2, i, str(UniActualResumen)
        SprUnidResumen(1).SetText 3, i, str(UniAnteriorResumen)
        If UniActualResumen = 0 Or ImpActualResumen = 0 Then
             SprUnidResumen(1).SetText 5, i, "0"
          Else
             SprUnidResumen(1).SetText 5, i, str(ImpActualResumen / UniActualResumen)
        End If
        If UniAnteriorResumen = 0 Or ImpAnteriorResumen = 0 Then
             SprUnidResumen(1).SetText 6, i, "0"
          Else
             SprUnidResumen(1).SetText 6, i, str(ImpAnteriorResumen / UniAnteriorResumen)
        End If
        ImpActualResumen = 0
        ImpAnteriorResumen = 0
        UniActualResumen = 0
        UniAnteriorResumen = 0
        
  Next i
   
End Sub
Private Sub L_LlenarResumen_Cia()

Dim i As Integer
Dim X As Integer
Dim ImpActual As Variant
Dim ImpAnterior As Variant
Dim ImpActualCia As Double
Dim ImpAnteriorCia As Double
Dim UniActual As Variant
Dim UniAnterior As Variant
Dim UniActualCia As Long
Dim UniAnteriorCia As Long
  

  For i = 1 To SprImpTotal.MaxRows - 1
    For X = 0 To 3
        SprImpResumen(X).GetText 2, i, ImpActual
        SprImpResumen(X).GetText 3, i, ImpAnterior
        SprUnidResumen(X).GetText 2, i, UniActual
        SprUnidResumen(X).GetText 3, i, UniAnterior
        'Acumulo los valores
        ImpActualCia = ImpActualCia + ImpActual
        ImpAnteriorCia = ImpAnteriorCia + ImpAnterior
        UniActualCia = UniActualCia + UniActual
        UniAnteriorCia = UniAnteriorCia + UniAnterior
    Next X
    
        'Lleno resumen
        SprImpTotal.SetText 2, i, str(ImpActualCia)
        SprImpTotal.SetText 3, i, str(ImpAnteriorCia)
        SprUnidTotal.SetText 2, i, str(UniActualCia)
        SprUnidTotal.SetText 3, i, str(UniAnteriorCia)
        If UniActualCia = 0 Or ImpActualCia = 0 Then
             SprUnidTotal.SetText 5, i, "0"
          Else
             SprUnidTotal.SetText 5, i, str(ImpActualCia / UniActualCia)
        End If
        If UniAnteriorCia = 0 Or ImpAnteriorCia = 0 Then
             SprUnidTotal.SetText 6, i, "0"
          Else
             SprUnidTotal.SetText 6, i, str(ImpAnteriorCia / UniAnteriorCia)
        End If
        ImpActualCia = 0
        ImpAnteriorCia = 0
        UniActualCia = 0
        UniAnteriorCia = 0
    
  Next i

End Sub
Private Sub L_LlenarResumen_Interior()

Dim i As Integer
Dim X As Integer
Dim Rubro As Variant
Dim ImpActual As Variant
Dim ImpAnterior As Variant
Dim ImpActualResumen As Double
Dim ImpAnteriorResumen As Double
Dim UniActual As Variant
Dim UniAnterior As Variant
Dim UniActualResumen As Integer
Dim UniAnteriorResumen As Integer
  
  For i = 1 To SprImpResumen(3).MaxRows - 1
      For X = 9 To 13
        SprImp(X).GetText 2, i, ImpActual
        SprImp(X).GetText 3, i, ImpAnterior
        SprUnid(X).GetText 2, i, UniActual
        SprUnid(X).GetText 3, i, UniAnterior
        'Acumulo los valores
        ImpActualResumen = ImpActualResumen + ImpActual
        ImpAnteriorResumen = ImpAnteriorResumen + ImpAnterior
        UniActualResumen = UniActualResumen + UniActual
        UniAnteriorResumen = UniAnteriorResumen + UniAnterior
      Next X
        'Lleno resumen
        SprImpResumen(3).SetText 2, i, str(ImpActualResumen)
        SprImpResumen(3).SetText 3, i, str(ImpAnteriorResumen)
        SprUnidResumen(3).SetText 2, i, str(UniActualResumen)
        SprUnidResumen(3).SetText 3, i, str(UniAnteriorResumen)
        If UniActualResumen = 0 Or ImpActualResumen = 0 Then
             SprUnidResumen(3).SetText 5, i, "0"
          Else
             SprUnidResumen(3).SetText 5, i, str(ImpActualResumen / UniActualResumen)
        End If
        If UniAnteriorResumen = 0 Or ImpAnteriorResumen = 0 Then
             SprUnidResumen(3).SetText 6, i, "0"
          Else
             SprUnidResumen(3).SetText 6, i, str(ImpAnteriorResumen / UniAnteriorResumen)
        End If
        ImpActualResumen = 0
        ImpAnteriorResumen = 0
        UniActualResumen = 0
        UniAnteriorResumen = 0
  Next i
   
End Sub

Private Sub L_LlenarResumen_Aep()
Dim i As Integer
Dim X As Integer
Dim Rubro As Variant
Dim ImpActual As Variant
Dim ImpAnterior As Variant
Dim ImpActualResumen As Double
Dim ImpAnteriorResumen As Double
Dim UniActual As Variant
Dim UniAnterior As Variant
Dim UniActualResumen As Long
Dim UniAnteriorResumen As Long
Dim PromActualResumen As Double
Dim PromAnteriorResumen As Double
  
    For i = 1 To SprImpResumen(2).MaxRows - 1
      For X = 7 To 8
        'Tomo los valores
        SprImp(X).GetText 2, i, ImpActual
        SprImp(X).GetText 3, i, ImpAnterior
        SprUnid(X).GetText 2, i, UniActual
        SprUnid(X).GetText 3, i, UniAnterior
        'Acumulo los valores
        ImpActualResumen = ImpActualResumen + ImpActual
        ImpAnteriorResumen = ImpAnteriorResumen + ImpAnterior
        UniActualResumen = UniActualResumen + UniActual
        UniAnteriorResumen = UniAnteriorResumen + UniAnterior
      Next X
        'Lleno resumen
        SprImpResumen(2).SetText 2, i, str(ImpActualResumen)
        SprImpResumen(2).SetText 3, i, str(ImpAnteriorResumen)
        SprUnidResumen(2).SetText 2, i, str(UniActualResumen)
        SprUnidResumen(2).SetText 3, i, str(UniAnteriorResumen)
        If UniActualResumen = 0 Or ImpActualResumen = 0 Then
             SprUnidResumen(2).SetText 5, i, "0"
          Else
             SprUnidResumen(2).SetText 5, i, str(ImpActualResumen / UniActualResumen)
        End If
        If UniAnteriorResumen = 0 Or ImpAnteriorResumen = 0 Then
             SprUnidResumen(2).SetText 6, i, "0"
          Else
             SprUnidResumen(2).SetText 6, i, str(ImpAnteriorResumen / UniAnteriorResumen)
        End If
        ImpActualResumen = 0
        ImpAnteriorResumen = 0
        UniActualResumen = 0
        UniAnteriorResumen = 0
        PromActualResumen = 0
        PromAnteriorResumen = 0
        
  Next i
  
End Sub

Private Sub L_Resaltar()

Dim fila As Integer
Dim i As Integer
Dim valor1 As Variant
Dim Valor2 As Variant

For i = 0 To 13
    SprImp(i).GetText 2, SprImp(i).MaxRows, valor1
    SprImp(i).GetText 3, SprImp(i).MaxRows, Valor2
    If Valor2 > 0 Then
    SprImp(i).SetText 4, SprImp(i).MaxRows, str(((valor1 - Valor2) * 100) / Valor2)
    End If
    SprUnid(i).GetText 2, SprUnid(i).MaxRows, valor1
    SprUnid(i).GetText 3, SprUnid(i).MaxRows, Valor2
    If Valor2 > 0 Then
    SprUnid(i).SetText 4, SprUnid(i).MaxRows, str(((valor1 - Valor2) * 100) / Valor2)
    End If
    SprImp(i).GetText 2, SprUnid(i).MaxRows, valor1
    SprUnid(i).GetText 2, SprUnid(i).MaxRows, Valor2
    If Valor2 > 0 Then
    SprUnid(i).SetText 5, SprUnid(i).MaxRows, str(valor1 / Valor2)
    End If
    SprImp(i).GetText 3, SprUnid(i).MaxRows, valor1
    SprUnid(i).GetText 3, SprUnid(i).MaxRows, Valor2
    If Valor2 > 0 Then
    SprUnid(i).SetText 6, SprUnid(i).MaxRows, str(valor1 / Valor2)
    End If
Next i

For i = 0 To 3
    SprImpResumen(i).GetText 2, SprImpResumen(i).MaxRows, valor1
    SprImpResumen(i).GetText 3, SprImpResumen(i).MaxRows, Valor2
    If Valor2 > 0 Then
    SprImpResumen(i).SetText 4, SprImpResumen(i).MaxRows, str(((valor1 - Valor2) * 100) / Valor2)
    End If
    SprUnidResumen(i).GetText 2, SprUnidResumen(i).MaxRows, valor1
    SprUnidResumen(i).GetText 3, SprUnidResumen(i).MaxRows, Valor2
    If Valor2 > 0 Then
    SprUnidResumen(i).SetText 4, SprUnidResumen(i).MaxRows, str(((valor1 - Valor2) * 100) / Valor2)
    End If
    SprImpResumen(i).GetText 2, SprUnidResumen(i).MaxRows, valor1
    SprUnidResumen(i).GetText 2, SprUnidResumen(i).MaxRows, Valor2
    If Valor2 > 0 Then
    SprUnidResumen(i).SetText 5, SprUnidResumen(i).MaxRows, str(valor1 / Valor2)
    End If
    SprImpResumen(i).GetText 3, SprUnidResumen(i).MaxRows, valor1
    SprUnidResumen(i).GetText 3, SprUnidResumen(i).MaxRows, Valor2
    If Valor2 > 0 Then
    SprUnidResumen(i).SetText 6, SprUnidResumen(i).MaxRows, str(valor1 / Valor2)
    End If
Next i

    SprImpTotal.GetText 2, SprImpTotal.MaxRows, valor1
    SprImpTotal.GetText 3, SprImpTotal.MaxRows, Valor2
    If Valor2 > 0 Then
    SprImpTotal.SetText 4, SprImpTotal.MaxRows, str(((valor1 - Valor2) * 100) / Valor2)
    End If
    SprUnidTotal.GetText 2, SprUnidTotal.MaxRows, valor1
    SprUnidTotal.GetText 3, SprUnidTotal.MaxRows, Valor2
    If Valor2 > 0 Then
    SprUnidTotal.SetText 4, SprUnidTotal.MaxRows, str(((valor1 - Valor2) * 100) / Valor2)
    End If
    SprImpTotal.GetText 2, SprUnidTotal.MaxRows, valor1
    SprUnidTotal.GetText 2, SprUnidTotal.MaxRows, Valor2
    If Valor2 > 0 Then
    SprUnidTotal.SetText 5, SprUnidTotal.MaxRows, str(valor1 / Valor2)
    End If
    SprImpTotal.GetText 3, SprUnidTotal.MaxRows, valor1
    SprUnidTotal.GetText 3, SprUnidTotal.MaxRows, Valor2
    If Valor2 > 0 Then
    SprUnidTotal.SetText 6, SprUnidTotal.MaxRows, str(valor1 / Valor2)
    End If

'Resalto las celdas

For i = 0 To 13
 For fila = 1 To SprImpTotal.MaxRows
   Spread.spread_ResaltarCelda SprImp(i), 4, fila
   Spread.spread_ResaltarCelda SprImp(i), 7, fila
   Spread.spread_ResaltarCelda SprUnid(i), 4, fila
   Spread.spread_ResaltarCelda SprUnid(i), 7, fila
Next fila
Next i

For i = 0 To 3
  For fila = 1 To SprImpTotal.MaxRows
   Spread.spread_ResaltarCelda SprImpResumen(i), 4, fila
   Spread.spread_ResaltarCelda SprImpResumen(i), 7, fila
   Spread.spread_ResaltarCelda SprUnidResumen(i), 4, fila
   Spread.spread_ResaltarCelda SprUnidResumen(i), 7, fila
  Next fila
Next i

For fila = 1 To SprImpTotal.MaxRows
   Spread.spread_ResaltarCelda SprImpTotal, 4, fila
   Spread.spread_ResaltarCelda SprImpTotal, 7, fila
   Spread.spread_ResaltarCelda SprUnidTotal, 4, fila
   Spread.spread_ResaltarCelda SprUnidTotal, 7, fila
Next fila



End Sub

Private Sub L_Participacion()
Dim i As Integer
Dim fila As Integer
Dim valor As Variant
Dim total As Variant
Dim ValorAnt As Variant
Dim TotalAnt As Variant

For i = 0 To 13
   
   SprImp(i).GetText 2, SprImp(i).MaxRows, total
   SprImp(i).GetText 3, SprImp(i).MaxRows, TotalAnt
   
   For fila = 1 To SprImpTotal.MaxRows - 1
       SprImp(i).GetText 2, fila, valor
       SprImp(i).GetText 3, fila, ValorAnt
       If total > 0 Then
       SprImp(i).SetText 5, fila, str(valor * 100 / total)
       End If
       If TotalAnt > 0 Then
       SprImp(i).SetText 6, fila, str(ValorAnt * 100 / TotalAnt)
       End If
   Next fila

Next i

For i = 0 To 3

   SprImpResumen(i).GetText 2, SprImpResumen(i).MaxRows, total
   SprImpResumen(i).GetText 3, SprImpResumen(i).MaxRows, TotalAnt

     For fila = 1 To SprImpTotal.MaxRows - 1
        SprImpResumen(i).GetText 2, fila, valor
        SprImpResumen(i).GetText 3, fila, ValorAnt
        If total > 0 Then
        SprImpResumen(i).SetText 5, fila, str(valor * 100 / total)
        End If
        If TotalAnt > 0 Then
        SprImpResumen(i).SetText 6, fila, str(ValorAnt * 100 / TotalAnt)
        End If
     Next fila
Next i
 

SprImpTotal.GetText 2, SprImpTotal.MaxRows, total
SprImpTotal.GetText 3, SprImpTotal.MaxRows, TotalAnt

For fila = 1 To SprImpTotal.MaxRows - 1
        SprImpTotal.GetText 2, fila, valor
        SprImpTotal.GetText 3, fila, ValorAnt
        SprImpTotal.SetText 5, fila, str(valor * 100 / total)
        SprImpTotal.SetText 6, fila, str(ValorAnt * 100 / TotalAnt)
Next fila
 
End Sub
Private Sub L_DecoEspigon()
Dim fecha As String, loc As String
Dim i As Integer, indDep As Integer, ind As Integer
Dim Importe As Variant
Dim dato As Variant
Dim X As Variant
Dim Rubro As Variant
Dim linea As Integer
 

Do While Not RsData.EOF

    For i = 0 To 13
        loc = RsData!cod_local
            For ind = 0 To 13
                If Mid(locales(ind), 8, 3) = loc Then
                    Exit For
                End If
            Next
        linea = 1
        Do While loc = RsData!cod_local
             SprImp(ind).GetText 1, linea, Rubro
             If Rubro = RsData!cod_rubr Then
                SprImp(ind).SetText 2, linea, str(RsData!Importe)
                SprUnid(ind).SetText 2, linea, str(RsData!unidades)
                If RsData!Importe <> 0 Or RsData!Importe <> Null Or RsData!unidades <> 0 Or RsData!unidades <> Null Then
                  SprUnid(ind).SetText 5, linea, str(RsData!Importe / RsData!unidades)
                 Else
                  SprUnid(ind).SetText 5, linea, "0"
                End If
                RsData.MoveNext
                If RsData.EOF Then
                 Exit Do
                End If
            End If
        linea = linea + 1
        Loop
        If RsData.EOF Then
         Exit For
        End If
    
    Next
Loop
End Sub


Private Sub L_DecoEspigonAnt()
Dim fecha As String
Dim i As Integer, indDep As Integer, ind As Integer
Dim dato As Variant
Dim X As Variant
Dim Rubro As Variant
Dim linea As Integer
Dim loc As String

Do While Not RsDataAnt.EOF

    For i = 0 To 13
        loc = RsDataAnt!cod_local
        linea = 1
            For ind = 0 To 13
                If Mid(locales(ind), 8, 3) = loc Then
                    Exit For
                End If
            Next
        
        Do While loc = RsDataAnt!cod_local
             SprImp(ind).GetText 1, linea, Rubro
             If Rubro = RsDataAnt!cod_rubr Then
                SprImp(ind).SetText 3, linea, str(RsDataAnt!Importe)
                SprUnid(ind).SetText 3, linea, str(RsDataAnt!unidades)
                If (RsDataAnt!unidades = 0) Or (RsDataAnt!unidades = Null) Then
                   SprUnid(ind).SetText 6, linea, "0"
                 Else
                   SprUnid(ind).SetText 6, linea, str(RsDataAnt!Importe / RsDataAnt!unidades)
                End If
                
                RsDataAnt.MoveNext
                If RsDataAnt.EOF Then
                 Exit Do
                End If
              Else
                SprImp(ind).SetText 3, linea, "0"
                SprUnid(ind).SetText 3, linea, "0"
                SprUnid(ind).SetText 6, linea, "0"
                
                'RsDataAnt.MoveNext
                If RsDataAnt.EOF Then
                 Exit Do
                End If
              
            End If
        linea = linea + 1
        Loop
        If RsDataAnt.EOF Then
         Exit For
        End If
    
    Next
Loop


End Sub

Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        L_Refrescar
    Case 1

        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        'tabRubro.Enabled = False
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
    L_TratarExcel "Ventas por Local / Rubro", "Espigón Internacional A"
End Sub

Private Sub CmdReport_Click()
Dim SQL

On Error GoTo ErrPRN:

frmVtaLocalRubro.caption = Aplicacion.SeteoProceso(frmVtaLocalRubro.caption)

Aplicacion.ComienzoTrans
If Aplicacion.EjecutarDAO("lock table estadis.prn_rubrolocal in exclusive mode nowait") Then
    L_PrepararTablaPRN
Else
    MsgBox "POR FAVOR REINTENTAR EN POCOS MINUTOS...", vbOKOnly, "ATENCION"
End If

Aplicacion.TerminarConExitoTrans
MeImprimir

frmVtaLocalRubro.caption = Aplicacion.SeteoFin

Exit Sub
ErrPRN:
    Resume

End Sub

Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabRubro

End Sub

Private Sub Form_Load()
Dim i

Width = Screen.Width * 0.83
Height = Screen.Height * 0.75
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

locales(0) = "EZEINTAL05"
locales(1) = "EZEINTAL06"
locales(2) = "EZEINTAL07"
locales(3) = "EZEINTAL08"
locales(4) = "EZEINTBL01"
locales(5) = "EZEINTBL02"
locales(6) = "EZEINTBL03"
locales(7) = "AEPAEP L09"
locales(8) = "AEPAEP L14"
locales(9) = "INTCORDL04"
locales(10) = "INTMENDL10"
locales(11) = "INTMDPLL15"
locales(12) = "INTMENDL16"
locales(13) = "INTBARIL18"

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







