VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIndicLS 
   Caption         =   "INDICADORES por Entradas/Salidas"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   9195
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   9135
      Top             =   1890
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   3810
      Left            =   135
      TabIndex        =   2
      Top             =   1680
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   6720
      _Version        =   327680
      Tabs            =   6
      TabsPerRow      =   6
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
      TabPicture(0)   =   "frmIndicLS_T.frx":0000
      Tab(0).ControlCount=   5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "botExcel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "borGrBar(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "sprTotal"
      Tab(0).Control(4).Enabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmIndicLS_T.frx":001C
      Tab(1).ControlCount=   4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "borGrBar(1)"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "sprEZEA"
      Tab(1).Control(3).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmIndicLS_T.frx":0038
      Tab(2).ControlCount=   4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(4)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2(5)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "borGrBar(2)"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "SprEZEB"
      Tab(2).Control(3).Enabled=   0   'False
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmIndicLS_T.frx":0054
      Tab(3).ControlCount=   4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2(6)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label2(7)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "borGrBar(3)"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "sprAEP"
      Tab(3).Control(3).Enabled=   0   'False
      TabCaption(4)   =   "IN FLIGHT"
      TabPicture(4)   =   "frmIndicLS_T.frx":0070
      Tab(4).ControlCount=   4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "SprFLIGHT"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "borGrBar(4)"
      Tab(4).Control(1).Enabled=   -1  'True
      Tab(4).Control(2)=   "Label2(9)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label2(8)"
      Tab(4).Control(3).Enabled=   0   'False
      TabCaption(5)   =   "TOTAL EZEIZA"
      TabPicture(5)   =   "frmIndicLS_T.frx":008C
      Tab(5).ControlCount=   4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label2(10)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label2(11)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "SprEZEIZA"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "borGrBar(5)"
      Tab(5).Control(3).Enabled=   -1  'True
      Begin FPSpread.vaSpread SprFLIGHT 
         Height          =   3045
         Left            =   -74850
         OleObjectBlob   =   "frmIndicLS_T.frx":00A8
         TabIndex        =   48
         Top             =   570
         Width           =   7830
      End
      Begin FPSpread.vaSpread SprEZEB 
         Height          =   3045
         Left            =   -74865
         OleObjectBlob   =   "frmIndicLS_T.frx":086C
         TabIndex        =   47
         Top             =   600
         Width           =   7830
      End
      Begin FPSpread.vaSpread SprEZEIZA 
         Height          =   3045
         Left            =   -74880
         OleObjectBlob   =   "frmIndicLS_T.frx":1030
         TabIndex        =   42
         Top             =   600
         Width           =   7830
      End
      Begin FPSpread.vaSpread sprTotal 
         Height          =   3045
         Left            =   165
         OleObjectBlob   =   "frmIndicLS_T.frx":17F4
         TabIndex        =   16
         Top             =   570
         Width           =   7800
      End
      Begin FPSpread.vaSpread sprEZEA 
         Height          =   3045
         Left            =   -74835
         OleObjectBlob   =   "frmIndicLS_T.frx":1FB8
         TabIndex        =   31
         Top             =   570
         Width           =   7830
      End
      Begin FPSpread.vaSpread sprAEP 
         Height          =   3045
         Left            =   -74850
         OleObjectBlob   =   "frmIndicLS_T.frx":277C
         TabIndex        =   36
         Top             =   585
         Width           =   7830
      End
      Begin VB.CommandButton borGrBar 
         Enabled         =   0   'False
         Height          =   510
         Index           =   5
         Left            =   -66990
         Picture         =   "frmIndicLS_T.frx":2F40
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3150
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Enabled         =   0   'False
         Height          =   510
         Index           =   4
         Left            =   -66990
         Picture         =   "frmIndicLS_T.frx":3382
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3090
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Enabled         =   0   'False
         Height          =   510
         Index           =   3
         Left            =   -66990
         Picture         =   "frmIndicLS_T.frx":37C4
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3105
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Enabled         =   0   'False
         Height          =   510
         Index           =   2
         Left            =   -67005
         Picture         =   "frmIndicLS_T.frx":3C06
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3120
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Enabled         =   0   'False
         Height          =   510
         Index           =   1
         Left            =   -66975
         Picture         =   "frmIndicLS_T.frx":4048
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3105
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Enabled         =   0   'False
         Height          =   510
         Index           =   0
         Left            =   8040
         Picture         =   "frmIndicLS_T.frx":448A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2415
         Width           =   780
      End
      Begin VB.CommandButton botExcel 
         Caption         =   "Excel"
         Height          =   510
         Left            =   8040
         Picture         =   "frmIndicLS_T.frx":48CC
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3075
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "LLEGADAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   11
         Left            =   -68670
         TabIndex        =   44
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "SALIDAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   10
         Left            =   -73440
         TabIndex        =   43
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "SALIDAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   9
         Left            =   -73170
         TabIndex        =   41
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "LLEGADAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   8
         Left            =   -68400
         TabIndex        =   40
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "LLEGADAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   7
         Left            =   -68655
         TabIndex        =   38
         Top             =   375
         Width           =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "SALIDAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   6
         Left            =   -73425
         TabIndex        =   37
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "LLEGADAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   5
         Left            =   -68640
         TabIndex        =   35
         Top             =   390
         Width           =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "SALIDAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   4
         Left            =   -73410
         TabIndex        =   34
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "LLEGADAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   3
         Left            =   -68625
         TabIndex        =   33
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "SALIDAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   2
         Left            =   -73395
         TabIndex        =   32
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "LLEGADAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   1
         Left            =   6450
         TabIndex        =   30
         Top             =   390
         Width           =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "SALIDAS -->"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   165
         Index           =   0
         Left            =   1680
         TabIndex        =   29
         Top             =   405
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1500
      Left            =   8220
      TabIndex        =   1
      Top             =   -30
      Width           =   825
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
         Left            =   135
         Picture         =   "frmIndicLS_T.frx":4E5E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   165
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
         Left            =   135
         Picture         =   "frmIndicLS_T.frx":4F60
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
         Left            =   135
         Picture         =   "frmIndicLS_T.frx":5062
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
         Width           =   8040
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
            Left            =   4980
            TabIndex        =   22
            Top             =   210
            Width           =   3015
            Begin VB.CommandButton botHelpFDAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmIndicLS_T.frx":5884
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmIndicLS_T.frx":59F6
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   600
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesdeAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   25
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
               TabIndex        =   26
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
               TabIndex        =   28
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
               TabIndex        =   27
               Top             =   600
               Width           =   1185
            End
         End
         Begin VB.OptionButton optFechas 
            Caption         =   "Acumulado Anual "
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   375
            Width           =   1800
         End
         Begin VB.OptionButton optFechas 
            Caption         =   "Acumulado Rango Mes Corriente"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   90
            TabIndex        =   14
            Top             =   735
            Value           =   -1  'True
            Width           =   1845
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
            Left            =   1935
            TabIndex        =   4
            Top             =   225
            Width           =   3015
            Begin VB.CommandButton botHelpFH 
               Height          =   345
               Left            =   2535
               Picture         =   "frmIndicLS_T.frx":5B68
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2550
               Picture         =   "frmIndicLS_T.frx":5CDA
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
               _Version        =   327680
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
      Height          =   480
      Left            =   150
      TabIndex        =   39
      Top             =   5415
      Width           =   8730
   End
End
Attribute VB_Name = "frmIndicLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset
Dim RsDataAnt  As Recordset
Dim RsDataEstim  As Recordset
Dim RsDataVol  As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas
                          'con porcentajes
Dim UsoPorIntB As Boolean
Dim UsoPorAero As Boolean
Dim UsoPorTotal As Boolean
Private Function L_Armarcondicion(anio As Integer) As String
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant

If optFechas(0).Value Then
    If anio = Year(Date) Then
        fechaDesde = "01-01-" & Trim(str(anio))
        fechaHasta = Format$(Date - 1, "dd-mm") & "-" & Trim(str(anio))
    Else
        fechaDesde = "01-01-" & Trim(str(anio))
        fechaHasta = Format$(Date - 1, "dd-mm") & "-" & Trim(str(anio))
    End If
Else
    If anio = Year(mskFDesde.FormattedText) Then
        fechaDesde = mskFDesde.FormattedText
        fechaHasta = mskFHasta.FormattedText
    Else
        fechaDesde = mskFDesdeAnt.FormattedText
        fechaHasta = mskFHastaAnt.FormattedText
    End If
End If

Cond = " WHERE fch_vta between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

L_Armarcondicion = Cond

End Function

Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant
Dim AepImpL As Double, IntaImpL As Double, IntbImpL As Double
Dim AepImpS As Double, IntaImpS As Double, IntbImpS As Double
Dim IFLImpL As Double, IFLImpS As Double
Dim AepTickL As Long, IntaTickL As Long, IntbTickL As Long
Dim AepTickS As Long, IntaTickS As Long, IntbTickS As Long
Dim IFLTickL As Double, IFLTickS As Double

AepImpL = 0
IntaImpL = 0
IntbImpL = 0
IFLImpL = 0
AepImpS = 0
IntaImpS = 0
IntbImpS = 0
IFLImpS = 0

AepTickL = 0
IntaTickL = 0
IntbTickL = 0
IFLTickL = 0
AepTickS = 0
IntaTickS = 0
IntbTickS = 0
IFLTickS = 0

Do While Not RsData.EOF
    Select Case RsData!cod_depn
    Case DSLoc(1).Dep
        Select Case RsData!trancito
        Case "L"
            AepImpL = AepImpL + RsData!imp
            AepTickL = AepTickL + RsData!cant_t
        Case "S"
            AepImpS = AepImpS + RsData!imp
            AepTickS = AepTickS + RsData!cant_t
        End Select
    Case DSLoc(2).Dep
        Select Case RsData!cod_sdep
        Case DSLoc(2).Sdep
            Select Case RsData!trancito
            Case "L"
               IntaImpL = IntaImpL + RsData!imp
               IntaTickL = IntaTickL + RsData!cant_t
            Case "S"
               IntaImpS = IntaImpS + RsData!imp
               IntaTickS = IntaTickS + RsData!cant_t
            End Select
        Case DSLoc(3).Sdep
            Select Case RsData!trancito
            Case "L"
                IntbImpL = IntbImpL + RsData!imp
                IntbTickL = IntbTickL + RsData!cant_t
            Case "S"
                IntbImpS = IntbImpS + RsData!imp
                IntbTickS = IntbTickS + RsData!cant_t
            End Select
        End Select
        
    Case DSLoc(12).Dep
            Select Case RsData!trancito
            Case "L"
                IFLImpL = IFLImpL + RsData!imp
                IFLTickL = IFLTickL + RsData!cant_t
            Case "S"
                IFLImpS = IFLImpS + RsData!imp
                IFLTickS = IFLTickS + RsData!cant_t
            End Select
        
        
    End Select
            
    RsData.MoveNext

Loop

SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Facturación"
SprAep.SetText 2, SprAep.MaxRows, str(AepImpS)
SprAep.SetText 6, SprAep.MaxRows, str(AepImpL)

SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Tickets"
SprAep.SetText 2, SprAep.MaxRows, str(AepTickS)
SprAep.SetText 6, SprAep.MaxRows, str(AepTickL)


SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Pr. x Tickets"
If AepTickS > 0 Then
    SprAep.SetText 2, SprAep.MaxRows, Format$((AepImpS / AepTickS), "##.##")
    SprAep.SetText 6, SprAep.MaxRows, Format$((AepImpL / AepTickL), "##.##")
End If

SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Pax Viajados"

SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Pr x Pax Viaj."

SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Captación"


SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Facturación"
SprEzeA.SetText 2, SprEzeA.MaxRows, str(IntaImpS)
SprEzeA.SetText 6, SprEzeA.MaxRows, str(IntaImpL)

SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Tickets"
SprEzeA.SetText 2, SprEzeA.MaxRows, str(IntaTickS)
SprEzeA.SetText 6, SprEzeA.MaxRows, str(IntaTickL)


SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Pr. x Tickets"
If IntaTickS > 0 Then
SprEzeA.SetText 2, SprEzeA.MaxRows, Format$((IntaImpS / IntaTickS), "##.##")
SprEzeA.SetText 6, SprEzeA.MaxRows, Format$((IntaImpL / IntaTickL), "##.##")
End If


SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Pax Viajados"

SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Pr x Pax Viaj."

SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Captación"


SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Facturación"
SprEzeB.SetText 2, SprEzeB.MaxRows, str(IntbImpS)
SprEzeB.SetText 6, SprEzeB.MaxRows, str(IntbImpL)

SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Tickets"
SprEzeB.SetText 2, SprEzeB.MaxRows, str(IntbTickS)
SprEzeB.SetText 6, SprEzeB.MaxRows, str(IntbTickL)

SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Pr. x Tickets"
If IntbTickS > 0 Then
SprEzeB.SetText 2, SprEzeB.MaxRows, Format$((IntbImpS / IntbTickS), "##.##")
SprEzeB.SetText 6, SprEzeB.MaxRows, Format$((IntbImpL / IntbTickL), "##.##")
End If


SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, 6, "Pax Viajados"

SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Pr x Pax Viaj."

SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Captación"

'''''''
SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, SprFlight.MaxRows, "Facturación"
SprFlight.SetText 2, SprFlight.MaxRows, str(IFLImpS)
SprFlight.SetText 6, SprFlight.MaxRows, str(IFLImpL)

SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, SprFlight.MaxRows, "Tickets"
SprFlight.SetText 2, SprFlight.MaxRows, str(IFLTickS)
SprFlight.SetText 6, SprFlight.MaxRows, str(IFLTickL)

SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, SprFlight.MaxRows, "Pr. x Tickets"
If IFLTickS > 0 Then
SprFlight.SetText 2, SprFlight.MaxRows, Format$((IFLImpS / IFLTickS), "##.##")
SprFlight.SetText 6, SprFlight.MaxRows, Format$((IFLImpL / IFLTickL), "##.##")
End If


SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, 6, "Pax Viajados"

SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, SprFlight.MaxRows, "Pr x Pax Viaj."

SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, SprFlight.MaxRows, "Captación"


End Sub
Private Sub L_DecoEspigonAnt()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant
Dim AepImpL As Double, IntaImpL As Double, IntbImpL As Double, IFLImpL As Double
Dim AepImpS As Double, IntaImpS As Double, IntbImpS As Double, IFLImpS As Double
Dim AepTickL As Long, IntaTickL As Long, IntbTickL As Long, IFLTickL As Long
Dim AepTickS As Long, IntaTickS As Long, IntbTickS As Long, IFLTickS As Long

AepImpL = 0
IntaImpL = 0
IntbImpL = 0
IFLImpL = 0
AepImpS = 0
IntaImpS = 0
IntbImpS = 0
IFLImpS = 0
AepTickL = 0
IntaTickL = 0
IntbTickL = 0
IFLTickL = 0
AepTickS = 0
IntaTickS = 0
IntbTickS = 0
IFLTickS = 0

Do While Not RsDataAnt.EOF
    Select Case RsDataAnt!cod_depn
    Case DSLoc(1).Dep
        Select Case RsDataAnt!trancito
        Case "L"
            AepImpL = AepImpL + RsDataAnt!imp
            AepTickL = AepTickL + RsDataAnt!cant_t
        Case "S"
            AepImpS = AepImpS + RsDataAnt!imp
            AepTickS = AepTickS + RsDataAnt!cant_t
        End Select
    Case DSLoc(2).Dep
        Select Case RsDataAnt!cod_sdep
        Case DSLoc(2).Sdep
            Select Case RsDataAnt!trancito
            Case "L"
               IntaImpL = IntaImpL + RsDataAnt!imp
               IntaTickL = IntaTickL + RsDataAnt!cant_t
            Case "S"
               IntaImpS = IntaImpS + RsDataAnt!imp
               IntaTickS = IntaTickS + RsDataAnt!cant_t
            End Select
        Case DSLoc(3).Sdep
            Select Case RsDataAnt!trancito
            Case "L"
                IntbImpL = IntbImpL + RsDataAnt!imp
                IntbTickL = IntbTickL + RsDataAnt!cant_t
            Case "S"
                IntbImpS = IntbImpS + RsDataAnt!imp
                IntbTickS = IntbTickS + RsDataAnt!cant_t
            End Select
        End Select
    Case DSLoc(2).Dep
            Select Case RsDataAnt!trancito
            Case "L"
                IFLImpL = IFLImpL + RsDataAnt!imp
                IFLTickL = IFLTickL + RsDataAnt!cant_t
            Case "S"
                IFLImpS = IFLImpS + RsDataAnt!imp
                IFLTickS = IFLTickS + RsDataAnt!cant_t
            End Select
        
    End Select
            
    RsDataAnt.MoveNext

Loop

SprAep.SetText 3, 1, str(AepImpS)
SprAep.SetText 7, 1, str(AepImpL)

SprAep.SetText 3, 2, str(AepTickS)
SprAep.SetText 7, 2, str(AepTickL)

If AepImpS = 0 Or AepImpL = 0 Or AepTickS = 0 Or AepTickL = 0 Then
        SprAep.SetText 3, 3, Format$(0, "##.##")
        SprAep.SetText 7, 3, Format$(0, "##.##")
  Else
        SprAep.SetText 3, 3, Format$((AepImpS / AepTickS), "##.##")
        SprAep.SetText 7, 3, Format$((AepImpL / AepTickL), "##.##")
End If


SprEzeA.SetText 3, 1, str(IntaImpS)
SprEzeA.SetText 7, 1, str(IntaImpL)

SprEzeA.SetText 3, 2, str(IntaTickS)
SprEzeA.SetText 7, 2, str(IntaTickL)

If IntaImpS = 0 Or IntaImpL = 0 Or IntaTickS = 0 Or IntaTickL = 0 Then
    SprEzeA.SetText 3, 3, Format$(0, "##.##")
    SprEzeA.SetText 7, 3, Format$(0, "##.##")
  Else
    SprEzeA.SetText 3, 3, Format$((IntaImpS / IntaTickS), "##.##")
    SprEzeA.SetText 7, 3, Format$((IntaImpL / IntaTickL), "##.##")
End If

SprEzeB.SetText 3, 1, str(IntbImpS)
SprEzeB.SetText 7, 1, str(IntbImpL)

SprEzeB.SetText 3, 2, str(IntbTickS)
SprEzeB.SetText 7, 2, str(IntbTickL)

If IntbImpS = 0 Or IntbImpL = 0 Or IntbTickS = 0 Or IntbTickL = 0 Then
        SprEzeB.SetText 3, 3, Format$(0, "##.##")
        SprEzeB.SetText 7, 3, Format$(0, "##.##")
     Else
        SprEzeB.SetText 3, 3, Format$((IntbImpS / IntbTickS), "##.##")
        SprEzeB.SetText 7, 3, Format$((IntbImpL / IntbTickL), "##.##")
End If

SprFlight.SetText 3, 1, str(IFLImpS)
SprFlight.SetText 7, 1, str(IFLImpL)

SprFlight.SetText 3, 2, str(IFLTickS)
SprFlight.SetText 7, 2, str(IFLTickL)

If IFLImpS = 0 Or IFLImpL = 0 Or IFLTickS = 0 Or IFLTickL = 0 Then
        SprFlight.SetText 3, 3, Format$(0, "##.##")
        SprFlight.SetText 7, 3, Format$(0, "##.##")
   Else
        SprFlight.SetText 3, 3, Format$((IFLImpS / IFLTickS), "##.##")
        SprFlight.SetText 7, 3, Format$((IFLImpL / IFLTickL), "##.##")
End If

End Sub
Private Sub L_DecoVolados(col As Integer)
Dim valor As Variant

Do While Not RsDataVol.EOF
        Select Case RsDataVol!cod_depn
            Case DSLoc(1).Dep
                If RsDataVol!Tipo = "S" Then
                   SprAep.SetText col, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprAep.GetText col, 1, valor
                      SprAep.SetText col, 5, str(valor / RsDataVol!cant_v)
                      SprAep.GetText col, 2, valor
                      SprAep.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
                Else
                   SprAep.SetText col + 4, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprAep.GetText col + 4, 1, valor
                      SprAep.SetText col + 4, 5, str(valor / RsDataVol!cant_v)
                      SprAep.GetText col + 4, 2, valor
                      SprAep.SetText col + 4, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
                
                End If
            Case DSLoc(2).Dep
                Select Case RsDataVol!cod_sdep
                    Case DSLoc(2).Sdep
                        If RsDataVol!Tipo = "S" Then
                            SprEzeA.SetText col, 4, str(RsDataVol!cant_v)
                            If RsDataVol!cant_v > 0 Then
                                SprEzeA.GetText col, 1, valor
                                SprEzeA.SetText col, 5, str(valor / RsDataVol!cant_v)
                                SprEzeA.GetText col, 2, valor
                                SprEzeA.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        Else
                            SprEzeA.SetText col + 4, 4, str(RsDataVol!cant_v)
                            If RsDataVol!cant_v > 0 Then
                                SprEzeA.GetText col + 4, 1, valor
                                SprEzeA.SetText col + 4, 5, str(valor / RsDataVol!cant_v)
                                SprEzeA.GetText col + 4, 2, valor
                                SprEzeA.SetText col + 4, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        
                        End If
                    Case DSLoc(3).Sdep
                        If RsDataVol!Tipo = "S" Then
                            SprEzeB.SetText col, 4, str(RsDataVol!cant_v)
                            If RsDataVol!cant_v > 0 Then
                                SprEzeB.GetText col, 1, valor
                                SprEzeB.SetText col, 5, (valor / RsDataVol!cant_v)
                                SprEzeB.GetText col, 2, valor
                                SprEzeB.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        Else
                            SprEzeB.SetText col + 4, 4, str(RsDataVol!cant_v)
                            If RsDataVol!cant_v > 0 Then
                                SprEzeB.GetText col + 4, 1, valor
                                SprEzeB.SetText col + 4, 5, (valor / RsDataVol!cant_v)
                                SprEzeB.GetText col + 4, 2, valor
                                SprEzeB.SetText col + 4, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        
                        End If
                End Select
                
            Case DSLoc(12).Dep
                If RsDataVol!Tipo = "S" Then
                   SprFlight.SetText col, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprFlight.GetText col, 1, valor
                      SprFlight.SetText col, 5, (valor / RsDataVol!cant_v)
                      SprFlight.GetText col, 2, valor
                      SprFlight.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
                Else
                   SprFlight.SetText col + 4, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprFlight.GetText col + 4, 1, valor
                      SprFlight.SetText col + 4, 5, (valor / RsDataVol!cant_v)
                      SprFlight.GetText col + 4, 2, valor
                      SprFlight.SetText col + 4, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
                
                End If
                
        End Select
                
        RsDataVol.MoveNext
Loop
End Sub



Private Sub L_LimpiarGrillas()
Dim i
    
    SprEzeA.MaxRows = 0
    SprEzeB.MaxRows = 0
    SprAep.MaxRows = 0
    SprTotal.MaxRows = 0
    SprFlight.MaxRows = 0
    SprEZEIZA.MaxRows = 0
    labInfo.caption = ""
    
End Sub

Private Sub L_LlenarTotales()
Dim Nom As String
Dim i As Integer

For i = 1 To 6
    SprTotal.MaxRows = SprTotal.MaxRows + 1
    Select Case i
        Case 1
            Nom = "Facturación"
        Case 2
            Nom = "Tickets"
        Case 3
            Nom = "Pr. x Tickets"
        Case 4
            Nom = "Pax Viajados"
        Case 5
            Nom = "Pr x Pax Viaj."
        Case 6
            Nom = "Captación"
    End Select
    
    SprTotal.SetText 1, i, Nom

        L_RellenarDatos i, 2, Nom
        L_RellenarDatos i, 3, Nom
        L_RellenarDatos i, 6, Nom
        L_RellenarDatos i, 7, Nom
Next
End Sub

Private Sub L_LlenarTotales_EZE()
Dim Nom As String
Dim i As Integer

For i = 1 To 6
    SprEZEIZA.MaxRows = SprEZEIZA.MaxRows + 1
    Select Case i
        Case 1
            Nom = "Facturación"
        Case 2
            Nom = "Tickets"
        Case 3
            Nom = "Pr. x Tickets"
        Case 4
            Nom = "Pax Viajados"
        Case 5
            Nom = "Pr x Pax Viaj."
        Case 6
            Nom = "Captación"
    End Select
    
    SprEZEIZA.SetText 1, i, Nom

        L_RellenarDatos_EZE i, 2, Nom
        L_RellenarDatos_EZE i, 3, Nom
        L_RellenarDatos_EZE i, 6, Nom
        L_RellenarDatos_EZE i, 7, Nom
Next
End Sub

Private Sub L_Refrescar()
Dim SQL As String
Dim sqlX As String
Dim rs As Recordset

On Error GoTo ErrInd:

frmIndicLS.caption = Aplicacion.SeteoProceso(frmIndicLS.caption)


SQL = " SELECT "
SQL = SQL & " PL.cod_depn,"
SQL = SQL & " DECODE(PL.cod_depn,'IFL','IFL',PL.cod_sdep) cod_sdep, "
SQL = SQL & " LT.TRANCITO, "
SQL = SQL & " sum(cant_tickets) cant_t, "
SQL = SQL & " sum(importe) imp "
SQL = SQL & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText))) & " PL ,"
SQL = SQL & " ESTADIS.Local_Trancito  LT"
SQL = SQL & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)))
SQL = SQL & " and PL.Cod_depn = LT.Cod_depn "
SQL = SQL & " and PL.Cod_sdep = LT.cod_sdep "
SQL = SQL & " and PL.COD_LOCAL=LT.COD_LOCAL "
SQL = SQL & " and TIPO_VUELO = 'PI' "
SQL = SQL & "group by PL.cod_depn,DECODE(PL.cod_depn,'IFL','IFL',PL.cod_sdep),LT.TRANCITO,LT.TIPO_VUELO "
'sql = sql & " order by cod_depn,cod_sdep"
SQL = SQL & " UNION "
SQL = SQL & " SELECT "
SQL = SQL & " PL.cod_depn, "
SQL = SQL & " DECODE(PL.cod_depn,'IFL','IFL',PL.cod_sdep) cod_sdep, "
SQL = SQL & " LT.TRANCITO, "
SQL = SQL & " sum(cant_tickets) cant_t, "
SQL = SQL & " sum(importe) imp "
SQL = SQL & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText))) & " PL ,"
SQL = SQL & " ESTADIS.Local_Trancito  LT"
SQL = SQL & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)))
SQL = SQL & " and PL.Cod_depn = LT.Cod_depn "
SQL = SQL & " and PL.Cod_sdep = LT.cod_sdep "
SQL = SQL & " and PL.COD_LOCAL=LT.COD_LOCAL "
SQL = SQL & " and TIPO_VUELO <> 'PI' "
SQL = SQL & " and  decode(mod(pl.cod_vuelo,2),0,'S','L') = LT.TRANCITO "
SQL = SQL & "group by PL.cod_depn,DECODE(PL.cod_depn,'IFL','IFL',PL.cod_sdep),LT.TRANCITO,LT.TIPO_VUELO, mod(pl.cod_vuelo,2) "


sqlX = " SELECT "
sqlX = sqlX & " PL.cod_depn,"
sqlX = sqlX & " DECODE(PL.cod_depn,'IFL','IFL',PL.cod_sdep) cod_sdep , "
sqlX = sqlX & " LT.TRANCITO, "
sqlX = sqlX & " sum(cant_tickets) cant_t, "
sqlX = sqlX & " sum(importe) imp "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)) - 1) & " PL ,"
sqlX = sqlX & " ESTADIS.Local_Trancito  LT"
sqlX = sqlX & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)) - 1)
sqlX = sqlX & " and PL.Cod_depn = LT.Cod_depn "
sqlX = sqlX & " and PL.Cod_sdep = LT.cod_sdep "
sqlX = sqlX & " and PL.COD_LOCAL=LT.COD_LOCAL "
sqlX = sqlX & " and TIPO_VUELO = 'PI' "
sqlX = sqlX & "group by PL.cod_depn,DECODE(PL.cod_depn,'IFL','IFL',PL.cod_sdep), LT.TRANCITO,LT.TIPO_VUELO "
'sqlX = sqlX & " order by cod_depn,cod_sdep"
sqlX = sqlX & " UNION "
sqlX = sqlX & " SELECT "
sqlX = sqlX & " PL.cod_depn , "
sqlX = sqlX & " DECODE(PL.cod_depn,'IFL','IFL',PL.cod_sdep) cod_sdep , "
sqlX = sqlX & " LT.TRANCITO, "
sqlX = sqlX & " sum(cant_tickets) cant_t, "
sqlX = sqlX & " sum(importe) imp "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)) - 1) & " PL ,"
sqlX = sqlX & " ESTADIS.Local_Trancito  LT"
sqlX = sqlX & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)) - 1)
sqlX = sqlX & " and PL.Cod_depn = LT.Cod_depn "
sqlX = sqlX & " and PL.Cod_sdep = LT.cod_sdep "
sqlX = sqlX & " and PL.COD_LOCAL=LT.COD_LOCAL "
sqlX = sqlX & " and TIPO_VUELO <> 'PI' "
sqlX = sqlX & " and  decode(mod(pl.cod_vuelo,2),0,'S','L') = LT.TRANCITO "
sqlX = sqlX & "group by PL.cod_depn, DECODE(PL.cod_depn,'IFL','IFL',PL.cod_sdep),LT.TRANCITO,LT.TIPO_VUELO, mod(pl.cod_vuelo,2) "


If Aplicacion.ObtenerRsDAO(SQL, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        TabEspigon.Enabled = True
        L_DecoEspigon
        If Aplicacion.ObtenerRsDAO(sqlX, RsDataAnt) Then
            L_DecoEspigonAnt
        End If
        L_TratarVolados
        L_TratarVoladosHIST
        L_LlenarTotales
        L_LlenarTotales_EZE
    
        L_Resaltar
        
        Aplicacion.CerrarDAO RsDataAnt
    End If
    
    Aplicacion.CerrarDAO RsData

End If

ErrInd:
    frmIndicLS.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub


Private Sub L_RellenarDatos(i As Integer, col As Integer, Nom As String)
Dim valor As Variant, divsor As Variant
Dim Impr As Double
    
    If i = 1 Or i = 2 Or i = 4 Then
        SprEzeA.GetText col, i, valor
        Impr = Impr + valor
        SprEzeB.GetText col, i, valor
        Impr = Impr + valor
        SprAep.GetText col, i, valor
        Impr = Impr + valor
        If i <> 4 Then
          SprFlight.GetText col, i, valor
          Impr = Impr + valor
        End If
    ElseIf i = 3 Then
        SprTotal.GetText col, 1, valor
        SprTotal.GetText col, 2, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 5 Then
        SprTotal.GetText col, 1, valor
        SprTotal.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 6 Then
        SprTotal.GetText col, 2, valor
        SprTotal.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor * 100
        End If
    
    End If
    
    SprTotal.SetText col, i, Format$((Impr), "#.00")

End Sub

Private Sub L_RellenarDatos_EZE(i As Integer, col As Integer, Nom As String)
Dim valor As Variant, divsor As Variant
Dim Impr As Double
    
    If i = 1 Or i = 2 Or i = 4 Then
        SprEzeA.GetText col, i, valor
        Impr = Impr + valor
        SprEzeB.GetText col, i, valor
        Impr = Impr + valor
        SprAep.GetText col, i, valor
        Impr = Impr + valor
    ElseIf i = 3 Then
        SprEZEIZA.GetText col, 1, valor
        SprEZEIZA.GetText col, 2, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 5 Then
        SprEZEIZA.GetText col, 1, valor
        SprEZEIZA.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 6 Then
        SprEZEIZA.GetText col, 2, valor
        SprEZEIZA.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor * 100
        End If
    
    End If
    
    SprEZEIZA.SetText col, i, Format$((Impr), "#.00")

End Sub
Private Sub L_Resaltar()
Dim fila As Integer

For fila = 1 To SprTotal.MaxRows
    Spread.spread_ResaltarCelda SprTotal, 4, fila
    Spread.spread_ResaltarCelda SprTotal, 5, fila
    Spread.spread_ResaltarCelda SprTotal, 8, fila
    Spread.spread_ResaltarCelda SprTotal, 9, fila
    
    Spread.spread_ResaltarCelda SprEzeA, 4, fila
    Spread.spread_ResaltarCelda SprEzeA, 5, fila
    Spread.spread_ResaltarCelda SprEzeA, 8, fila
    Spread.spread_ResaltarCelda SprEzeA, 9, fila

    Spread.spread_ResaltarCelda SprEzeB, 4, fila
    Spread.spread_ResaltarCelda SprEzeB, 5, fila
    Spread.spread_ResaltarCelda SprEzeB, 8, fila
    Spread.spread_ResaltarCelda SprEzeB, 9, fila

    Spread.spread_ResaltarCelda SprAep, 4, fila
    Spread.spread_ResaltarCelda SprAep, 5, fila
    Spread.spread_ResaltarCelda SprAep, 8, fila
    Spread.spread_ResaltarCelda SprAep, 9, fila
    
    Spread.spread_ResaltarCelda SprFlight, 4, fila
    Spread.spread_ResaltarCelda SprFlight, 5, fila
    Spread.spread_ResaltarCelda SprFlight, 8, fila
    Spread.spread_ResaltarCelda SprFlight, 9, fila

    Spread.spread_ResaltarCelda SprEZEIZA, 4, fila
    Spread.spread_ResaltarCelda SprEZEIZA, 5, fila
    Spread.spread_ResaltarCelda SprEZEIZA, 8, fila
    Spread.spread_ResaltarCelda SprEZEIZA, 9, fila
    
Next

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
        Case total
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
        Case total
            UsoPorTotal = False
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorAero Or UsoPorTotal) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub


Private Sub L_TratarVolados()
Dim SQL As String
Dim FD As String, FH As String
Dim rs_pax As Recordset

If optFechas(0).Value Then
    FD = "01-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText)
    FH = mskFDesde.FormattedText
Else
    FD = mskFDesde.FormattedText
    FH = mskFHasta.FormattedText
End If

SQL = "select min(fch_actualizado) Fch from estadis.control_carga WHERE cod_sdep <> 'INT' "

If Aplicacion.ObtenerRsDAO(SQL, rs_pax) Then
    
    If CDate(FH) <= CDate(rs_pax!fch) Then
  
        SQL = " SELECT "
        SQL = SQL & "tipo,"
        SQL = SQL & "cod_depn,"
        SQL = SQL & "DECODE(COD_DEPN,'IFL','IFL',COD_SDEP) COD_SDEP,"
        SQL = SQL & " SUM(pax_AT_VS_vol.PAX_VOL) CANT_V "
        SQL = SQL & " From estadis.pax_AT_VS_vol "
        SQL = SQL & " Where "
        SQL = SQL & " fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
        SQL = SQL & " and tipo <> 'I' "
        SQL = SQL & " GROUP BY COD_DEPN,DECODE(cod_depN,'IFL','IFL',COD_SDEP),tipo"
        
        If Aplicacion.ObtenerRsDAO(SQL, RsDataVol) Then
            L_DecoVolados 2
            Aplicacion.CerrarDAO RsDataVol
        End If
    Else
        labInfo.caption = "Datos de pasajeros transitados Actualizados al día " & rs_pax!fch
    End If
End If


End Sub

Private Sub L_TratarVoladosHIST()
Dim SQL As String
Dim FD As String, FH As String

If optFechas(0).Value Then
    FD = "01-01-" & Year(mskFDesde.FormattedText) - 1
    FH = Day(mskFDesde.FormattedText) & "-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText) - 1
Else
    FD = mskFDesdeAnt.FormattedText
    FH = mskFHastaAnt.FormattedText
End If

        SQL = " SELECT "
        SQL = SQL & "tipo,"
        SQL = SQL & "cod_depn,"
        SQL = SQL & "cod_sdep,"
        SQL = SQL & " SUM(pax_volados.cantidad) CANT_V "
        SQL = SQL & " From estadis.pax_volados "
        SQL = SQL & " Where "
        SQL = SQL & " fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
        SQL = SQL & " GROUP BY COD_DEPN,COD_SDEP,tipo"
        
        If Aplicacion.ObtenerRsDAO(SQL, RsDataVol) Then
            L_DecoVolados 3
            Aplicacion.CerrarDAO RsDataVol
        End If

End Sub


Private Sub borGrBar_Click(Index As Integer)
Dim col_datos As Collection

    Set col_datos = New Collection
    Select Case Index
        Case 0
            L_LLenarColeccion col_datos, SprTotal
            FrmGrafbar.CargarGrafico total, "", col_datos
        Case 1
            L_LLenarColeccion col_datos, SprEzeA
            FrmGrafbar.CargarGrafico EZEA, "", col_datos
        Case 2
            L_LLenarColeccion col_datos, SprEzeB
            FrmGrafbar.CargarGrafico EZEB, "", col_datos
        Case 3
            L_LLenarColeccion col_datos, SprAep
            FrmGrafbar.CargarGrafico AERO, "", col_datos
    End Select
        
    
End Sub

Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        L_Refrescar
    Case 1

        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        TabEspigon.Enabled = False
        L_LimpiarGrillas
        
    Case 2
        Unload Me
End Select
End Sub



Private Sub botExcel_Click()
    If optFechas(0).Value Then
        L_TratarExcel SprTotal, "INDICADORES ACUMULADOS ANUALES (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
        "del 1 de Ene al " & Format$(mskFDesde.FormattedText, "dd-mmm"), "", Exl_Blanco
    Else
        L_TratarExcel SprTotal, "INDICADORES ACUMULADOS MES CORRIENTE (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
        " del " & Format$(mskFDesde.FormattedText, "dd-mmm") & " al " & Format$(mskFHasta.FormattedText, "dd-mmm"), "", Exl_Blanco 'Exl_Ros
    End If
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

Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabEspigon
TabEspigon.TabEnabled(4) = False

End Sub

Private Sub L_LLenarColeccion(ByRef col As Collection, spr As control)
Dim cl_dato As CLlgi
Dim valor As Variant
Dim item As Variant

spr.Row = spr.ActiveRow
spr.GetText 1, spr.Row, item

Set cl_dato = New CLlgi
    
spr.GetText 2, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
col.Add cl_dato
    
Set cl_dato = New CLlgi
spr.GetText 3, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
col.Add cl_dato

Set cl_dato = New CLlgi
spr.GetText 6, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
col.Add cl_dato

End Sub


Private Sub L_TratarExcel(spr As control, Titulo As String, subTit As String, Info As String, color As Integer)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim fila As Integer
Dim i
Dim tit As Variant
Dim nombre As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

frmIndicLS.caption = Aplicacion.SeteoProceso(frmIndicLS.caption)

If nombre <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
'    AppExcel.application.Visible = True
    
    ReDim titCol(spr.MaxCols)
    col = 1
    fila = 3
    
    For i = 1 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 1, 1, Titulo
    rango = Exl_rangos(1, 1, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, color
    
    Exl_PonerValor AppExcel, fila, col, subTit
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    fila = fila + 1
    Exl_PonerValor AppExcel, fila, col, "Total Buenos Aires"
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba
    
    AppExcel.Application.Range(rango).Merge
        
    fila = fila + 1
    Exl_PonerValor AppExcel, fila, col + 1, "SALIDAS"
    Exl_PonerValor AppExcel, fila, col + 5, "LLEGADAS"
    fila = fila + 1
    Exl_BajarGrillaExel spr, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + spr.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango

    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "INTERNACIONAL A "
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_PonerValor AppExcel, fila, col + 1, "SALIDAS"
    Exl_PonerValor AppExcel, fila, col + 5, "LLEGADAS"
    fila = fila + 1
    Exl_BajarGrillaExel SprEzeA, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprEzeA.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '-----------------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "INTERNACIONAL B "
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_PonerValor AppExcel, fila, col + 1, "SALIDAS"
    Exl_PonerValor AppExcel, fila, col + 5, "LLEGADAS"
    fila = fila + 1
    Exl_BajarGrillaExel SprEzeB, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprEzeB.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "AEROPARQUE "
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba
    
    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_PonerValor AppExcel, fila, col + 1, "SALIDAS"
    Exl_PonerValor AppExcel, fila, col + 5, "LLEGADAS"
    fila = fila + 1
    Exl_BajarGrillaExel SprAep, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprAep.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
        
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris

    Exl_AnchoCol AppExcel, 5, 5, 7
    Exl_AnchoCol AppExcel, 9, 9, 7
    
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    'AppExcel.application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen

    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If

    AppExcel.SaveAs nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmIndicLS.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6300
Me.Width = 9300

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

frmPrincipal.lstForms.AddItem "frmIndic"

End Sub


Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmIndic"
End Sub


Private Sub mskFDesde_LostFocus()

    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    Else
    'If (Year(mskFDesde.FormattedText) < Year(Date)) Then
    '(Month(mskFDesde.FormattedText) < Month(Date))
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
        mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
    Else
    'If Year(CDate(mskFDesdeAnt.FormattedText)) >= Year(Date) Then
    '   mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
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

Private Sub optFechas_Click(Index As Integer)
Select Case Index
    Case 0
        mskFHasta.Visible = False
        botHelpFH.Visible = False
        Label1(0).caption = "Fecha Tope"
        Label1(1).Visible = False
        frPerAnt.Visible = False
    Case 1
        mskFHasta.Visible = True
        botHelpFH.Visible = True
        Label1(0).caption = "Fecha Desde"
        Label1(1).Visible = True
        frPerAnt.Visible = True
End Select
End Sub

Private Sub tabEspigon_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    Select Case TabEspigon.Tab
        Case 0
            SprTotal.SetFocus
        Case 1
            SprEzeA.SetFocus
        Case 2
            SprEzeB.SetFocus
        Case 3
            SprAep.SetFocus
        Case 4
            SprFlight.SetFocus
    End Select
    
    
ErrT:
    Exit Sub



End Sub

Private Sub Timer1_Timer()
labInfo.Visible = Not labInfo.Visible
End Sub


