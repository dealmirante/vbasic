VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCaj 
   Caption         =   "Ventas por Cajeros"
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
      Left            =   90
      TabIndex        =   4
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
      TabPicture(0)   =   "frmCaj.frx":0000
      Tab(0).ControlCount=   4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frTotPorc0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frTot0"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "botPorcTot(2)"
      Tab(0).Control(2).Enabled=   -1  'True
      Tab(0).Control(3)=   "botPorcTot(0)"
      Tab(0).Control(3).Enabled=   -1  'True
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmCaj.frx":001C
      Tab(1).ControlCount=   6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frEzeAPorc0"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frEzeA0"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "botExcelEzeA"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "botPorc(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "botPorc(0)"
      Tab(1).Control(5).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmCaj.frx":0038
      Tab(2).ControlCount=   6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frEzeBPorc0"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frEzeB0"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame5"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "botExcelEzeB"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "botPorcB(2)"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "botPorcB(0)"
      Tab(2).Control(5).Enabled=   -1  'True
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmCaj.frx":0054
      Tab(3).ControlCount=   6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frAepPorc0"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "frAep0"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame6"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "botExcelAep"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "botPorcAep(2)"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "botPorcAep(0)"
      Tab(3).Control(5).Enabled=   -1  'True
      TabCaption(4)   =   "INTERIOR"
      TabPicture(4)   =   "frmCaj.frx":0070
      Tab(4).ControlCount=   6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrIntPorc0"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "FrInt0"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "botPorcInt(0)"
      Tab(4).Control(2).Enabled=   -1  'True
      Tab(4).Control(3)=   "botPorcInt(2)"
      Tab(4).Control(3).Enabled=   -1  'True
      Tab(4).Control(4)=   "botExcelInt"
      Tab(4).Control(4).Enabled=   -1  'True
      Tab(4).Control(5)=   "Frame7"
      Tab(4).Control(5).Enabled=   0   'False
      Begin VB.Frame Frame7 
         Caption         =   "Ordenamiento"
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
         Height          =   1560
         Left            =   -67845
         TabIndex        =   66
         Top             =   2115
         Width           =   1650
         Begin VB.OptionButton optSort 
            Caption         =   "Por Imp. (Desc)"
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
            Index           =   14
            Left            =   60
            TabIndex        =   70
            Top             =   945
            Width           =   1530
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Código"
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
            Index           =   12
            Left            =   60
            TabIndex        =   69
            Top             =   270
            Width           =   1170
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Imp. (Asc)"
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
            Index           =   13
            Left            =   60
            TabIndex        =   68
            Top             =   600
            Width           =   1470
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Prom."
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
            Index           =   15
            Left            =   60
            TabIndex        =   67
            Top             =   1260
            Width           =   1530
         End
      End
      Begin VB.CommandButton botExcelInt 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67215
         Picture         =   "frmCaj.frx":008C
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1545
         Width           =   780
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
         Left            =   -67200
         Picture         =   "frmCaj.frx":061E
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1005
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
         Left            =   -67200
         Picture         =   "frmCaj.frx":0CE4
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   465
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
         Left            =   -67290
         Picture         =   "frmCaj.frx":1656
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   450
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
         Index           =   2
         Left            =   -67275
         Picture         =   "frmCaj.frx":1FC8
         Style           =   1  'Graphical
         TabIndex        =   54
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
         Left            =   -67170
         Picture         =   "frmCaj.frx":268E
         Style           =   1  'Graphical
         TabIndex        =   53
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
         Left            =   -67170
         Picture         =   "frmCaj.frx":3000
         Style           =   1  'Graphical
         TabIndex        =   52
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
         Left            =   7770
         Picture         =   "frmCaj.frx":36C6
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   420
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
         Left            =   7770
         Picture         =   "frmCaj.frx":4038
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   945
         Width           =   765
      End
      Begin VB.CommandButton botExcelAep 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67290
         Picture         =   "frmCaj.frx":46FE
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1515
         Width           =   780
      End
      Begin VB.CommandButton botExcelEzeA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   7770
         Picture         =   "frmCaj.frx":4C90
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1485
         Width           =   765
      End
      Begin VB.CommandButton botExcelEzeB 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67170
         Picture         =   "frmCaj.frx":5222
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1515
         Width           =   780
      End
      Begin VB.Frame Frame6 
         Caption         =   "Ordenamiento"
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
         Height          =   1590
         Left            =   -67965
         TabIndex        =   40
         Top             =   2070
         Width           =   1650
         Begin VB.OptionButton optSort 
            Caption         =   "Por Prom."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   75
            TabIndex        =   58
            Top             =   1260
            Width           =   1455
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Imp. (Desc)"
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
            Index           =   10
            Left            =   75
            TabIndex        =   43
            Top             =   930
            Width           =   1530
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Código"
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
            Index           =   8
            Left            =   90
            TabIndex        =   42
            Top             =   225
            Width           =   1170
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Imp. (Asc)"
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
            Index           =   9
            Left            =   75
            TabIndex        =   41
            Top             =   570
            Width           =   1470
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ordenamiento"
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
         Height          =   1560
         Left            =   -67950
         TabIndex        =   36
         Top             =   2070
         Width           =   1650
         Begin VB.OptionButton optSort 
            Caption         =   "Por Prom."
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
            Index           =   7
            Left            =   30
            TabIndex        =   57
            Top             =   1260
            Width           =   1530
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Imp. (Asc)"
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
            Index           =   5
            Left            =   45
            TabIndex        =   39
            Top             =   600
            Width           =   1470
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Código"
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
            Index           =   4
            Left            =   30
            TabIndex        =   38
            Top             =   240
            Width           =   1170
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Imp. (Desc)"
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
            Index           =   6
            Left            =   30
            TabIndex        =   37
            Top             =   945
            Width           =   1530
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Ordenamiento"
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
         Height          =   1635
         Left            =   7140
         TabIndex        =   32
         Top             =   1995
         Width           =   1620
         Begin VB.OptionButton optSort 
            Caption         =   "Por Prom."
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
            Index           =   3
            Left            =   30
            TabIndex        =   56
            Top             =   1290
            Width           =   1350
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Imp. (Desc)"
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
            Left            =   30
            TabIndex        =   35
            Top             =   960
            Width           =   1530
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Código"
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
            Left            =   30
            TabIndex        =   34
            Top             =   240
            Width           =   1170
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Imp. (Asc)"
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
            Left            =   30
            TabIndex        =   33
            Top             =   585
            Width           =   1470
         End
      End
      Begin VB.Frame frTotPorc0 
         Height          =   3015
         Left            =   -71220
         TabIndex        =   18
         Top             =   435
         Width           =   2655
         Begin FPSpread.vaSpread sprTotPorc 
            Height          =   2700
            Index           =   0
            Left            =   90
            OleObjectBlob   =   "frmCaj.frx":57B4
            TabIndex        =   19
            Top             =   195
            Width           =   6600
         End
      End
      Begin VB.Frame frTot0 
         Height          =   3000
         Left            =   -74760
         TabIndex        =   16
         Top             =   435
         Width           =   3150
         Begin FPSpread.vaSpread sprTot 
            Height          =   2700
            Index           =   0
            Left            =   90
            OleObjectBlob   =   "frmCaj.frx":5B89
            TabIndex        =   17
            Top             =   210
            Width           =   6600
         End
      End
      Begin VB.CommandButton botPorcTot 
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
         Height          =   465
         Index           =   2
         Left            =   -67185
         TabIndex        =   6
         Top             =   1515
         Width           =   885
      End
      Begin VB.CommandButton botPorcTot 
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
         Height          =   450
         Index           =   0
         Left            =   -67185
         TabIndex        =   5
         Top             =   435
         Width           =   885
      End
      Begin VB.Frame frEzeB0 
         Height          =   3000
         Left            =   -74850
         TabIndex        =   24
         Top             =   500
         Width           =   6800
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2700
            Left            =   105
            OleObjectBlob   =   "frmCaj.frx":5FD6
            TabIndex        =   28
            Top             =   195
            Width           =   6600
         End
      End
      Begin VB.Frame frEzeBPorc0 
         Height          =   3000
         Left            =   -74850
         TabIndex        =   25
         Top             =   500
         Width           =   6800
         Begin FPSpread.vaSpread sprEzeBPorc 
            Height          =   2700
            Left            =   90
            OleObjectBlob   =   "frmCaj.frx":6557
            TabIndex        =   29
            Top             =   180
            Width           =   6600
         End
      End
      Begin VB.Frame frEzeA0 
         Height          =   3000
         Left            =   150
         TabIndex        =   20
         Top             =   500
         Width           =   6800
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2700
            Left            =   105
            OleObjectBlob   =   "frmCaj.frx":692C
            TabIndex        =   21
            Top             =   180
            Width           =   6600
         End
      End
      Begin VB.Frame frEzeAPorc0 
         Height          =   3000
         Left            =   150
         TabIndex        =   22
         Top             =   495
         Width           =   6800
         Begin FPSpread.vaSpread sprEzeAPorc 
            Height          =   2700
            Left            =   75
            OleObjectBlob   =   "frmCaj.frx":6E85
            TabIndex        =   23
            Top             =   195
            Width           =   6600
         End
      End
      Begin VB.Frame FrInt0 
         Height          =   3000
         Left            =   -74835
         TabIndex        =   61
         Top             =   405
         Width           =   6800
         Begin FPSpread.vaSpread SprInt 
            Height          =   2700
            Left            =   105
            OleObjectBlob   =   "frmCaj.frx":725A
            TabIndex        =   62
            Top             =   225
            Width           =   6600
         End
      End
      Begin VB.Frame FrIntPorc0 
         Height          =   2985
         Left            =   -74835
         TabIndex        =   59
         Top             =   420
         Width           =   6795
         Begin FPSpread.vaSpread sprIntPorc 
            Height          =   2700
            Left            =   90
            OleObjectBlob   =   "frmCaj.frx":77CF
            TabIndex        =   60
            Top             =   195
            Width           =   6615
         End
      End
      Begin VB.Frame frAep0 
         Height          =   3000
         Left            =   -74850
         TabIndex        =   26
         Top             =   500
         Width           =   6800
         Begin FPSpread.vaSpread sprAep 
            Height          =   2700
            Left            =   105
            OleObjectBlob   =   "frmCaj.frx":7BA4
            TabIndex        =   30
            Top             =   225
            Width           =   6600
         End
      End
      Begin VB.Frame frAepPorc0 
         Height          =   2985
         Left            =   -74850
         TabIndex        =   27
         Top             =   500
         Width           =   6795
         Begin FPSpread.vaSpread sprAepPorc 
            Height          =   2700
            Left            =   90
            OleObjectBlob   =   "frmCaj.frx":8119
            TabIndex        =   31
            Top             =   195
            Width           =   6615
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1500
      Left            =   7950
      TabIndex        =   3
      Top             =   -30
      Width           =   1005
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
         Left            =   210
         Picture         =   "frmCaj.frx":84EE
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1005
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
         Left            =   210
         Picture         =   "frmCaj.frx":8D10
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   585
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
         Index           =   0
         Left            =   210
         Picture         =   "frmCaj.frx":8E12
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   180
         Width           =   570
      End
   End
   Begin VB.Frame frdatos 
      Height          =   1530
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   7845
      Begin VB.Frame Frame1 
         Caption         =   "Fechas de Consultas "
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
         TabIndex        =   7
         Top             =   240
         Width           =   4920
         Begin VB.OptionButton optFechas 
            Caption         =   "Acumulado"
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
            Index           =   1
            Left            =   3330
            TabIndex        =   15
            Top             =   645
            Width           =   1305
         End
         Begin VB.OptionButton optFechas 
            Caption         =   "Un Día"
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
            Left            =   3330
            TabIndex        =   14
            Top             =   285
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.CommandButton botHelpFH 
            Height          =   345
            Left            =   2655
            Picture         =   "frmCaj.frx":8F14
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   615
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton botHelpFD 
            Height          =   345
            Left            =   2655
            Picture         =   "frmCaj.frx":9086
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   255
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskFDesde 
            Height          =   285
            Left            =   1380
            TabIndex        =   10
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
            Left            =   1395
            TabIndex        =   13
            Top             =   615
            Visible         =   0   'False
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
            Left            =   135
            TabIndex        =   9
            Top             =   615
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha "
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
            Left            =   135
            TabIndex        =   8
            Top             =   270
            Width           =   1185
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
         Height          =   1080
         Index           =   0
         Left            =   5445
         TabIndex        =   1
         Top             =   300
         Width           =   2175
         Begin VB.ComboBox cboGrupo 
            Height          =   315
            Left            =   165
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   405
            Width           =   1845
         End
      End
   End
End
Attribute VB_Name = "frmCaj"
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
Dim UsoPorInterior As Boolean

Dim UsoPorTotal As Boolean
Private Function L_Armarcondicion() As String
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant

If optFechas(0).Value Then
    Cond = " WHERE fch_ticket =" & func_ToDate(mskFDesde.FormattedText)
Else
    Cond = " WHERE fch_ticket between " & func_ToDate(mskFDesde.FormattedText) & " And " & func_ToDate(mskFHasta.FormattedText)
End If

If cboGrupo.Text <> "TODOS" Then
    Cond = Cond & " and grupo_venta = '" & cboGrupo.Text & "'"
End If

L_Armarcondicion = Cond

End Function

Private Sub L_DecoEspigon()
Dim fecha As String, desc As String, sql As String
Dim i As Integer, indDep As Integer

Do While Not RsData.EOF
        
        sql = "SELECT decode(apellido || ', ' || nombre,'INTERBAIRES S.A.,  ','NO HAY DATOS',apellido || ', ' || nombre) AS DESCRIP "
        sql = sql & "FROM personal.empleado Emp ,ventas.cajeros Caj "
        sql = sql & "WHERE Emp.legajo = Caj.legajo "
        sql = sql & " AND caj.cod_cajero = " & RsData!cod_cajero

        Call Func_ObtenerDesc(sql, desc)

        Select Case RsData!cod_depn
            Case DSLoc(1).Dep
                sprAEP.MaxRows = sprAEP.MaxRows + 1
                sprAEP.SetText 1, sprAEP.MaxRows, desc
                sprAEP.SetText 2, sprAEP.MaxRows, Trim(str(RsData!cod_cajero))
                sprAEP.SetText 3, sprAEP.MaxRows, str(RsData!imp)
                sprAEP.SetText 4, sprAEP.MaxRows, str(RsData!cant_t)
                sprAEP.SetText 5, sprAEP.MaxRows, str(RsData!cant_p)
                
            Case DSLoc(2).Dep
                Select Case RsData!cod_sdep
                    Case DSLoc(2).Sdep
                        sprEzeA.MaxRows = sprEzeA.MaxRows + 1
                        sprEzeA.SetText 1, sprEzeA.MaxRows, desc
                        sprEzeA.SetText 2, sprEzeA.MaxRows, Trim(str(RsData!cod_cajero))
                        sprEzeA.SetText 3, sprEzeA.MaxRows, str(RsData!imp)
                        sprEzeA.SetText 4, sprEzeA.MaxRows, str(RsData!cant_t)
                        sprEzeA.SetText 5, sprEzeA.MaxRows, str(RsData!cant_p)
                    Case DSLoc(3).Sdep
                        sprEzeB.MaxRows = sprEzeB.MaxRows + 1
                        sprEzeB.SetText 1, sprEzeB.MaxRows, desc
                        sprEzeB.SetText 2, sprEzeB.MaxRows, Trim(str(RsData!cod_cajero))
                        sprEzeB.SetText 3, sprEzeB.MaxRows, str(RsData!imp)
                        sprEzeB.SetText 4, sprEzeB.MaxRows, str(RsData!cant_t)
                        sprEzeB.SetText 5, sprEzeB.MaxRows, str(RsData!cant_p)

                End Select
        
            Case DSLoc(4).Dep
                        sprInt.MaxRows = sprInt.MaxRows + 1
                        sprInt.SetText 1, sprInt.MaxRows, desc
                        sprInt.SetText 2, sprInt.MaxRows, Trim(str(RsData!cod_cajero))
                        sprInt.SetText 3, sprInt.MaxRows, str(RsData!imp)
                        sprInt.SetText 4, sprInt.MaxRows, str(RsData!cant_t)
                        sprInt.SetText 5, sprInt.MaxRows, str(RsData!cant_p)
                    
        End Select
        RsData.MoveNext
        If RsData.EOF Then
            Exit Do
        End If

Loop

Spread_TotalesGrillas sprEzeA, sprEzeA.MaxCols - 3, 3
Spread_TotalesGrillas sprEzeB, sprEzeB.MaxCols - 3, 3
Spread_TotalesGrillas sprAEP, sprAEP.MaxCols - 3, 3
Spread_TotalesGrillas sprInt, sprInt.MaxCols - 3, 3

L_RefrescarFrames

End Sub
Private Sub L_LimpiarGrillas()

   Dim i
   sprEzeA.MaxRows = 0
   sprEzeB.MaxRows = 0
   sprAEP.MaxRows = 0
   sprInt.MaxRows = 0
   
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
            For j = 3 To sprDato.MaxCols
                sprPorc.SetText j, i, "100"
            Next
         Else
            sprDato.GetText 1, i, dato
            Result = 0
            sprPorc.MaxRows = i
        
            sprPorc.SetText 1, i, Trim(dato)
            
            sprDato.GetText 2, i, dato
            Result = 0
            sprPorc.MaxRows = i
        
            sprPorc.SetText 2, i, Trim(str(dato))
            
            For j = 3 To sprDato.MaxCols - 2
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
    frEzeAPorc0.Refresh
    frEzeBPorc0.Refresh
    frAepPorc0.Refresh
    FrIntPorc0.Refresh
    frTot0.Refresh
    Frame4.Refresh
    Frame5.Refresh
    Frame6.Refresh
    Frame7.Refresh
    
End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset

On Error GoTo ErrCaj:

frmCaj.caption = Aplicacion.SeteoProceso(frmCaj.caption)


sql = " SELECT "
sql = sql & " cod_cajero, "
'sql = sql & " grupo_venta,"
sql = sql & " cod_depn, "
sql = sql & " cod_sdep, "
sql = sql & " sum(cant_tick) cant_t, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cant_nro_pax) cant_p "
sql = sql & "FROM " & funcLocal_VistaTicket("view_ticket_h", Month(mskFDesde.FormattedText), Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion
sql = sql & "group by cod_cajero,cod_depn,cod_sdep"
sql = sql & " order by cod_cajero,cod_depn,cod_sdep"


If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        L_DecoEspigon
        Call optSort_Click(2)
        Call optSort_Click(5)
        Call optSort_Click(8)
    End If
    
    Aplicacion.CerrarDAO RsData
    
End If

ErrCaj:
    frmCaj.caption = Aplicacion.SeteoFin
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
        Case INTE
            UsoPorInterior = False
        Case total
            UsoPorTotal = False
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntB Or UsoPorAero Or UsoPorInterior Or UsoPorTotal) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub

Private Sub L_TotalesLinea(spr As control)

    'Select a block of cells
    spr.col = 1
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
L_TratarExcel sprAEP, "AEROPARQUE", "Informe por Cajeros (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", ""
End Sub
Private Sub L_TratarExcel(spr As control, Titulo As String, subTit As String, Info As String)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim fila As Integer, filaant As Integer
Dim i
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmCaj.caption = Aplicacion.SeteoProceso(frmCaj.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    AppExcel.Application.Visible = True
    
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
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, col, subTit
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    rango = Exl_rangos(fila, fila, 1, 1)
    
    Exl_PonerValor AppExcel, fila, col, "Grupo :" & cboGrupo.Text
    
    Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
    
    fila = fila + 2
    
    Exl_BajarGrillaExel spr, AppExcel, fila, col, titCol
    filaant = fila
    fila = fila + spr.MaxRows
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    rango = Exl_rangos(filaant + 1, fila, col + 4, spr.MaxCols)
    Exl_Format AppExcel, rango

    Exl_AnchoCol AppExcel, 1, 1, 25
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

    frmCaj.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub

Private Sub botExcelEzeA_Click()

L_TratarExcel sprEzeA, "ESPIGON INTERNACIONAL A", "Informe por Cajeros (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", ""

End Sub

Private Sub botExcelEzeb_Click()
L_TratarExcel sprEzeB, "ESPIGON INTERNACIONAL B", "Informe por Cajeros (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", ""
End Sub






Private Sub botExcelInt_Click()
L_TratarExcel sprInt, "INTERIOR", "Informe por Cajeros (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", ""
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
        botPorc(0).Enabled = False
        L_SeteoEjecutar False, EZEA
        frEzeA0.Refresh
    Case 2
        frEzeA0.Visible = False
        L_Porcentages sprEzeAPorc, sprEzeA, "COL"
        
        botPorc(0).Enabled = True
        frEzeAPorc0.Refresh
        L_SeteoEjecutar True, EZEA
        
End Select

End Sub





Private Sub botPorcInt_Click(Index As Integer)
Select Case Index
    Case 0
        FrInt0.Visible = True
        botPorcint(0).Enabled = False
        L_SeteoEjecutar False, INTE
        FrInt0.Refresh
    Case 2
        FrInt0.Visible = False
        L_Porcentages SprIntPorc, sprInt, "COL"
        
        botPorcint(0).Enabled = True
        FrIntPorc0.Refresh
        L_SeteoEjecutar True, INTE
        
End Select
End Sub




Private Sub botPorcB_Click(Index As Integer)

Select Case Index
    Case 0
        frEzeB0.Visible = True
        botPorcB(0).Enabled = False
        L_SeteoEjecutar False, EZEB
        frEzeB0.Refresh
    Case 2
        frEzeB0.Visible = False
        L_Porcentages sprEzeBPorc, sprEzeB, "COL"
        
        botPorcB(0).Enabled = True
        frEzeBPorc0.Refresh
        L_SeteoEjecutar True, EZEB
        
End Select

End Sub



Private Sub botPorcTot_Click(Index As Integer)

Select Case Index
    Case 0
        frTot0.Visible = True
        botPorcTot(0).Enabled = False
        L_SeteoEjecutar False, total
        frTot0.Refresh
    Case 1
        frTot0.Visible = False
        L_Porcentages sprTotPorc(0), SprTot(0), "FILAS"
        botPorcTot(0).Enabled = True
        frTotPorc0.Refresh
        L_SeteoEjecutar True, total
    Case 2
        frTot0.Visible = False
        L_Porcentages sprTotPorc(0), SprTot(0), "COL"
        
        botPorcTot(0).Enabled = True
        frTotPorc0.Refresh
        L_SeteoEjecutar True, total
        
End Select


End Sub









Private Sub Form_Activate()
FuncLocal_SeteoTABS tabEspigon
tabEspigon.TabVisible(0) = False

optSort(2).Value = True
optSort(6).Value = True
optSort(10).Value = True
optSort(14).Value = True

L_RefrescarFrames

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

mskFDesde.Text = func_Dia1SegunMes_Anio(Month(Date), Year(Date))
mskFHasta.Text = Format$(Date - 1, FTOFECHA)

'FuncLocal_SeteoTABS tabEspigon

UsoPorIntA = False
UsoPorIntB = False
UsoPorTotal = False
UsoPorAero = False

L_RefrescarFrames

L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmCaj"

End Sub


Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmCaj"
End Sub


Private Sub mskFDesde_LostFocus()

    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    Else
    'If (Month(mskFDesde.FormattedText) < Month(Date)) Then
    '    mskFDesde.Text = Format$("01-" & Month(Date) & "-" & Year(Date), FTOFECHA)
    End If
    
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)

    If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
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


Private Sub optFechas_Click(Index As Integer)
If Index = 0 Then
    Label1(0).caption = "Fecha "
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
    Label1(1).Visible = False
    mskFHasta.Visible = False
    botHelpFH.Visible = False
Else
    Label1(0).caption = "Fecha Desde"
    Label1(1).Visible = True
    mskFHasta.Visible = True
    botHelpFH.Visible = True

End If
End Sub


Private Sub optSort_Click(Index As Integer)


Select Case Index
    Case 0
        sprEzeA.Row = 1
        sprEzeA.col = 1
        sprEzeA.Row2 = sprEzeA.MaxRows - 1
        sprEzeA.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeA.SortBy = SS_SORT_BY_ROW

        sprEzeA.SortKey(1) = 2
        sprEzeA.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprEzeA.Action = SS_ACTION_SORT
    Case 1
        sprEzeA.Row = 1
        sprEzeA.col = 1
        sprEzeA.Row2 = sprEzeA.MaxRows - 1
        sprEzeA.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeA.SortBy = SS_SORT_BY_ROW
        
        sprEzeA.SortKey(1) = 3
        sprEzeA.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprEzeA.Action = SS_ACTION_SORT
    Case 2
        sprEzeA.Row = 1
        sprEzeA.col = 1
        sprEzeA.Row2 = sprEzeA.MaxRows - 1
        sprEzeA.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeA.SortBy = SS_SORT_BY_ROW
        
        sprEzeA.SortKey(1) = 3
        sprEzeA.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprEzeA.Action = SS_ACTION_SORT
    Case 3
        sprEzeA.Row = 1
        sprEzeA.col = 1
        sprEzeA.Row2 = sprEzeA.MaxRows - 1
        sprEzeA.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeA.SortBy = SS_SORT_BY_ROW
        
        sprEzeA.SortKey(1) = 6
        sprEzeA.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprEzeA.Action = SS_ACTION_SORT
            
    Case 4
        sprEzeB.Row = 1
        sprEzeB.col = 1
        sprEzeB.Row2 = sprEzeB.MaxRows - 1
        sprEzeB.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeB.SortBy = SS_SORT_BY_ROW
        
        sprEzeB.SortKey(1) = 2
        sprEzeB.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprEzeB.Action = SS_ACTION_SORT
    Case 5
        sprEzeB.Row = 1
        sprEzeB.col = 1
        sprEzeB.Row2 = sprEzeB.MaxRows - 1
        sprEzeB.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeB.SortBy = SS_SORT_BY_ROW
        sprEzeB.SortKey(1) = 3
        sprEzeB.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprEzeB.Action = SS_ACTION_SORT
    Case 6
        sprEzeB.Row = 1
        sprEzeB.col = 1
        sprEzeB.Row2 = sprEzeB.MaxRows - 1
        sprEzeB.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeB.SortBy = SS_SORT_BY_ROW
        sprEzeB.SortKey(1) = 6
        sprEzeB.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprEzeB.Action = SS_ACTION_SORT
    Case 8
        sprAEP.Row = 1
        sprAEP.col = 1
        sprAEP.Row2 = sprAEP.MaxRows - 1
        sprAEP.Col2 = 5
        
        ' Set sort definition for key 1
        sprAEP.SortBy = SS_SORT_BY_ROW
        sprAEP.SortKey(1) = 2
        sprAEP.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprAEP.Action = SS_ACTION_SORT
    Case 9
        sprAEP.Row = 1
        sprAEP.col = 1
        sprAEP.Row2 = sprAEP.MaxRows - 1
        sprAEP.Col2 = 5
        
        ' Set sort definition for key 1
        sprAEP.SortBy = SS_SORT_BY_ROW
        sprAEP.SortKey(1) = 3
        sprAEP.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprAEP.Action = SS_ACTION_SORT
    Case 10
        sprAEP.Row = 1
        sprAEP.col = 1
        sprAEP.Row2 = sprAEP.MaxRows - 1
        sprAEP.Col2 = 5
        
        ' Set sort definition for key 1
        sprAEP.SortBy = SS_SORT_BY_ROW
        sprAEP.SortKey(1) = 3
        sprAEP.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprAEP.Action = SS_ACTION_SORT
    Case 11
        sprAEP.Row = 1
        sprAEP.col = 1
        sprAEP.Row2 = sprAEP.MaxRows - 1
        sprAEP.Col2 = 5
        
        ' Set sort definition for key 1
        sprAEP.SortBy = SS_SORT_BY_ROW
        sprAEP.SortKey(1) = 6
        sprAEP.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprAEP.Action = SS_ACTION_SORT
    Case 12
    
        sprInt.Row = 1
        sprInt.col = 1
        sprInt.Row2 = sprInt.MaxRows - 1
        sprInt.Col2 = 5
        
        ' Set sort definition for key 1
        sprInt.SortBy = SS_SORT_BY_ROW

        sprInt.SortKey(1) = 2
        sprInt.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprInt.Action = SS_ACTION_SORT
    Case 13
        sprInt.Row = 1
        sprInt.col = 1
        sprInt.Row2 = sprInt.MaxRows - 1
        sprInt.Col2 = 5
        
        ' Set sort definition for key 1
        sprInt.SortBy = SS_SORT_BY_ROW
        
        sprInt.SortKey(1) = 3
        sprInt.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprInt.Action = SS_ACTION_SORT
    Case 14
        sprInt.Row = 1
        sprInt.col = 1
        sprInt.Row2 = sprInt.MaxRows - 1
        sprInt.Col2 = 5
        
        ' Set sort definition for key 1
        sprInt.SortBy = SS_SORT_BY_ROW
        
        sprInt.SortKey(1) = 3
        sprInt.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprInt.Action = SS_ACTION_SORT
    Case 15
        sprInt.Row = 1
        sprInt.col = 1
        sprInt.Row2 = sprInt.MaxRows - 1
        sprInt.Col2 = 5
        
        ' Set sort definition for key 1
        sprInt.SortBy = SS_SORT_BY_ROW
        
        sprInt.SortKey(1) = 6
        sprInt.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprInt.Action = SS_ACTION_SORT


End Select




End Sub


Private Sub sprAep_DblClick(ByVal col As Long, ByVal Row As Long)
Dim valor As Variant

If col = 2 Then
sprAEP.GetText col, Row, valor

If valor <> "" Then
frmDatosCajero.DatosEmpleado valor
End If
End If

End Sub


Private Sub sprEzeA_DblClick(ByVal col As Long, ByVal Row As Long)
Dim valor As Variant

If col = 2 Then
sprEzeA.GetText col, Row, valor

If valor <> "" Then
frmDatosCajero.DatosEmpleado valor
End If
End If
End Sub


Private Sub sprEzeB_DblClick(ByVal col As Long, ByVal Row As Long)
Dim valor As Variant

If col = 2 Then
sprEzeB.GetText col, Row, valor

If valor <> "" Then
frmDatosCajero.DatosEmpleado valor
End If
End If
End Sub


