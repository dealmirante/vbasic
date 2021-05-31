VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLocNac 
   Caption         =   "Ventas por Nacionalidad"
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
      Picture         =   "frmLocNac.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
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
      Picture         =   "frmLocNac.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   31
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
      Picture         =   "frmLocNac.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   30
      Width           =   570
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   4185
      Left            =   15
      TabIndex        =   1
      Top             =   1365
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   7382
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   459
      Enabled         =   0   'False
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "RESUMEN"
      TabPicture(0)   =   "frmLocNac.frx":0A26
      Tab(0).ControlCount=   6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frPorc(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frNum(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "botPorcResumen(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "botPorcResumen(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "BotExcelResumen"
      Tab(0).Control(5).Enabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmLocNac.frx":0A42
      Tab(1).ControlCount=   16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkEzeA(3)"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "chkEzeA(5)"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "chkEzeA(7)"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "chkEzeA(6)"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "chkEzeA(4)"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "chkEzeA(2)"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "chkEzeA(9)"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "chkEzeA(8)"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "chkEzeA(1)"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "Frame4(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "botPorc(2)"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "botPorc(0)"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "botExcelEzeA"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "chkEzeA(0)"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "frNum(0)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "frPorc(0)"
      Tab(1).Control(15).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmLocNac.frx":0A5E
      Tab(2).ControlCount=   11
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frPorc(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frNum(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "chkEzeB(0)"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "chkEzeB(1)"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "chkEzeB(3)"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "chkEzeB(4)"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "botExcelEzeB"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "botPorcB(0)"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "botPorcB(2)"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "Frame4(1)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "chkEzeB(2)"
      Tab(2).Control(10).Enabled=   -1  'True
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmLocNac.frx":0A7A
      Tab(3).ControlCount=   9
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "chkAep(2)"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "botPorcAep(2)"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "botPorcAep(0)"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "botExcelAep"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "chkAep(1)"
      Tab(3).Control(5).Enabled=   -1  'True
      Tab(3).Control(6)=   "chkAep(0)"
      Tab(3).Control(6).Enabled=   -1  'True
      Tab(3).Control(7)=   "frNum(2)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "frPorc(2)"
      Tab(3).Control(8).Enabled=   0   'False
      Begin VB.CheckBox chkEzeB 
         Caption         =   "L 01 (2)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   2
         Left            =   -72825
         TabIndex        =   78
         Top             =   465
         Value           =   1  'Checked
         Width           =   960
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 22"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   3
         Left            =   -69465
         TabIndex        =   4
         Top             =   405
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 08"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   5
         Left            =   -69765
         TabIndex        =   6
         Top             =   405
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 06"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   7
         Left            =   -72285
         TabIndex        =   8
         Top             =   435
         Value           =   1  'Checked
         Width           =   900
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 08"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   6
         Left            =   -73170
         TabIndex        =   7
         Top             =   420
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 07 (1)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   4
         Left            =   -70410
         TabIndex        =   5
         Top             =   405
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 06"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   2
         Left            =   -70605
         TabIndex        =   3
         Top             =   405
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 22"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   9
         Left            =   -74010
         TabIndex        =   77
         Top             =   420
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 21"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   8
         Left            =   -74730
         TabIndex        =   76
         Top             =   420
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 05(1)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   1
         Left            =   -70830
         TabIndex        =   75
         Top             =   405
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton BotExcelResumen 
         Caption         =   "Excel"
         Height          =   510
         Left            =   7515
         Picture         =   "frmLocNac.frx":0A96
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1515
         Width           =   765
      End
      Begin VB.CommandButton botPorcResumen 
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
         Left            =   7530
         Picture         =   "frmLocNac.frx":1028
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   480
         Width           =   765
      End
      Begin VB.CommandButton botPorcResumen 
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
         Left            =   7530
         Picture         =   "frmLocNac.frx":199A
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1005
         Width           =   765
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
         Height          =   1725
         Index           =   3
         Left            =   7425
         TabIndex        =   67
         Top             =   2325
         Width           =   1590
         Begin VB.OptionButton optSort 
            Caption         =   "Por Pr. (Desc)"
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
            Left            =   30
            TabIndex        =   71
            Top             =   1230
            Width           =   1530
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
            Index           =   14
            Left            =   15
            TabIndex        =   70
            Top             =   900
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
            Index           =   13
            Left            =   30
            TabIndex        =   69
            Top             =   255
            Value           =   -1  'True
            Width           =   1290
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
            Index           =   12
            Left            =   30
            TabIndex        =   68
            Top             =   570
            Width           =   1470
         End
      End
      Begin VB.Frame frNum 
         Height          =   3390
         Index           =   3
         Left            =   120
         TabIndex        =   63
         Top             =   650
         Width           =   7300
         Begin FPSpread.vaSpread SprResumen 
            Height          =   2955
            Left            =   75
            OleObjectBlob   =   "frmLocNac.frx":2060
            TabIndex        =   64
            Top             =   285
            Width           =   7155
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
         Height          =   1560
         Index           =   2
         Left            =   -67530
         TabIndex        =   58
         Top             =   2250
         Width           =   1545
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
            Index           =   11
            Left            =   30
            TabIndex        =   62
            Top             =   600
            Width           =   1470
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
            Index           =   10
            Left            =   30
            TabIndex        =   61
            Top             =   255
            Value           =   -1  'True
            Width           =   1290
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
            Index           =   9
            Left            =   30
            TabIndex        =   60
            Top             =   930
            Width           =   1530
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Pr. (Desc)"
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
            Left            =   30
            TabIndex        =   59
            Top             =   1230
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
         Height          =   1560
         Index           =   1
         Left            =   -67575
         TabIndex        =   53
         Top             =   2355
         Width           =   1575
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
            Index           =   7
            Left            =   30
            TabIndex        =   57
            Top             =   600
            Width           =   1470
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
            Index           =   6
            Left            =   30
            TabIndex        =   56
            Top             =   255
            Value           =   -1  'True
            Width           =   1290
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
            Index           =   5
            Left            =   30
            TabIndex        =   55
            Top             =   930
            Width           =   1530
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Pr. (Desc)"
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
            TabIndex        =   54
            Top             =   1230
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
         Height          =   1725
         Index           =   0
         Left            =   -67575
         TabIndex        =   48
         Top             =   2175
         Width           =   1590
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
            TabIndex        =   52
            Top             =   600
            Width           =   1470
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
            Left            =   30
            TabIndex        =   51
            Top             =   255
            Value           =   -1  'True
            Width           =   1290
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
            TabIndex        =   50
            Top             =   930
            Width           =   1530
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Por Pr. (Desc)"
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
            TabIndex        =   49
            Top             =   1230
            Width           =   1530
         End
      End
      Begin VB.CheckBox chkAep 
         Caption         =   "L 14 (1)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   2
         Left            =   -72720
         TabIndex        =   33
         Top             =   450
         Value           =   1  'Checked
         Width           =   1035
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
         Left            =   -67155
         Picture         =   "frmLocNac.frx":2579
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   975
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
         Left            =   -67155
         Picture         =   "frmLocNac.frx":2C3F
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Left            =   -67230
         Picture         =   "frmLocNac.frx":35B1
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   960
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
         Left            =   -67245
         Picture         =   "frmLocNac.frx":3C77
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   -66855
         Picture         =   "frmLocNac.frx":45E9
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   -66855
         Picture         =   "frmLocNac.frx":4CAF
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   450
         Width           =   765
      End
      Begin VB.CommandButton botExcelAep 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67155
         Picture         =   "frmLocNac.frx":5621
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1500
         Width           =   765
      End
      Begin VB.CommandButton botExcelEzeB 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67230
         Picture         =   "frmLocNac.frx":5BB3
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1470
         Width           =   750
      End
      Begin VB.CommandButton botExcelEzeA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -66855
         Picture         =   "frmLocNac.frx":6145
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1500
         Width           =   765
      End
      Begin VB.CheckBox chkAep 
         Caption         =   "L 14"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   1
         Left            =   -73680
         TabIndex        =   14
         Top             =   450
         Value           =   1  'Checked
         Width           =   720
      End
      Begin VB.CheckBox chkAep 
         Caption         =   "L 09"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   -74820
         TabIndex        =   13
         Top             =   450
         Value           =   1  'Checked
         Width           =   720
      End
      Begin VB.CheckBox chkEzeB 
         Caption         =   "L 03"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   4
         Left            =   -70710
         TabIndex        =   12
         Top             =   465
         Value           =   1  'Checked
         Width           =   720
      End
      Begin VB.CheckBox chkEzeB 
         Caption         =   "L 02"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   3
         Left            =   -71685
         TabIndex        =   11
         Top             =   465
         Value           =   1  'Checked
         Width           =   720
      End
      Begin VB.CheckBox chkEzeB 
         Caption         =   "L 01 (1)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   1
         Left            =   -73920
         TabIndex        =   10
         Top             =   450
         Value           =   1  'Checked
         Width           =   960
      End
      Begin VB.CheckBox chkEzeB 
         Caption         =   "L 01"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   -74895
         TabIndex        =   9
         Top             =   450
         Value           =   1  'Checked
         Width           =   720
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 05"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   -71325
         TabIndex        =   2
         Top             =   405
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Frame frNum 
         Height          =   3390
         Index           =   1
         Left            =   -74900
         TabIndex        =   38
         Top             =   650
         Width           =   7300
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2955
            Left            =   90
            OleObjectBlob   =   "frmLocNac.frx":66D7
            TabIndex        =   42
            Top             =   225
            Width           =   7155
         End
      End
      Begin VB.Frame frNum 
         Height          =   3390
         Index           =   2
         Left            =   -74900
         TabIndex        =   39
         Top             =   650
         Width           =   7300
         Begin FPSpread.vaSpread sprAep 
            Height          =   2955
            Left            =   75
            OleObjectBlob   =   "frmLocNac.frx":6BF0
            TabIndex        =   40
            Top             =   285
            Width           =   7155
         End
      End
      Begin VB.Frame frPorc 
         Height          =   3390
         Index           =   2
         Left            =   -74865
         TabIndex        =   44
         Top             =   660
         Width           =   7305
         Begin FPSpread.vaSpread sprPorcAep 
            Height          =   2895
            Left            =   345
            OleObjectBlob   =   "frmLocNac.frx":7109
            TabIndex        =   47
            Top             =   300
            Width           =   6600
         End
      End
      Begin VB.Frame frPorc 
         Height          =   3390
         Index           =   1
         Left            =   -74895
         TabIndex        =   43
         Top             =   645
         Width           =   7305
         Begin FPSpread.vaSpread sprPorcB 
            Height          =   2895
            Left            =   300
            OleObjectBlob   =   "frmLocNac.frx":7482
            TabIndex        =   46
            Top             =   300
            Width           =   6600
         End
      End
      Begin VB.Frame frNum 
         Height          =   3390
         Index           =   0
         Left            =   -74880
         TabIndex        =   36
         Top             =   650
         Width           =   7300
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2955
            Left            =   75
            OleObjectBlob   =   "frmLocNac.frx":77FB
            TabIndex        =   37
            Top             =   285
            Width           =   7155
         End
      End
      Begin VB.Frame frPorc 
         Height          =   3390
         Index           =   3
         Left            =   120
         TabIndex        =   65
         Top             =   645
         Width           =   7305
         Begin FPSpread.vaSpread SprPorcResumen 
            Height          =   2895
            Left            =   375
            OleObjectBlob   =   "frmLocNac.frx":7D14
            TabIndex        =   66
            Top             =   315
            Width           =   6600
         End
      End
      Begin VB.Frame frPorc 
         Height          =   3390
         Index           =   0
         Left            =   -74895
         TabIndex        =   41
         Top             =   660
         Width           =   7305
         Begin FPSpread.vaSpread sprPorcA 
            Height          =   2895
            Left            =   375
            OleObjectBlob   =   "frmLocNac.frx":808D
            TabIndex        =   45
            Top             =   300
            Width           =   6600
         End
      End
   End
   Begin VB.Frame frdatos 
      Height          =   1380
      Left            =   15
      TabIndex        =   0
      Top             =   -45
      Width           =   7560
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
         Height          =   675
         Index           =   0
         Left            =   3480
         TabIndex        =   34
         Top             =   360
         Width           =   2940
         Begin VB.ComboBox cboGrupo 
            Height          =   315
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   225
            Width           =   1845
         End
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   2490
         Picture         =   "frmLocNac.frx":8406
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   690
         Width           =   375
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2505
         Picture         =   "frmLocNac.frx":8578
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   210
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1335
         TabIndex        =   20
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
         Left            =   1335
         TabIndex        =   21
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
         Left            =   135
         TabIndex        =   23
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
         Left            =   135
         TabIndex        =   22
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmLocNac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsDataEzeA  As Recordset
Dim RsDataEzeB  As Recordset
Dim RsDataAep  As Recordset

Dim RsDataEzeAPax  As Recordset
Dim RsDataEzeBPax  As Recordset
Dim RsDataAepPax  As Recordset

Dim RsDataResumen As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas
                       'con porcentajes
Dim UsoPorIntB As Boolean
Dim UsoPorAero As Boolean
Dim UsoPorTotal As Boolean



Private Function L_Locales(Esp As String) As String
Dim i
Dim LC As String

LC = ""
Select Case Esp
    Case AERO
        For i = 0 To 2
            If chkAep(i).Value = 1 Then
                LC = LC & chkAep(i).caption & ", "
            End If
        Next
    Case EZEA
        For i = 0 To 7
            If chkEzeA(i).Value = 1 Then
                LC = LC & chkEzeA(i).caption & ", "
            End If
        Next
    
    Case EZEB
        For i = 0 To 3
            If chkEzeB(i).Value = 1 Then
                LC = LC & chkEzeB(i).caption & ", "
            End If
        Next
    
End Select

L_Locales = Left(LC, Len(LC) - 2)

End Function

Private Function L_Armarcondicion(Esp As String) As String
Dim Cond As String, condLoc As String
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant, i

If cboGrupo.Text <> "" Then
    Cond = Cond & " and  grupo_venta = '" & cboGrupo.Text & "' "
End If

Select Case Esp
    Case EZEA
        
        condLoc = ""
        For i = 6 To 9 'NroLocINTA
           If i = 6 Then
              If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L08' ) "
              End If
           ElseIf i = 7 Then
              If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L06' ) "
              End If
           ElseIf i = 8 Then
              If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L21' ) "
              End If
           ElseIf i = 9 Then
              If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L22' ) "
              End If
           End If
        Next
        If condLoc <> "" Then
            Cond = Cond & " and (Cod_depn = 'EZE' and cod_sdep = 'INTA' and " & Right(condLoc, Len(condLoc) - 4) & ") "
        Else
            Cond = Cond & " and  cod_local = '00'"
        End If
    Case EZEB
        condLoc = ""
        For i = 0 To 4
            If chkEzeB(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = '" & Left(DSLoc(3).locales(i + 1), 3) & "' and cod_sloc = " & Right(DSLoc(3).locales(i + 1), 1) & ") "
            End If
        Next
        If condLoc <> "" Then
            Cond = Cond & " and (Cod_depn = 'EZE' and cod_sdep = 'INTB' and " & Right(condLoc, Len(condLoc) - 4) & ") "
        Else
            Cond = Cond & " and  cod_local = '00'"
        End If
    Case AERO
        condLoc = ""
        For i = 0 To NroLocAERO - 1
            If chkAep(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = '" & Left(DSLoc(1).locales(i + 1), 3) & "' and cod_sloc = " & Right(DSLoc(1).locales(i + 1), 1) & ") "
            End If
        Next
        If condLoc <> "" Then
            Cond = Cond & " and (Cod_depn = 'AEP' and cod_sdep = 'AEP' and " & Right(condLoc, Len(condLoc) - 4) & ") "
        Else
            Cond = Cond & " and  cod_local = '00'"
        End If
End Select

L_Armarcondicion = Cond

End Function


Private Function L_Armarcondicion_Resumen() As String

Dim Cond As String, condLoc As String
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant, i

If cboGrupo.Text <> "" Then
    Cond = Cond & " and  grupo_venta = '" & cboGrupo.Text & "' "    ''''& " and ("
  Else
    '''Cond = " and ("
End If
    

        For i = 7 To 9
           If i = 7 Then
              If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L06' ) "
              End If
           ElseIf i = 8 Then
              If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L21' ) "
              End If
           ElseIf i = 9 Then
              If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L22' ) "
              End If
           End If
        Next

        For i = 0 To 4
            If chkEzeB(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = '" & Left(DSLoc(3).locales(i + 1), 3) & "' and cod_sloc = " & Right(DSLoc(3).locales(i + 1), 1) & ") "
            End If
        Next

        For i = 0 To NroLocAERO - 1
            If chkAep(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = '" & Left(DSLoc(1).locales(i + 1), 3) & "' and cod_sloc = " & Right(DSLoc(1).locales(i + 1), 1) & ") "
            End If
        Next

'Cond = Cond & condLoc & ")"

        If condLoc <> "" Then
            Cond = Cond & "  and (" & Right(condLoc, Len(condLoc) - 4) & ") "
        Else
            Cond = Cond & " and  cod_local = '00'"
        End If

L_Armarcondicion_Resumen = Cond

End Function

Private Sub L_LimpiarGrillas()
Dim i

sprEzeA.MaxRows = 0
sprEzeB.MaxRows = 0
sprAep.MaxRows = 0
SprResumen.MaxRows = 0

End Sub

Private Function L_Locales_Resumen() As String
Dim i
Dim LC As String

LC = ""

For i = 0 To 2
   If chkAep(i).Value = 1 Then
      LC = LC & chkAep(i).caption & ", "
   End If
Next

For i = 0 To 7
   If chkEzeA(i).Value = 1 Then
      LC = LC & chkEzeA(i).caption & ", "
   End If
Next
    
For i = 0 To 3
   If chkEzeB(i).Value = 1 Then
      LC = LC & chkEzeB(i).caption & ", "
   End If
Next
    
L_Locales_Resumen = Left(LC, Len(LC) - 2)

End Function

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
            sprDato.GetText 1, i, fecha
            Result = 0
            sprPorc.MaxRows = i
        
            sprPorc.SetText 1, i, Format$(fecha, "dd-mm-yy")
            
            For j = 2 To sprDato.MaxCols - 2
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
Dim i

For i = 0 To 3
    frNum(i).Refresh
    Frame4(i).Refresh
Next

End Sub

Private Sub L_RefrescarEzeA()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrEA:

frmLocNac.caption = Aplicacion.SeteoProceso(frmLocNac.caption)

sql = " SELECT "
sql = sql & " P.descrip, "
sql = sql & " sum(cant_tickets) cant_t, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cant_pax) cant_p "
sql = sql & "FROM " & funcLocal_Vista("pax_grupo", Year(mskFHasta.FormattedText))
sql = sql & " V, ventas.paises P "
sql = sql & " WHERE nacionalidad=cod_pais "
sql = sql & L_Armarcondicion(EZEA)
sql = sql & " And fch_vta BETWEEN  " & func_ToDate(mskFDesde.FormattedText) & " And " & func_ToDate(mskFHasta.FormattedText)
sql = sql & " And cod_sdep  = 'INTA' "
sql = sql & "group by p.descrip "
sql = sql & " order by p.descrip "


If Aplicacion.ObtenerRsDAO(sql, RsDataEzeA) Then
    
    If Aplicacion.CantReg(RsDataEzeA) > 0 Then
        L_LlenarGrillaINTA
        L_Porcentages sprPorcA, sprEzeA, "COL"
    End If

    Aplicacion.CerrarDAO RsDataEzeA
    
End If

ErrEA:
    frmLocNac.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub

Private Sub L_RefrescarEzeB()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrEB:

frmLocNac.caption = Aplicacion.SeteoProceso(frmLocNac.caption)

sql = " SELECT "
sql = sql & " P.descrip, "
sql = sql & " sum(cant_tickets) cant_t, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cant_pax) cant_p "
sql = sql & "FROM " & funcLocal_Vista("pax_grupo", Year(mskFHasta.FormattedText))
sql = sql & " V, ventas.paises P "
sql = sql & " WHERE nacionalidad=cod_pais "
sql = sql & L_Armarcondicion(EZEB)
sql = sql & " And fch_vta BETWEEN  " & func_ToDate(mskFDesde.FormattedText) & " And " & func_ToDate(mskFHasta.FormattedText)
sql = sql & "group by p.descrip "
sql = sql & " order by p.descrip "


If Aplicacion.ObtenerRsDAO(sql, RsDataEzeB) Then
    
    If Aplicacion.CantReg(RsDataEzeB) > 0 Then
        L_LlenarGrillaINTB
        L_Porcentages sprPorcB, sprEzeB, "COL"
    End If

    Aplicacion.CerrarDAO RsDataEzeB
    
End If

ErrEB:
    frmLocNac.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub

Private Sub L_RefrescarResumen()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrEB:

frmLocNac.caption = Aplicacion.SeteoProceso(frmLocNac.caption)

sql = " SELECT "
sql = sql & " P.descrip, "
sql = sql & " sum(cant_tickets) cant_t, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cant_pax) cant_p "
sql = sql & "FROM " & funcLocal_Vista("pax_grupo", Year(mskFHasta.FormattedText))
sql = sql & " V, ventas.paises P "
sql = sql & " WHERE nacionalidad=cod_pais "
sql = sql & L_Armarcondicion_Resumen
sql = sql & " And fch_vta BETWEEN  " & func_ToDate(mskFDesde.FormattedText) & " And " & func_ToDate(mskFHasta.FormattedText)
sql = sql & "group by p.descrip "
sql = sql & " order by p.descrip "


If Aplicacion.ObtenerRsDAO(sql, RsDataResumen) Then
    
    If Aplicacion.CantReg(RsDataResumen) > 0 Then
        L_LlenarGrillaResumen
        L_Porcentages SprPorcResumen, SprResumen, "COL"
    End If

    Aplicacion.CerrarDAO RsDataResumen
    
End If

ErrEB:
    frmLocNac.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub

Private Sub L_RefrescarAep()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrEB:

frmLocNac.caption = Aplicacion.SeteoProceso(frmLocNac.caption)

sql = " SELECT "
sql = sql & " P.descrip, "
sql = sql & " sum(cant_tickets) cant_t, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cant_pax) cant_p "
sql = sql & "FROM " & funcLocal_Vista("pax_grupo", Year(mskFHasta.FormattedText))
sql = sql & " V, ventas.paises P "
sql = sql & " WHERE nacionalidad=cod_pais "
sql = sql & L_Armarcondicion(AERO)
sql = sql & " And fch_vta BETWEEN  " & func_ToDate(mskFDesde.FormattedText) & " And " & func_ToDate(mskFHasta.FormattedText)
sql = sql & "group by p.descrip "
sql = sql & " order by p.descrip "


If Aplicacion.ObtenerRsDAO(sql, RsDataAep) Then
    
    If Aplicacion.CantReg(RsDataAep) > 0 Then
        L_LlenarGrillaAEP
        L_Porcentages sprPorcAep, sprAep, "COL"
    End If

    Aplicacion.CerrarDAO RsDataAep
    
End If

ErrEB:
    frmLocNac.caption = Aplicacion.SeteoFin
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
        Case Total
            UsoPorTotal = True
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorAero Or UsoPorTotal) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub

Private Sub L_SeteosGral()
Dim i

    frdatos.Enabled = False
    botEjecutar(0).Enabled = False
    tabEspigon.Enabled = True

End Sub

Private Sub botEjecutar_Click(Index As Integer)

Select Case Index
    
    Case 0
        L_RefrescarEzeA
        L_RefrescarEzeB
        L_RefrescarAep
        L_RefrescarResumen
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
L_TratarExcel sprAep, sprPorcAep, "AEROPARQUE ", "Informe por Nacionalidad (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", AERO
End Sub

Private Sub botExcelEzeA_Click()
    L_TratarExcel sprEzeA, sprPorcA, "ESPIGON INTERNACIONAL A", "Informe por Nacionalidad (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", EZEA
End Sub

Private Sub botExcelEzeb_Click()
    L_TratarExcel sprEzeB, sprPorcB, "ESPIGON INTERNACIONAL B", "Informe por Nacionalidad (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", EZEB
End Sub

Private Sub L_TratarExcel(spr As control, spr2 As control, titulo As String, subTit As String, Info As String)
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

frmLocNac.caption = Aplicacion.SeteoProceso(frmLocNac.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.application.Visible = True
    
    ReDim titCol(spr.MaxCols)
    Col = 1
    fila = 4
    
    For i = 1 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 2, 1, titulo
    rango = Exl_rangos(2, 2, 1, spr.MaxCols)
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
    Exl_PonerValor AppExcel, fila, Col, "Locales : " & L_Locales_Resumen
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, "Grupo :" & cboGrupo.Text
        
    fila = fila + 2
    Exl_BajarGrillaExel spr, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + spr.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 2
    Exl_BajarGrillaExel spr2, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + spr.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    'AppExcel.application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If

    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmLocNac.caption = Aplicacion.SeteoFin
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

Private Sub BotExcelResumen_Click()
    L_TratarExcel SprResumen, SprPorcResumen, "ESPIGON INTERNACIONAL A,B y AEP", "Informe por Nacionalidad (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", Total
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
        frNum(0).Visible = True
        botPorc(0).Enabled = False
        L_SeteoEjecutar False, EZEA
        frNum(0).Refresh
    Case 2
        frNum(0).Visible = False
        L_Porcentages sprPorcA, sprEzeA, "COL"
        botPorc(0).Enabled = True
        frPorc(0).Refresh
        L_SeteoEjecutar True, EZEA
        
End Select

End Sub

Private Sub botPorcaep_Click(Index As Integer)

Select Case Index
    Case 0
        frNum(2).Visible = True
        botPorcAep(0).Enabled = False
        L_SeteoEjecutar False, AERO
        frNum(2).Refresh
    Case 2
        frNum(2).Visible = False
        L_Porcentages sprPorcAep, sprAep, "COL"

        botPorcAep(0).Enabled = True
        frPorc(2).Refresh
        L_SeteoEjecutar True, AERO
        
End Select


End Sub

Private Sub botPorcB_Click(Index As Integer)

Select Case Index
    Case 0
        frNum(1).Visible = True
        botPorcB(0).Enabled = False
        L_SeteoEjecutar False, EZEB
        frNum(1).Refresh
    Case 2
        frNum(1).Visible = False
        L_Porcentages sprPorcB, sprEzeB, "COL"
        botPorcB(0).Enabled = True
        frPorc(1).Refresh
        L_SeteoEjecutar True, EZEB
        
End Select

End Sub



Private Sub botPorcResumen_Click(Index As Integer)
Dim i


Select Case Index
    Case 0
        frNum(3).Visible = True
        botPorcResumen(0).Enabled = False
        L_SeteoEjecutar False, Total
        frNum(3).Refresh
    Case 2
        frNum(3).Visible = False
        L_Porcentages SprPorcResumen, SprResumen, "COL"
        botPorcResumen(0).Enabled = True
        frPorc(3).Refresh
        L_SeteoEjecutar True, Total
        
End Select
End Sub

Private Sub cboGrupo_KeyPress(KeyAscii As Integer)

If KeyAscii = 32 Then
    cboGrupo.ListIndex = -1
End If

End Sub

Private Sub chkAep_Click(Index As Integer)
sprAep.MaxRows = 0
L_RefrescarAep
L_RefrescarResumen
End Sub

Private Sub chkEzeA_Click(Index As Integer)
sprEzeA.MaxRows = 0
L_RefrescarEzeA
L_RefrescarResumen
End Sub

Private Sub chkEzeB_Click(Index As Integer)
sprEzeB.MaxRows = 0
L_RefrescarEzeB
L_RefrescarResumen
End Sub


Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabEspigon

L_RefrescarFrames

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

cboGrupo.AddItem "A"
cboGrupo.AddItem "B"
cboGrupo.AddItem "C"

UsoPorIntA = False
UsoPorIntB = False
UsoPorTotal = False
UsoPorAero = False


L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmLocNac"

End Sub




Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmGLC"
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
Dim fecha As String
Dim i As Integer
Dim Total As Variant
Dim TOTImp As Long
Dim totCant As Long


sprEzeA.MaxRows = 0

Do While Not RsDataEzeA.EOF
    sprEzeA.MaxRows = sprEzeA.MaxRows + 1
    
    sprEzeA.SetText 1, sprEzeA.MaxRows, Trim(RsDataEzeA!Descrip)
    sprEzeA.SetText 2, sprEzeA.MaxRows, str(RsDataEzeA!imp)
    sprEzeA.SetText 3, sprEzeA.MaxRows, str(RsDataEzeA!cant_t)
    sprEzeA.SetText 4, sprEzeA.MaxRows, str(RsDataEzeA!cant_p)
    
    RsDataEzeA.MoveNext
Loop

Spread_TotalesGrillas sprEzeA, sprEzeA.MaxCols - 3, 2

End Sub

Private Sub L_LlenarGrillaResumen()
Dim fecha As String
Dim i As Integer
Dim Total As Variant
Dim TOTImp As Long
Dim totCant As Long


SprResumen.MaxRows = 0

Do While Not RsDataResumen.EOF
    SprResumen.MaxRows = SprResumen.MaxRows + 1
    
    SprResumen.SetText 1, SprResumen.MaxRows, Trim(RsDataResumen!Descrip)
    SprResumen.SetText 2, SprResumen.MaxRows, str(RsDataResumen!imp)
    SprResumen.SetText 3, SprResumen.MaxRows, str(RsDataResumen!cant_t)
    SprResumen.SetText 4, SprResumen.MaxRows, str(RsDataResumen!cant_p)
    
    RsDataResumen.MoveNext
Loop

Spread_TotalesGrillas SprResumen, SprResumen.MaxCols - 3, 2

End Sub

Private Sub L_LlenarGrillaINTB()

sprEzeB.MaxRows = 0

Do While Not RsDataEzeB.EOF
    sprEzeB.MaxRows = sprEzeB.MaxRows + 1
    
    sprEzeB.SetText 1, sprEzeB.MaxRows, Trim(RsDataEzeB!Descrip)
    sprEzeB.SetText 2, sprEzeB.MaxRows, str(RsDataEzeB!imp)
    sprEzeB.SetText 3, sprEzeB.MaxRows, str(RsDataEzeB!cant_t)
    sprEzeB.SetText 4, sprEzeB.MaxRows, str(RsDataEzeB!cant_p)
    
    RsDataEzeB.MoveNext
Loop

Spread_TotalesGrillas sprEzeB, sprEzeB.MaxCols - 3, 2
End Sub


Private Sub L_LlenarGrillaAEP()
Dim fecha As String
Dim i As Integer
Dim Total As Variant
Dim TOTImp As Long
Dim totCant As Long


sprAep.MaxRows = 0

Do While Not RsDataAep.EOF
    sprAep.MaxRows = sprAep.MaxRows + 1
    
    sprAep.SetText 1, sprAep.MaxRows, Trim(RsDataAep!Descrip)
    sprAep.SetText 2, sprAep.MaxRows, str(RsDataAep!imp)
    sprAep.SetText 3, sprAep.MaxRows, str(RsDataAep!cant_t)
    sprAep.SetText 4, sprAep.MaxRows, str(RsDataAep!cant_p)
    
    RsDataAep.MoveNext
Loop

Spread_TotalesGrillas sprAep, sprAep.MaxCols - 3, 2

''MsgBox ("Termino")

End Sub

Private Sub optSort_Click(Index As Integer)


Select Case Index
    Case 0
        sprEzeA.Row = 1
        sprEzeA.Col = 1
        sprEzeA.Row2 = sprEzeA.MaxRows - 1
        sprEzeA.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeA.SortBy = SS_SORT_BY_ROW

        sprEzeA.SortKey(1) = 1
        sprEzeA.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprEzeA.Action = SS_ACTION_SORT
    Case 1
        sprEzeA.Row = 1
        sprEzeA.Col = 1
        sprEzeA.Row2 = sprEzeA.MaxRows - 1
        sprEzeA.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeA.SortBy = SS_SORT_BY_ROW
        
        sprEzeA.SortKey(1) = 2
        sprEzeA.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprEzeA.Action = SS_ACTION_SORT
    Case 2
        sprEzeA.Row = 1
        sprEzeA.Col = 1
        sprEzeA.Row2 = sprEzeA.MaxRows - 1
        sprEzeA.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeA.SortBy = SS_SORT_BY_ROW
        
        sprEzeA.SortKey(1) = 2
        sprEzeA.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprEzeA.Action = SS_ACTION_SORT
    Case 3
        sprEzeA.Row = 1
        sprEzeA.Col = 1
        sprEzeA.Row2 = sprEzeA.MaxRows - 1
        sprEzeA.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeA.SortBy = SS_SORT_BY_ROW
        
        sprEzeA.SortKey(1) = 5
        sprEzeA.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprEzeA.Action = SS_ACTION_SORT
            
    Case 4
        sprEzeB.Row = 1
        sprEzeB.Col = 1
        sprEzeB.Row2 = sprEzeB.MaxRows - 1
        sprEzeB.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeB.SortBy = SS_SORT_BY_ROW
        
        sprEzeB.SortKey(1) = 1
        sprEzeB.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprEzeB.Action = SS_ACTION_SORT
    Case 5
        sprEzeB.Row = 1
        sprEzeB.Col = 1
        sprEzeB.Row2 = sprEzeB.MaxRows - 1
        sprEzeB.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeB.SortBy = SS_SORT_BY_ROW
        sprEzeB.SortKey(1) = 2
        sprEzeB.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprEzeB.Action = SS_ACTION_SORT
    Case 6
        sprEzeB.Row = 1
        sprEzeB.Col = 1
        sprEzeB.Row2 = sprEzeB.MaxRows - 1
        sprEzeB.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeB.SortBy = SS_SORT_BY_ROW
        sprEzeB.SortKey(1) = 2
        sprEzeB.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprEzeB.Action = SS_ACTION_SORT
    Case 7
        sprEzeB.Row = 1
        sprEzeB.Col = 1
        sprEzeB.Row2 = sprEzeB.MaxRows - 1
        sprEzeB.Col2 = 5
        
        ' Set sort definition for key 1
        sprEzeB.SortBy = SS_SORT_BY_ROW
        sprEzeB.SortKey(1) = 5
        sprEzeB.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprEzeB.Action = SS_ACTION_SORT
    
    Case 8
        sprAep.Row = 1
        sprAep.Col = 1
        sprAep.Row2 = sprAep.MaxRows - 1
        sprAep.Col2 = 5
        
        ' Set sort definition for key 1
        sprAep.SortBy = SS_SORT_BY_ROW
        sprAep.SortKey(1) = 1
        sprAep.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprAep.Action = SS_ACTION_SORT
    Case 9
        sprAep.Row = 1
        sprAep.Col = 1
        sprAep.Row2 = sprAep.MaxRows - 1
        sprAep.Col2 = 5
        
        ' Set sort definition for key 1
        sprAep.SortBy = SS_SORT_BY_ROW
        sprAep.SortKey(1) = 2
        sprAep.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprAep.Action = SS_ACTION_SORT
    Case 10
        sprAep.Row = 1
        sprAep.Col = 1
        sprAep.Row2 = sprAep.MaxRows - 1
        sprAep.Col2 = 5
        
        ' Set sort definition for key 1
        sprAep.SortBy = SS_SORT_BY_ROW
        sprAep.SortKey(1) = 2
        sprAep.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprAep.Action = SS_ACTION_SORT
    Case 11
        sprAep.Row = 1
        sprAep.Col = 1
        sprAep.Row2 = sprAep.MaxRows - 1
        sprAep.Col2 = 5
        
        ' Set sort definition for key 1
        sprAep.SortBy = SS_SORT_BY_ROW
        sprAep.SortKey(1) = 5
        sprAep.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprAep.Action = SS_ACTION_SORT
        
    Case 13
        SprResumen.Row = 1
        SprResumen.Col = 1
        SprResumen.Row2 = SprResumen.MaxRows - 1
        SprResumen.Col2 = 5
        
        ' Set sort definition for key 1
        SprResumen.SortBy = SS_SORT_BY_ROW
        SprResumen.SortKey(1) = 1
        SprResumen.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        SprResumen.Action = SS_ACTION_SORT
    
    Case 12
        SprResumen.Row = 1
        SprResumen.Col = 1
        SprResumen.Row2 = SprResumen.MaxRows - 1
        SprResumen.Col2 = 5
        
        ' Set sort definition for key 1
        SprResumen.SortBy = SS_SORT_BY_ROW
        SprResumen.SortKey(1) = 2
        SprResumen.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        SprResumen.Action = SS_ACTION_SORT
    
    Case 14
        SprResumen.Row = 1
        SprResumen.Col = 1
        SprResumen.Row2 = SprResumen.MaxRows - 1
        SprResumen.Col2 = 5
        
        ' Set sort definition for key 1
        SprResumen.SortBy = SS_SORT_BY_ROW
        SprResumen.SortKey(1) = 2
        SprResumen.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        SprResumen.Action = SS_ACTION_SORT
    
    Case 15
        SprResumen.Row = 1
        SprResumen.Col = 1
        SprResumen.Row2 = SprResumen.MaxRows - 1
        SprResumen.Col2 = 5
        
        ' Set sort definition for key 1
        SprResumen.SortBy = SS_SORT_BY_ROW
        SprResumen.SortKey(1) = 5
        SprResumen.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        SprResumen.Action = SS_ACTION_SORT
    
    End Select
End Sub


