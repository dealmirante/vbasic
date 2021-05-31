VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGLC 
   Caption         =   "Ventas por Grupo"
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
      Picture         =   "frmGLC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   94
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
      Picture         =   "frmGLC.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   93
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
      Picture         =   "frmGLC.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   30
      Width           =   570
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   4185
      Left            =   60
      TabIndex        =   9
      Top             =   1395
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   7382
      _Version        =   327680
      Tabs            =   4
      Tab             =   1
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
      TabPicture(0)   =   "frmGLC.frx":0A26
      Tab(0).ControlCount=   0
      Tab(0).ControlEnabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmGLC.frx":0A42
      Tab(1).ControlCount=   15
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "chkEzeA(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "tabEzeA"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "botGrTortaEA"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "botGrafEvoEZEA"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkEzeA(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chkEzeA(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkEzeA(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkEzeA(4)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkEzeA(5)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkEzeA(6)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "chkEzeA(7)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "botExcelEzeA"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "botPorc(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "botPorc(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "botPorc(2)"
      Tab(1).Control(14).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmGLC.frx":0A5E
      Tab(2).ControlCount=   12
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkEzeB(4)"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "botPorcB(2)"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "botPorcB(1)"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "botPorcB(0)"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "botExcelEzeB"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "chkEzeB(3)"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "chkEzeB(2)"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "chkEzeB(1)"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "chkEzeB(0)"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "botGrEvoEzeB"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "botGrTortaEB"
      Tab(2).Control(10).Enabled=   -1  'True
      Tab(2).Control(11)=   "tabEzeB"
      Tab(2).Control(11).Enabled=   0   'False
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmGLC.frx":0A7A
      Tab(3).ControlCount=   10
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tabAep"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "botGrTortaAEP"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "botGrEvoAep"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "chkAep(0)"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "chkAep(1)"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "botExcelAep"
      Tab(3).Control(5).Enabled=   -1  'True
      Tab(3).Control(6)=   "botPorcAep(0)"
      Tab(3).Control(6).Enabled=   -1  'True
      Tab(3).Control(7)=   "botPorcAep(1)"
      Tab(3).Control(7).Enabled=   -1  'True
      Tab(3).Control(8)=   "botPorcAep(2)"
      Tab(3).Control(8).Enabled=   -1  'True
      Tab(3).Control(9)=   "chkAep(2)"
      Tab(3).Control(9).Enabled=   -1  'True
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
         Left            =   -70335
         TabIndex        =   97
         Top             =   450
         Width           =   720
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
         TabIndex        =   95
         Top             =   450
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
         Left            =   -67500
         Picture         =   "frmGLC.frx":0A96
         Style           =   1  'Graphical
         TabIndex        =   91
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
         Picture         =   "frmGLC.frx":115C
         Style           =   1  'Graphical
         TabIndex        =   90
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
         Picture         =   "frmGLC.frx":17B2
         Style           =   1  'Graphical
         TabIndex        =   89
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
         Picture         =   "frmGLC.frx":2124
         Style           =   1  'Graphical
         TabIndex        =   88
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
         Picture         =   "frmGLC.frx":27EA
         Style           =   1  'Graphical
         TabIndex        =   87
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
         Picture         =   "frmGLC.frx":2E40
         Style           =   1  'Graphical
         TabIndex        =   86
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
         Picture         =   "frmGLC.frx":37B2
         Style           =   1  'Graphical
         TabIndex        =   85
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
         Picture         =   "frmGLC.frx":3E78
         Style           =   1  'Graphical
         TabIndex        =   84
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
         Picture         =   "frmGLC.frx":44CE
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   450
         Width           =   765
      End
      Begin VB.CommandButton botExcelAep 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67500
         Picture         =   "frmGLC.frx":4E40
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   3135
         Width           =   765
      End
      Begin VB.CommandButton botExcelEzeB 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67500
         Picture         =   "frmGLC.frx":53D2
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   3150
         Width           =   765
      End
      Begin VB.CommandButton botExcelEzeA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   7500
         Picture         =   "frmGLC.frx":5964
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   3135
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
         TabIndex        =   67
         Top             =   450
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
         TabIndex        =   66
         Top             =   450
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
         Left            =   -71430
         TabIndex        =   65
         Top             =   450
         Width           =   720
      End
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
         Left            =   -72675
         TabIndex        =   64
         Top             =   450
         UseMaskColor    =   -1  'True
         Width           =   1005
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
         Left            =   -73935
         TabIndex        =   63
         Top             =   450
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
         TabIndex        =   62
         Top             =   435
         Value           =   1  'Checked
         Width           =   720
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 08 (2)"
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
         Left            =   5610
         TabIndex        =   61
         Top             =   450
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 08 (1)"
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
         Left            =   5295
         TabIndex        =   60
         Top             =   450
         Visible         =   0   'False
         Width           =   915
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
         Left            =   5055
         TabIndex        =   59
         Top             =   435
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
         Index           =   4
         Left            =   3945
         TabIndex        =   58
         Top             =   465
         Width           =   915
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 22 (S)"
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
         Left            =   2130
         TabIndex        =   57
         Top             =   450
         Width           =   930
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
         Index           =   3
         Left            =   3195
         TabIndex        =   56
         Top             =   450
         Width           =   720
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 21(5)"
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
         Left            =   135
         TabIndex        =   55
         Top             =   450
         Value           =   1  'Checked
         Width           =   900
      End
      Begin VB.CommandButton botGrEvoAep 
         Height          =   500
         Left            =   -67500
         Picture         =   "frmGLC.frx":5EF6
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2595
         Width           =   765
      End
      Begin VB.CommandButton botGrEvoEzeB 
         Height          =   495
         Left            =   -67500
         Picture         =   "frmGLC.frx":6338
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2610
         Width           =   765
      End
      Begin VB.CommandButton botGrafEvoEZEA 
         Height          =   500
         Left            =   7500
         Picture         =   "frmGLC.frx":677A
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2595
         Width           =   765
      End
      Begin VB.CommandButton botGrTortaAEP 
         Height          =   500
         Left            =   -67500
         Picture         =   "frmGLC.frx":6BBC
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2055
         Width           =   765
      End
      Begin VB.CommandButton botGrTortaEB 
         Height          =   510
         Left            =   -67500
         Picture         =   "frmGLC.frx":6FFE
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2055
         Width           =   765
      End
      Begin VB.CommandButton botGrTortaEA 
         Height          =   500
         Left            =   7500
         Picture         =   "frmGLC.frx":7440
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2040
         Width           =   765
      End
      Begin TabDlg.SSTab tabEzeA 
         Height          =   3330
         Left            =   105
         TabIndex        =   10
         Top             =   700
         Width           =   7000
         _ExtentX        =   12356
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmGLC.frx":7882
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frEzeAPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frEzeA0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets (Un.)"
         TabPicture(1)   =   "frmGLC.frx":789E
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeAPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frEzeA1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmGLC.frx":78BA
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frEzeAPorc2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "frEzeA2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick"
         TabPicture(3)   =   "frmGLC.frx":78D6
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprEzeA(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom Imp x Pax"
         TabPicture(4)   =   "frmGLC.frx":78F2
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprEzeA(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2700
            Index           =   4
            Left            =   -74865
            OleObjectBlob   =   "frmGLC.frx":790E
            TabIndex        =   69
            Top             =   225
            Width           =   6600
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2700
            Index           =   3
            Left            =   -74880
            OleObjectBlob   =   "frmGLC.frx":7DC4
            TabIndex        =   68
            Top             =   180
            Width           =   6600
         End
         Begin VB.Frame frEzeA0 
            Height          =   2985
            Left            =   75
            TabIndex        =   13
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":827A
               TabIndex        =   25
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeAPorc0 
            Height          =   2985
            Left            =   90
            TabIndex        =   19
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeAPorc 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":875C
               TabIndex        =   49
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeA1 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   16
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":8B16
               TabIndex        =   43
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeAPorc1 
            Height          =   2985
            Left            =   -74925
            TabIndex        =   20
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeAPorc 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":8FF8
               TabIndex        =   50
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeA2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   31
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":93B2
               TabIndex        =   44
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeAPorc2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   32
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeAPorc 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":9894
               TabIndex        =   33
               Top             =   195
               Width           =   6600
            End
         End
      End
      Begin TabDlg.SSTab tabEzeB 
         Height          =   3330
         Left            =   -74925
         TabIndex        =   11
         Top             =   720
         Width           =   7005
         _ExtentX        =   12356
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmGLC.frx":9C4E
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frEzeBPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frEzeB0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets (Un.)"
         TabPicture(1)   =   "frmGLC.frx":9C6A
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeBPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frEzeB1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmGLC.frx":9C86
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frEzeBPorc2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "frEzeB2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick"
         TabPicture(3)   =   "frmGLC.frx":9CA2
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprEzeB(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom Imp x Pax"
         TabPicture(4)   =   "frmGLC.frx":9CBE
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprEzeB(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2700
            Index           =   4
            Left            =   -74850
            OleObjectBlob   =   "frmGLC.frx":9CDA
            TabIndex        =   71
            Top             =   210
            Width           =   6645
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2700
            Index           =   3
            Left            =   -74850
            OleObjectBlob   =   "frmGLC.frx":A195
            TabIndex        =   70
            Top             =   210
            Width           =   6645
         End
         Begin VB.Frame frEzeB0 
            Height          =   2985
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   6795
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":A650
               TabIndex        =   26
               Top             =   195
               Width           =   6645
            End
         End
         Begin VB.Frame frEzeB1 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   17
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":AB37
               TabIndex        =   45
               Top             =   195
               Width           =   6645
            End
         End
         Begin VB.Frame frEzeB2 
            Height          =   3000
            Left            =   -74910
            TabIndex        =   34
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":B01E
               TabIndex        =   46
               Top             =   180
               Width           =   6645
            End
         End
         Begin VB.Frame frEzeBPorc2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   35
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":B505
               TabIndex        =   52
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeBPorc1 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   22
            Top             =   75
            Width           =   6810
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":B8C4
               TabIndex        =   51
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeBPorc0 
            Height          =   3000
            Left            =   75
            TabIndex        =   21
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":BC83
               TabIndex        =   28
               Top             =   200
               Width           =   6600
            End
         End
      End
      Begin TabDlg.SSTab tabAep 
         Height          =   3330
         Left            =   -74940
         TabIndex        =   12
         Top             =   700
         Width           =   7000
         _ExtentX        =   12356
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmGLC.frx":C042
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frAepPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frAep0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets (Un.)"
         TabPicture(1)   =   "frmGLC.frx":C05E
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frAep1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frAepPorc1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmGLC.frx":C07A
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frAepPorc2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "frAep2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick"
         TabPicture(3)   =   "frmGLC.frx":C096
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprAep(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom Imp x Pax"
         TabPicture(4)   =   "frmGLC.frx":C0B2
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprAep(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprAep 
            Height          =   2700
            Index           =   4
            Left            =   -74865
            OleObjectBlob   =   "frmGLC.frx":C0CE
            TabIndex        =   73
            Top             =   195
            Width           =   6600
         End
         Begin FPSpread.vaSpread sprAep 
            Height          =   2700
            Index           =   3
            Left            =   -74805
            OleObjectBlob   =   "frmGLC.frx":C57A
            TabIndex        =   72
            Top             =   150
            Width           =   6600
         End
         Begin VB.Frame frAep0 
            Height          =   3000
            Left            =   90
            TabIndex        =   15
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":CA26
               TabIndex        =   27
               Top             =   180
               Width           =   6600
            End
         End
         Begin VB.Frame frAep1 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   18
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":CEFE
               TabIndex        =   47
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frAep2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   36
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":D3C8
               TabIndex        =   48
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frAepPorc2 
            Height          =   3000
            Left            =   -74925
            TabIndex        =   37
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   2
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":D892
               TabIndex        =   54
               Top             =   200
               Width           =   6615
            End
         End
         Begin VB.Frame frAepPorc1 
            Height          =   2985
            Left            =   -74940
            TabIndex        =   24
            Top             =   30
            Width           =   6810
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   1
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":DC28
               TabIndex        =   53
               Top             =   200
               Width           =   6615
            End
         End
         Begin VB.Frame frAepPorc0 
            Height          =   2985
            Left            =   60
            TabIndex        =   23
            Top             =   45
            Width           =   6825
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmGLC.frx":DFBE
               TabIndex        =   29
               Top             =   200
               Width           =   6615
            End
         End
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 21(6)"
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
         Left            =   1020
         TabIndex        =   96
         Top             =   450
         Width           =   900
      End
   End
   Begin VB.Frame frdatos 
      Height          =   1380
      Left            =   15
      TabIndex        =   0
      Top             =   -45
      Width           =   7560
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   2490
         Picture         =   "frmGLC.frx":E354
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   690
         Width           =   375
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2505
         Picture         =   "frmGLC.frx":E4C6
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   210
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
         Height          =   1140
         Index           =   1
         Left            =   3045
         TabIndex        =   1
         Top             =   120
         Width           =   4125
         Begin VB.ComboBox cboComi 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   675
            Width           =   2340
         End
         Begin VB.TextBox txtComi 
            Height          =   285
            Left            =   3885
            TabIndex        =   7
            Top             =   150
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.ListBox lstComi 
            Height          =   255
            Left            =   4395
            TabIndex        =   6
            Top             =   195
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
            Left            =   165
            TabIndex        =   5
            Top             =   495
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
            Left            =   1485
            TabIndex        =   4
            Top             =   210
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
            Left            =   165
            TabIndex        =   3
            Top             =   780
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
            Left            =   165
            TabIndex        =   2
            Top             =   195
            Value           =   -1  'True
            Width           =   1005
         End
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1335
         TabIndex        =   79
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
         TabIndex        =   80
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
         TabIndex        =   82
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
         TabIndex        =   81
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmGLC"
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

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas
                       'con porcentajes
Dim UsoPorIntB As Boolean
Dim UsoPorAero As Boolean
Dim UsoPorTotal As Boolean

Private Function L_Armarcondicion(Esp As String) As String
Dim Cond As String, condLoc As String
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant, i

fechaDesde = mskFDesde.FormattedText
fechaHasta = mskFHasta.FormattedText

Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)


If optComi(1).Value Then
    Cond = Cond & " and  Comitente IN ('IOSC','IBAIR') "
ElseIf optComi(0).Value Then
    Cond = Cond & " and  Comitente NOT IN ('IOSC','IBAIR','T') "
ElseIf optComi(2).Value Then
    Cond = Cond & " and  Comitente = '" & lstComi.List(cboComi.ListIndex) & "'"
Else
    Cond = Cond & " and  Comitente = 'T' "
End If

Select Case Esp
    Case EZEA
        
        condLoc = ""
        For i = 0 To 4
          If i = 0 Then
            If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L21' and cod_sloc = 0 ) "
            End If
          ElseIf i = 1 Then
            If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L21' and cod_sloc = 1 ) "
            End If
          ElseIf i = 2 Then
            If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L22' and cod_sloc in (0,1,2) ) "
            End If
          ElseIf i = 3 Then
            If chkEzeA(i).Value = 1 Then
                'condLoc = condLoc & " or ( cod_local = '" & Left(DSLoc(2).locales(i), 3) & "' and cod_sloc = " & Right(DSLoc(2).locales(i), 1) & ") "
                condLoc = condLoc & " or ( cod_local = 'L06' )"
            End If
          ElseIf i = 4 Then
            If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L22' and cod_sloc in (3,4) ) "
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
        For i = 0 To NroLocINTB - 1
            If chkEzeB(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = '" & Left(DSLoc(3).locales(i + 1), 3) & "' and cod_sloc = " & Right(DSLoc(3).locales(i + 1), 1) & ") "
                'condLoc = condLoc & " or ( cod_local = '" & chkEzeB(i).caption & "') "
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

Private Function L_ArmarcondicionPax(Esp As String) As String
Dim Cond, condLoc As String
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant, i

fechaDesde = mskFDesde.FormattedText
fechaHasta = mskFHasta.FormattedText

Cond = " WHERE fch_vta between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

Select Case Esp
    Case EZEA
        
        condLoc = ""
        For i = 0 To 2
            
          If i = 0 Then
            If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L21' and cod_sloc = 0 ) "
            End If
          ElseIf i = 1 Then
            If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L21' and cod_sloc = 1 ) "
            End If
          ElseIf i = 2 Then
            If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = 'L22' and cod_sloc = 0 ) "
            End If
          Else
            If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = '" & Left(DSLoc(2).locales(i + 1), 3) & "' and cod_sloc = " & Right(DSLoc(2).locales(i + 1), 1) & ") "
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
        For i = 0 To NroLocINTB - 1
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

L_ArmarcondicionPax = Cond

End Function


Private Sub L_LimpiarGrillas()
Dim i
For i = 0 To 4
    sprEzeA(i).MaxRows = 0
    sprEzeB(i).MaxRows = 0
    sprAEP(i).MaxRows = 0
Next


End Sub

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
        For i = 0 To 6
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
    frEzeAPorc2.Refresh
    frEzeBPorc0.Refresh
    frEzeBPorc1.Refresh
    frEzeBPorc2.Refresh
    frAepPorc0.Refresh
    frAepPorc1.Refresh
    frAepPorc2.Refresh
    frEzeA0.Refresh
    frEzeA1.Refresh
    frEzeA2.Refresh
    frEzeB0.Refresh
    frEzeB1.Refresh
    frEzeB2.Refresh
    frAep0.Refresh
    frAep1.Refresh
    frAep2.Refresh

End Sub

Private Sub L_RefrescarEzeA()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i
On Error GoTo ErrGLC:

frmGLC.caption = Aplicacion.SeteoProceso(frmGLC.caption)

sql = " SELECT "
sql = sql & " fch_ticket, "
sql = sql & " grupo_venta, "
sql = sql & " sum(cantidad) cant, "
sql = sql & " sum(importe) imp "
sql = sql & "FROM " & funcLocal_Vista("venta_lgi", Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion(EZEA)
sql = sql & "group by fch_ticket,grupo_venta "
sql = sql & " order by fch_ticket,grupo_venta"

sqlX = " SELECT "
sqlX = sqlX & " fch_vta, "
sqlX = sqlX & " grupo_venta, "
sqlX = sqlX & " sum(importe) imp, "
sqlX = sqlX & " sum(cant_pax) cant "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_grupo", Year(mskFDesde.FormattedText))
sqlX = sqlX & L_ArmarcondicionPax(EZEA)
sqlX = sqlX & "group by fch_vta,grupo_venta "
sqlX = sqlX & " order by fch_vta,grupo_venta"

If Aplicacion.ObtenerRsDAO(sql, RsDataEzeA) Then
    
        L_LlenarGrillaINTA
    If Aplicacion.CantReg(RsDataEzeA) > 0 Then

        If optComi(3).Value Then
            If Aplicacion.ObtenerRsDAO(sqlX, RsDataEzeAPax) Then
                L_LlenarGrillaINTAPax
            End If
        End If
        FuncLocal_PromediosFilaSPR sprEzeA(0), sprEzeA(1), sprEzeA(3), sprEzeA(0).MaxCols
        FuncLocal_PromediosFilaSPR sprEzeA(0), sprEzeA(2), sprEzeA(4), sprEzeA(0).MaxCols
        
        For i = 0 To 4
            Spread_PintarfinSemana sprEzeA(i)
        Next
    
    End If

    Aplicacion.CerrarDAO RsDataEzeA

ErrGLC:
    frmGLC.caption = Aplicacion.SeteoFin
    Exit Sub
    
End If

End Sub

Private Sub L_RefrescarEzeB()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrGLC:

frmGLC.caption = Aplicacion.SeteoProceso(frmGLC.caption)

sql = " SELECT "
sql = sql & " fch_ticket, "
sql = sql & " grupo_venta, "
sql = sql & " sum(cantidad) cant, "
sql = sql & " sum(importe) imp "
sql = sql & "FROM " & funcLocal_Vista("venta_lgi", Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion(EZEB)
sql = sql & "group by fch_ticket,grupo_venta "
sql = sql & " order by fch_ticket,grupo_venta"

sqlX = " SELECT "
sqlX = sqlX & " fch_vta, "
sqlX = sqlX & " grupo_venta, "
sqlX = sqlX & " sum(importe) imp, "
sqlX = sqlX & " sum(cant_pax) cant "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_grupo", Year(mskFDesde.FormattedText))
sqlX = sqlX & L_ArmarcondicionPax(EZEB)
sqlX = sqlX & "group by fch_vta,grupo_venta "
sqlX = sqlX & " order by fch_vta,grupo_venta"

If Aplicacion.ObtenerRsDAO(sql, RsDataEzeB) Then
    
    L_LlenarGrillaINTB
    If Aplicacion.CantReg(RsDataEzeB) > 0 Then

        If optComi(3).Value Then
            If Aplicacion.ObtenerRsDAO(sqlX, RsDataEzeBPax) Then
                L_LlenarGrillaINTBPax
            End If
        End If
        FuncLocal_PromediosFilaSPR sprEzeB(0), sprEzeB(1), sprEzeB(3), sprEzeB(0).MaxCols
        FuncLocal_PromediosFilaSPR sprEzeB(0), sprEzeB(2), sprEzeB(4), sprEzeB(0).MaxCols
        
        For i = 0 To 4
            Spread_PintarfinSemana sprEzeB(i)
        Next
    
    End If
    Aplicacion.CerrarDAO RsDataEzeB
    
End If

ErrGLC:
    frmGLC.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub

Private Sub L_RefrescarAep()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrGLC:

frmGLC.caption = Aplicacion.SeteoProceso(frmGLC.caption)

sql = " SELECT "
sql = sql & " fch_ticket, "
sql = sql & " grupo_venta, "
sql = sql & " sum(cantidad) cant, "
sql = sql & " sum(importe) imp "
sql = sql & "FROM " & funcLocal_Vista("venta_lgi", Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion(AERO)
sql = sql & "group by fch_ticket,grupo_venta "
sql = sql & " order by fch_ticket,grupo_venta"

sqlX = " SELECT "
sqlX = sqlX & " fch_vta, "
sqlX = sqlX & " grupo_venta, "
sqlX = sqlX & " sum(importe) imp, "
sqlX = sqlX & " sum(cant_pax) cant "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_grupo", Year(mskFDesde.FormattedText))
sqlX = sqlX & L_ArmarcondicionPax(AERO)
sqlX = sqlX & "group by fch_vta,grupo_venta "
sqlX = sqlX & " order by fch_vta,grupo_venta"

If Aplicacion.ObtenerRsDAO(sql, RsDataAep) Then
    
        L_LlenarGrillaAEP
    If Aplicacion.CantReg(RsDataAep) > 0 Then

        If optComi(3).Value Then
            If Aplicacion.ObtenerRsDAO(sqlX, RsDataAepPax) Then
                L_LlenarGrillaAEPPx
            End If
        End If
        FuncLocal_PromediosFilaSPR sprAEP(0), sprAEP(1), sprAEP(3), sprAEP(0).MaxCols
        FuncLocal_PromediosFilaSPR sprAEP(0), sprAEP(2), sprAEP(4), sprAEP(0).MaxCols
        
        For i = 0 To 4
            Spread_PintarfinSemana sprAEP(i)
        Next
    
    End If
        
    Aplicacion.CerrarDAO RsDataAep
    
End If

ErrGLC:
    frmGLC.caption = Aplicacion.SeteoFin
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
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorAero) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub

Private Sub L_SeteosGral()
Dim i

    frdatos.Enabled = False
    botEjecutar(0).Enabled = False
    tabEspigon.Enabled = True

    If optComi(3).Value Then
        For i = 1 To 4
            TabEzeA.TabEnabled(i) = True
            TabEzeB.TabEnabled(i) = True
            TabAep.TabEnabled(i) = True
        Next
    Else
        For i = 1 To 4
            TabEzeA.TabEnabled(i) = False
            TabEzeB.TabEnabled(i) = False
            TabAep.TabEnabled(i) = False
        Next
    End If


End Sub

Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        Select Case Aplicacion.Perfil
            Case "TODOS", "COMUN"
                L_RefrescarEzeA
                L_RefrescarEzeB
                L_RefrescarAep
            Case "INTA"
                L_RefrescarEzeA
            Case "INTA"
                L_RefrescarEzeB
            Case "AEP"
                L_RefrescarAep
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
L_TratarExcel "AEROPARQUE ", "Informe por Grupo (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", AERO, sprAEP(0).MaxCols
End Sub

Private Sub botExcelEzeA_Click()
    L_TratarExcel "ESPIGON INTERNACIONAL A", "Informe por Grupo (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", EZEA, sprEzeA(0).MaxCols
End Sub

Private Sub botExcelEzeb_Click()
    L_TratarExcel "ESPIGON INTERNACIONAL B", "Informe por Grupo (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", EZEB, sprEzeB(0).MaxCols
End Sub

Private Sub L_TratarExcel(titulo As String, subTit As String, Esp As String, CantCol As Integer)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer, filaTit, filaant
Dim i As Integer
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmGLC.caption = Aplicacion.SeteoProceso(frmGLC.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
 '   AppExcel.application.Visible = True
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
               
    fila = fila + 2
    
    Exl_PonerValor AppExcel, fila, Col, "Locales : " & L_Locales(Esp)
        
    filaTit = fila
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

           Next
    
    End Select
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    'AppExcel.application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
        
    AppExcel.SaveAs NOMBRE & ".xls"
'    AppExcel.Workbooks.Open nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmGLC.caption = Aplicacion.SeteoFin
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

Private Sub botGrEvoAep_Click()
    
    Select Case TabAep.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico AERO, "Importes", sprAEP(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico AERO, "Tickets", sprAEP(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico AERO, "Pasajeros", sprAEP(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select


End Sub

Private Sub botGrEvoEzeB_Click()
    
    Select Case TabEzeB.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico EZEB, "Importes", sprEzeB(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico EZEB, "Tickets", sprEzeB(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico EZEB, "Pasajeros", sprEzeB(2), mskFDesde.FormattedText, mskFHasta.FormattedText
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



Private Sub chkAep_Click(Index As Integer)
L_RefrescarAep
End Sub

Private Sub chkEzeA_Click(Index As Integer)
L_RefrescarEzeA
End Sub

Private Sub chkEzeB_Click(Index As Integer)
L_RefrescarEzeB
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


FuncCbo_LlenarCbosLST cboComi, lstComi, "COMITENTE"

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

L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmGLC"

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



Private Sub L_LlenarGrillaINTA()
Dim fecha As String
Dim i As Integer
Dim Total As Variant
Dim TOTImp As Long
Dim totCant As Long

For i = 0 To 4
    sprEzeA(i).MaxRows = 0
Next

Do While Not RsDataEzeA.EOF
    fecha = Format$(RsDataEzeA!Fch_Ticket, FTOFECHA)
    
    sprEzeA(0).MaxRows = sprEzeA(0).MaxRows + 1
    sprEzeA(0).SetText 1, sprEzeA(0).MaxRows, Format$(fecha, "dd-mm-yy")
    
    If optComi(3).Value Then
        sprEzeA(1).MaxRows = sprEzeA(1).MaxRows + 1
        sprEzeA(3).MaxRows = sprEzeA(3).MaxRows + 1
    
        sprEzeA(1).SetText 1, sprEzeA(1).MaxRows, Format$(fecha, "dd-mm-yy")
        sprEzeA(3).SetText 1, sprEzeA(3).MaxRows, Format$(fecha, "dd-mm-yy")
    End If
    
    TOTImp = 0
    totCant = 0
    
    Do While fecha = Format$(RsDataEzeA!Fch_Ticket, FTOFECHA)
        Select Case RsDataEzeA.Grupo_Venta
            Case "A"
                sprEzeA(0).SetText 2, sprEzeA(0).MaxRows, str(RsDataEzeA!imp)
                If optComi(3).Value Then
                    sprEzeA(1).SetText 2, sprEzeA(1).MaxRows, str(RsDataEzeA!cant)
                    If RsDataEzeA!cant > 0 Then
                    sprEzeA(3).SetText 2, sprEzeA(3).MaxRows, str(RsDataEzeA!imp / RsDataEzeA!cant)
                    End If
                End If
                
            Case "B"
                sprEzeA(0).SetText 3, sprEzeA(0).MaxRows, str(RsDataEzeA!imp)
                If optComi(3).Value Then
                    sprEzeA(1).SetText 3, sprEzeA(1).MaxRows, str(RsDataEzeA!cant)
                    If RsDataEzeA!cant > 0 Then
                    sprEzeA(3).SetText 3, sprEzeA(3).MaxRows, str(RsDataEzeA!imp / RsDataEzeA!cant)
                    End If
                End If
                
            Case "C"
                sprEzeA(0).SetText 4, sprEzeA(0).MaxRows, str(RsDataEzeA!imp)
                If optComi(3).Value Then
                    sprEzeA(1).SetText 4, sprEzeA(1).MaxRows, str(RsDataEzeA!cant)
                    If RsDataEzeA!cant > 0 Then
                    sprEzeA(3).SetText 4, sprEzeA(3).MaxRows, str(RsDataEzeA!imp / RsDataEzeA!cant)
                    End If
                End If
                
        End Select
        TOTImp = TOTImp + RsDataEzeA!imp
        totCant = totCant + RsDataEzeA!cant
        
        RsDataEzeA.MoveNext
        If RsDataEzeA.EOF Then
            Exit Do
        End If
    Loop
    If optComi(3).Value Then
        sprEzeA(3).SetText 5, sprEzeA(3).MaxRows, Format$((TOTImp / totCant), "#.0")
    End If
Loop
For i = 0 To 1
    Spread_TotalesGrillas sprEzeA(i), sprEzeA(i).MaxCols - 1, 2
Next

End Sub

Private Sub L_LlenarGrillaINTAPax()
Dim fecha As String
Dim i As Integer
Dim Total As Variant
Dim TOTImp As Long
Dim totCant As Long


Do While Not RsDataEzeAPax.EOF
    fecha = Format$(RsDataEzeAPax!fch_vta, FTOFECHA)
    
    sprEzeA(2).MaxRows = sprEzeA(2).MaxRows + 1
    sprEzeA(4).MaxRows = sprEzeA(4).MaxRows + 1
    
    sprEzeA(2).SetText 1, sprEzeA(2).MaxRows, Format$(fecha, "dd-mm-yy")
    sprEzeA(4).SetText 1, sprEzeA(4).MaxRows, Format$(fecha, "dd-mm-yy")
    
    TOTImp = 0
    totCant = 0
    
    Do While fecha = Format$(RsDataEzeAPax!fch_vta, FTOFECHA)
        Select Case RsDataEzeAPax.Grupo_Venta
            Case "A"
                sprEzeA(2).SetText 2, sprEzeA(2).MaxRows, str(RsDataEzeAPax!cant)
                If RsDataEzeAPax!cant > 0 Then
                sprEzeA(4).SetText 2, sprEzeA(4).MaxRows, str(RsDataEzeAPax!imp / RsDataEzeAPax!cant)
                End If
            Case "B"
                sprEzeA(2).SetText 3, sprEzeA(2).MaxRows, str(RsDataEzeAPax!cant)
                If RsDataEzeAPax!cant > 0 Then
                sprEzeA(4).SetText 3, sprEzeA(4).MaxRows, str(RsDataEzeAPax!imp / RsDataEzeAPax!cant)
                End If
            Case "C"
                sprEzeA(2).SetText 4, sprEzeA(2).MaxRows, str(RsDataEzeAPax!cant)
                If RsDataEzeAPax!cant > 0 Then
                sprEzeA(4).SetText 4, sprEzeA(4).MaxRows, str(RsDataEzeAPax!imp / RsDataEzeAPax!cant)
                End If
        End Select
        TOTImp = TOTImp + RsDataEzeAPax!imp
        totCant = totCant + RsDataEzeAPax!cant
        
        RsDataEzeAPax.MoveNext
        If RsDataEzeAPax.EOF Then
            Exit Do
        End If
    Loop
    sprEzeA(4).SetText 5, sprEzeA(4).MaxRows, Format$((TOTImp / totCant), "#.0")
Loop
    Spread_TotalesGrillas sprEzeA(2), sprEzeA(2).MaxCols - 1, 2
End Sub


Private Sub L_LlenarGrillaINTB()
Dim depn As String, Sdep As String, fecha As String
Dim i As Integer
Dim Total As Variant
Dim TOTImp As Long
Dim totCant As Long


For i = 0 To 4
    sprEzeB(i).MaxRows = 0
Next

Do While Not RsDataEzeB.EOF
    fecha = Format$(RsDataEzeB!Fch_Ticket, FTOFECHA)
    
    sprEzeB(0).MaxRows = sprEzeB(0).MaxRows + 1
    sprEzeB(0).SetText 1, sprEzeB(0).MaxRows, Format$(fecha, "dd-mm-yy")
    
    If optComi(3).Value Then
        sprEzeB(1).MaxRows = sprEzeB(1).MaxRows + 1
        sprEzeB(3).MaxRows = sprEzeB(3).MaxRows + 1
            
        sprEzeB(1).SetText 1, sprEzeB(1).MaxRows, Format$(fecha, "dd-mm-yy")
        sprEzeB(3).SetText 1, sprEzeB(3).MaxRows, Format$(fecha, "dd-mm-yy")
    End If
    TOTImp = 0
    totCant = 0

    Do While fecha = Format$(RsDataEzeB!Fch_Ticket, FTOFECHA)
        Select Case RsDataEzeB.Grupo_Venta
            Case "A"
                sprEzeB(0).SetText 2, sprEzeB(0).MaxRows, str(RsDataEzeB!imp)
                If optComi(3).Value Then
                    sprEzeB(1).SetText 2, sprEzeB(1).MaxRows, str(RsDataEzeB!cant)
                    If RsDataEzeB!cant > 0 Then
                    sprEzeB(3).SetText 2, sprEzeB(3).MaxRows, str(RsDataEzeB!imp / RsDataEzeB!cant)
                    End If
                End If
            Case "B"
                sprEzeB(0).SetText 3, sprEzeB(0).MaxRows, str(RsDataEzeB!imp)
                If optComi(3).Value Then
                    sprEzeB(1).SetText 3, sprEzeB(1).MaxRows, str(RsDataEzeB!cant)
                    If RsDataEzeB!cant > 0 Then
                    sprEzeB(3).SetText 3, sprEzeB(3).MaxRows, str(RsDataEzeB!imp / RsDataEzeB!cant)
                    End If
                End If
            Case "C"
                sprEzeB(0).SetText 4, sprEzeB(0).MaxRows, str(RsDataEzeB!imp)
                If optComi(3).Value Then
                    sprEzeB(1).SetText 4, sprEzeB(1).MaxRows, str(RsDataEzeB!cant)
                    If RsDataEzeB!cant > 0 Then
                    sprEzeB(3).SetText 4, sprEzeB(3).MaxRows, str(RsDataEzeB!imp / RsDataEzeB!cant)
                    End If
                End If
        End Select
        TOTImp = TOTImp + RsDataEzeB!imp
        totCant = totCant + RsDataEzeB!cant

        RsDataEzeB.MoveNext
        If RsDataEzeB.EOF Then
            Exit Do
        End If
    Loop
    If optComi(3).Value Then
        sprEzeB(3).SetText 5, sprEzeB(3).MaxRows, Format$((TOTImp / totCant), "#.0")
    End If
Loop
For i = 0 To 1
    Spread_TotalesGrillas sprEzeB(i), sprEzeB(i).MaxCols - 1, 2
Next


End Sub
Private Sub L_LlenarGrillaINTBPax()
Dim fecha As String
Dim i As Integer
Dim Total As Variant
Dim TOTImp As Long
Dim totCant As Long


Do While Not RsDataEzeBPax.EOF
    fecha = Format$(RsDataEzeBPax!fch_vta, FTOFECHA)
    
    sprEzeB(2).MaxRows = sprEzeB(2).MaxRows + 1
    sprEzeB(4).MaxRows = sprEzeB(4).MaxRows + 1
    
    sprEzeB(2).SetText 1, sprEzeB(2).MaxRows, Format$(fecha, "dd-mm-yy")
    sprEzeB(4).SetText 1, sprEzeB(4).MaxRows, Format$(fecha, "dd-mm-yy")
    
    TOTImp = 0
    totCant = 0
    
    Do While fecha = Format$(RsDataEzeBPax!fch_vta, FTOFECHA)
        Select Case RsDataEzeBPax.Grupo_Venta
            Case "A"
                sprEzeB(2).SetText 2, sprEzeB(2).MaxRows, str(RsDataEzeBPax!cant)
                If RsDataEzeBPax!cant > 0 Then
                sprEzeB(4).SetText 2, sprEzeB(4).MaxRows, str(RsDataEzeBPax!imp / RsDataEzeBPax!cant)
                End If
                
            Case "B"
                sprEzeB(2).SetText 3, sprEzeB(2).MaxRows, str(RsDataEzeBPax!cant)
                If RsDataEzeBPax!cant > 0 Then
                sprEzeB(4).SetText 3, sprEzeB(4).MaxRows, str(RsDataEzeBPax!imp / RsDataEzeBPax!cant)
                End If
                
            Case "C"
                sprEzeB(2).SetText 4, sprEzeB(2).MaxRows, str(RsDataEzeBPax!cant)
                If RsDataEzeBPax!cant > 0 Then
                sprEzeB(4).SetText 4, sprEzeB(4).MaxRows, str(RsDataEzeBPax!imp / RsDataEzeBPax!cant)
                End If
        End Select
        TOTImp = TOTImp + RsDataEzeBPax!imp
        totCant = totCant + RsDataEzeBPax!cant
        
        RsDataEzeBPax.MoveNext
        If RsDataEzeBPax.EOF Then
            Exit Do
        End If
    Loop
    sprEzeB(4).SetText 5, sprEzeB(4).MaxRows, Format$((TOTImp / totCant), "#.0")
Loop
    Spread_TotalesGrillas sprEzeB(2), sprEzeB(2).MaxCols - 1, 2
End Sub


Private Sub L_LlenarGrillaAEP()
Dim depn As String, Sdep As String, fecha As String
Dim i As Integer
Dim Total As Variant
Dim TOTImp As Long
Dim totCant As Long


For i = 0 To 4
    sprAEP(i).MaxRows = 0
Next

Do While Not RsDataAep.EOF
    fecha = Format$(RsDataAep!Fch_Ticket, FTOFECHA)
    
    sprAEP(0).MaxRows = sprAEP(0).MaxRows + 1
    sprAEP(0).SetText 1, sprAEP(0).MaxRows, Format$(fecha, "dd-mm-yy")
    
    If optComi(3).Value Then
        sprAEP(1).MaxRows = sprAEP(1).MaxRows + 1
        sprAEP(3).MaxRows = sprAEP(3).MaxRows + 1
        
        sprAEP(1).SetText 1, sprAEP(1).MaxRows, Format$(fecha, "dd-mm-yy")
        sprAEP(3).SetText 1, sprAEP(3).MaxRows, Format$(fecha, "dd-mm-yy")
    End If
    
    TOTImp = 0
    totCant = 0
    Do While fecha = Format$(RsDataAep!Fch_Ticket, FTOFECHA)
        Select Case RsDataAep.Grupo_Venta
            Case "A"
                sprAEP(0).SetText 2, sprAEP(0).MaxRows, str(RsDataAep!imp)
                If optComi(3).Value Then
                    sprAEP(1).SetText 2, sprAEP(1).MaxRows, str(RsDataAep!cant)
                    If RsDataAep!cant > 0 Then
                    sprAEP(3).SetText 2, sprAEP(3).MaxRows, str(RsDataAep!imp / RsDataAep!cant)
                    End If
                End If
            Case "B"
                sprAEP(0).SetText 3, sprAEP(0).MaxRows, str(RsDataAep!imp)
                If optComi(3).Value Then
                    sprAEP(1).SetText 3, sprAEP(1).MaxRows, str(RsDataAep!cant)
                    If RsDataAep!cant > 0 Then
                    sprAEP(3).SetText 3, sprAEP(3).MaxRows, str(RsDataAep!imp / RsDataAep!cant)
                    End If
                End If
            Case "C"
                sprAEP(0).SetText 4, sprAEP(0).MaxRows, str(RsDataAep!imp)
                If optComi(3).Value Then
                    sprAEP(1).SetText 4, sprAEP(1).MaxRows, str(RsDataAep!cant)
                    If RsDataAep!cant > 0 Then
                    sprAEP(3).SetText 4, sprAEP(3).MaxRows, str(RsDataAep!imp / RsDataAep!cant)
                    End If
                End If
        End Select
        TOTImp = TOTImp + RsDataAep!imp
        totCant = totCant + RsDataAep!cant

        RsDataAep.MoveNext
        If RsDataAep.EOF Then
            Exit Do
        End If
    Loop
    If optComi(3).Value Then
        sprAEP(3).SetText 5, sprAEP(3).MaxRows, Format$((TOTImp / totCant), "#.0")
    End If
Loop
For i = 0 To 1
    Spread_TotalesGrillas sprAEP(i), sprAEP(i).MaxCols - 1, 2
Next

End Sub
Private Sub L_LlenarGrillaAEPPx()
Dim fecha As String
Dim i As Integer
Dim Total As Variant
Dim TOTImp As Long
Dim totCant As Long


Do While Not RsDataAepPax.EOF
    fecha = Format$(RsDataAepPax!fch_vta, FTOFECHA)
    
    sprAEP(2).MaxRows = sprAEP(2).MaxRows + 1
    sprAEP(4).MaxRows = sprAEP(4).MaxRows + 1
    
    sprAEP(2).SetText 1, sprAEP(2).MaxRows, Format$(fecha, "dd-mm-yy")
    sprAEP(4).SetText 1, sprAEP(4).MaxRows, Format$(fecha, "dd-mm-yy")
    
    TOTImp = 0
    totCant = 0
    
    Do While fecha = Format$(RsDataAepPax!fch_vta, FTOFECHA)
        Select Case RsDataAepPax.Grupo_Venta
            Case "A"
                sprAEP(2).SetText 2, sprAEP(2).MaxRows, str(RsDataAepPax!cant)
                If RsDataAepPax!cant > 0 Then
                sprAEP(4).SetText 2, sprAEP(4).MaxRows, str(RsDataAepPax!imp / RsDataAepPax!cant)
                End If
            Case "B"
                sprAEP(2).SetText 3, sprAEP(2).MaxRows, str(RsDataAepPax!cant)
                If RsDataAepPax!cant > 0 Then
                sprAEP(4).SetText 3, sprAEP(4).MaxRows, str(RsDataAepPax!imp / RsDataAepPax!cant)
                End If
            Case "C"
                sprAEP(2).SetText 4, sprAEP(2).MaxRows, str(RsDataAepPax!cant)
                If RsDataAepPax!cant > 0 Then
                sprAEP(4).SetText 4, sprAEP(4).MaxRows, str(RsDataAepPax!imp / RsDataAepPax!cant)
                End If
        End Select
        TOTImp = TOTImp + RsDataAepPax!imp
        totCant = totCant + RsDataAepPax!cant
        
        RsDataAepPax.MoveNext
        If RsDataAepPax.EOF Then
            Exit Do
        End If
    Loop
    sprAEP(4).SetText 5, sprAEP(4).MaxRows, Format$((TOTImp / totCant), "#.0")
Loop
    Spread_TotalesGrillas sprAEP(2), sprAEP(2).MaxCols - 1, 2
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
        botGrEvoAep.Visible = True
    Case 3, 4
        botPorcAep(0).Visible = False
        botPorcAep(1).Visible = False
        botPorcAep(2).Visible = False
        botGrTortaAEP.Visible = False
        botGrEvoAep.Visible = False
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
        botGrEvoEzeB.Visible = True
    Case 3, 4
        botPorcB(0).Visible = False
        botPorcB(1).Visible = False
        botPorcB(2).Visible = False
        botGrTortaEB.Visible = False
        botGrEvoEzeB.Visible = False
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

