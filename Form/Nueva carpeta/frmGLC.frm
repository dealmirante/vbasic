VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
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
      Tab(1).Control(0)=   "tabEzeA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "botGrTortaEA"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "botGrafEvoEZEA"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkEzeA(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkEzeA(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chkEzeA(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkEzeA(4)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkEzeA(5)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkEzeA(6)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkEzeA(7)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "botExcelEzeA"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "botPorc(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "botPorc(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "botPorc(2)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chkEzeA(1)"
      Tab(1).Control(14).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmGLC.frx":0A5E
      Tab(2).ControlCount=   11
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabEzeB"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "botGrTortaEB"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "botGrEvoEzeB"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "chkEzeB(0)"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "chkEzeB(1)"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "chkEzeB(2)"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "chkEzeB(3)"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "botExcelEzeB"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "botPorcB(0)"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "botPorcB(1)"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "botPorcB(2)"
      Tab(2).Control(10).Enabled=   -1  'True
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
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 05 (1)"
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
         Left            =   930
         TabIndex        =   96
         Top             =   450
         Width           =   900
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
         Index           =   3
         Left            =   -71445
         TabIndex        =   65
         Top             =   450
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
         Index           =   2
         Left            =   -72675
         TabIndex        =   64
         Top             =   450
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
         Top             =   450
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
         Left            =   6345
         TabIndex        =   61
         Top             =   450
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
         Left            =   5355
         TabIndex        =   60
         Top             =   450
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
         Left            =   4605
         TabIndex        =   59
         Top             =   450
         Width           =   720
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
         Left            =   3600
         TabIndex        =   58
         Top             =   450
         Width           =   915
      End
      Begin VB.CheckBox chkEzeA 
         Caption         =   "L 07"
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
         Left            =   2790
         TabIndex        =   57
         Top             =   450
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
         Index           =   2
         Left            =   1980
         TabIndex        =   56
         Top             =   450
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
         Left            =   135
         TabIndex        =   55
         Top             =   450
         Value           =   1  'Checked
         Width           =   720
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
            OleObjectBlob   =   "frmGLC.frx":7DB8
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
               OleObjectBlob   =   "frmGLC.frx":8262
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
               OleObjectBlob   =   "frmGLC.frx":8738
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
               OleObjectBlob   =   "frmGLC.frx":8AE6
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
               OleObjectBlob   =   "frmGLC.frx":8FBC
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
               OleObjectBlob   =   "frmGLC.frx":936A
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
               OleObjectBlob   =   "frmGLC.frx":9840
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
         TabPicture(0)   =   "frmGLC.frx":9BEE
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frEzeBPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frEzeB0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets (Un.)"
         TabPicture(1)   =   "frmGLC.frx":9C0A
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeBPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frEzeB1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmGLC.frx":9C26
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frEzeBPorc2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "frEzeB2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick"
         TabPicture(3)   =   "frmGLC.frx":9C42
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprEzeB(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom Imp x Pax"
         TabPicture(4)   =   "frmGLC.frx":9C5E
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprEzeB(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2700
            Index           =   4
            Left            =   -74850
            OleObjectBlob   =   "frmGLC.frx":9C7A
            TabIndex        =   71
            Top             =   210
            Width           =   6645
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2700
            Index           =   3
            Left            =   -74850
            OleObjectBlob   =   "frmGLC.frx":A129
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
               OleObjectBlob   =   "frmGLC.frx":A5D8
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
               OleObjectBlob   =   "frmGLC.frx":AAB3
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
               OleObjectBlob   =   "frmGLC.frx":AF8E
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
               OleObjectBlob   =   "frmGLC.frx":B469
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
               OleObjectBlob   =   "frmGLC.frx":B81C
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
               OleObjectBlob   =   "frmGLC.frx":BBCF
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
         TabPicture(0)   =   "frmGLC.frx":BF82
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frAepPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frAep0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Tickets (Un.)"
         TabPicture(1)   =   "frmGLC.frx":BF9E
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frAep1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frAepPorc1"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "Pasajeros"
         TabPicture(2)   =   "frmGLC.frx":BFBA
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frAepPorc2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "frAep2"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "Prom Imp x Tick"
         TabPicture(3)   =   "frmGLC.frx":BFD6
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprAep(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Prom Imp x Pax"
         TabPicture(4)   =   "frmGLC.frx":BFF2
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprAep(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprAep 
            Height          =   2700
            Index           =   4
            Left            =   -74865
            OleObjectBlob   =   "frmGLC.frx":C00E
            TabIndex        =   73
            Top             =   195
            Width           =   6600
         End
         Begin FPSpread.vaSpread sprAep 
            Height          =   2700
            Index           =   3
            Left            =   -74805
            OleObjectBlob   =   "frmGLC.frx":C4AE
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
               OleObjectBlob   =   "frmGLC.frx":C94E
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
               OleObjectBlob   =   "frmGLC.frx":CE1A
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
               OleObjectBlob   =   "frmGLC.frx":D2D8
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
               OleObjectBlob   =   "frmGLC.frx":D796
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
               OleObjectBlob   =   "frmGLC.frx":DB20
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
               OleObjectBlob   =   "frmGLC.frx":DEAA
               TabIndex        =   29
               Top             =   200
               Width           =   6615
            End
         End
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
         Picture         =   "frmGLC.frx":E234
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   690
         Width           =   375
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2505
         Picture         =   "frmGLC.frx":E3A6
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
         _Version        =   327680
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
        For i = 0 To NroLocINTA - 1
            If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = '" & Left(DSLoc(2).locales(i + 1), 3) & "' and cod_sloc = " & Right(DSLoc(2).locales(i + 1), 1) & ") "
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
        For i = 0 To NroLocINTA - 1
            If chkEzeA(i).Value = 1 Then
                condLoc = condLoc & " or ( cod_local = '" & Left(DSLoc(2).locales(i + 1), 3) & "' and cod_sloc = " & Right(DSLoc(2).locales(i + 1), 1) & ") "
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
    SprEzeA(i).MaxRows = 0
    SprEzeB(i).MaxRows = 0
    SprAep(i).MaxRows = 0
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
Dim SQL As String
Dim sqlX As String
Dim rs As Recordset
Dim i
On Error GoTo ErrGLC:

frmGLC.caption = Aplicacion.SeteoProceso(frmGLC.caption)

SQL = " SELECT "
SQL = SQL & " fch_ticket, "
SQL = SQL & " grupo_venta, "
SQL = SQL & " sum(cantidad) cant, "
SQL = SQL & " sum(importe) imp "
SQL = SQL & "FROM " & funcLocal_Vista("venta_lgi", Year(mskFDesde.FormattedText))
SQL = SQL & L_Armarcondicion(EZEA)
SQL = SQL & "group by fch_ticket,grupo_venta "
SQL = SQL & " order by fch_ticket,grupo_venta"

sqlX = " SELECT "
sqlX = sqlX & " fch_vta, "
sqlX = sqlX & " grupo_venta, "
sqlX = sqlX & " sum(importe) imp, "
sqlX = sqlX & " sum(cant_pax) cant "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_grupo", Year(mskFDesde.FormattedText))
sqlX = sqlX & L_ArmarcondicionPax(EZEA)
sqlX = sqlX & "group by fch_vta,grupo_venta "
sqlX = sqlX & " order by fch_vta,grupo_venta"

If Aplicacion.ObtenerRsDAO(SQL, RsDataEzeA) Then
    
        L_LlenarGrillaINTA
    If Aplicacion.CantReg(RsDataEzeA) > 0 Then

        If optComi(3).Value Then
            If Aplicacion.ObtenerRsDAO(sqlX, RsDataEzeAPax) Then
                L_LlenarGrillaINTAPax
            End If
        End If
        FuncLocal_PromediosFilaSPR SprEzeA(0), SprEzeA(1), SprEzeA(3), SprEzeA(0).MaxCols
        FuncLocal_PromediosFilaSPR SprEzeA(0), SprEzeA(2), SprEzeA(4), SprEzeA(0).MaxCols
        
        For i = 0 To 4
            Spread_PintarfinSemana SprEzeA(i)
        Next
    
    End If

    Aplicacion.CerrarDAO RsDataEzeA

ErrGLC:
    frmGLC.caption = Aplicacion.SeteoFin
    Exit Sub
    
End If

End Sub

Private Sub L_RefrescarEzeB()
Dim SQL As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrGLC:

frmGLC.caption = Aplicacion.SeteoProceso(frmGLC.caption)

SQL = " SELECT "
SQL = SQL & " fch_ticket, "
SQL = SQL & " grupo_venta, "
SQL = SQL & " sum(cantidad) cant, "
SQL = SQL & " sum(importe) imp "
SQL = SQL & "FROM " & funcLocal_Vista("venta_lgi", Year(mskFDesde.FormattedText))
SQL = SQL & L_Armarcondicion(EZEB)
SQL = SQL & "group by fch_ticket,grupo_venta "
SQL = SQL & " order by fch_ticket,grupo_venta"

sqlX = " SELECT "
sqlX = sqlX & " fch_vta, "
sqlX = sqlX & " grupo_venta, "
sqlX = sqlX & " sum(importe) imp, "
sqlX = sqlX & " sum(cant_pax) cant "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_grupo", Year(mskFDesde.FormattedText))
sqlX = sqlX & L_ArmarcondicionPax(EZEB)
sqlX = sqlX & "group by fch_vta,grupo_venta "
sqlX = sqlX & " order by fch_vta,grupo_venta"

If Aplicacion.ObtenerRsDAO(SQL, RsDataEzeB) Then
    
    L_LlenarGrillaINTB
    If Aplicacion.CantReg(RsDataEzeB) > 0 Then

        If optComi(3).Value Then
            If Aplicacion.ObtenerRsDAO(sqlX, RsDataEzeBPax) Then
                L_LlenarGrillaINTBPax
            End If
        End If
        FuncLocal_PromediosFilaSPR SprEzeB(0), SprEzeB(1), SprEzeB(3), SprEzeB(0).MaxCols
        FuncLocal_PromediosFilaSPR SprEzeB(0), SprEzeB(2), SprEzeB(4), SprEzeB(0).MaxCols
        
        For i = 0 To 4
            Spread_PintarfinSemana SprEzeB(i)
        Next
    
    End If
    Aplicacion.CerrarDAO RsDataEzeB
    
End If

ErrGLC:
    frmGLC.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub

Private Sub L_RefrescarAep()
Dim SQL As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrGLC:

frmGLC.caption = Aplicacion.SeteoProceso(frmGLC.caption)

SQL = " SELECT "
SQL = SQL & " fch_ticket, "
SQL = SQL & " grupo_venta, "
SQL = SQL & " sum(cantidad) cant, "
SQL = SQL & " sum(importe) imp "
SQL = SQL & "FROM " & funcLocal_Vista("venta_lgi", Year(mskFDesde.FormattedText))
SQL = SQL & L_Armarcondicion(AERO)
SQL = SQL & "group by fch_ticket,grupo_venta "
SQL = SQL & " order by fch_ticket,grupo_venta"

sqlX = " SELECT "
sqlX = sqlX & " fch_vta, "
sqlX = sqlX & " grupo_venta, "
sqlX = sqlX & " sum(importe) imp, "
sqlX = sqlX & " sum(cant_pax) cant "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_grupo", Year(mskFDesde.FormattedText))
sqlX = sqlX & L_ArmarcondicionPax(AERO)
sqlX = sqlX & "group by fch_vta,grupo_venta "
sqlX = sqlX & " order by fch_vta,grupo_venta"

If Aplicacion.ObtenerRsDAO(SQL, RsDataAep) Then
    
        L_LlenarGrillaAEP
    If Aplicacion.CantReg(RsDataAep) > 0 Then

        If optComi(3).Value Then
            If Aplicacion.ObtenerRsDAO(sqlX, RsDataAepPax) Then
                L_LlenarGrillaAEPPx
            End If
        End If
        FuncLocal_PromediosFilaSPR SprAep(0), SprAep(1), SprAep(3), SprAep(0).MaxCols
        FuncLocal_PromediosFilaSPR SprAep(0), SprAep(2), SprAep(4), SprAep(0).MaxCols
        
        For i = 0 To 4
            Spread_PintarfinSemana SprAep(i)
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
            Case "TODOS"
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
L_TratarExcel "AEROPARQUE ", "Informe por Grupo (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", AERO, SprAep(0).MaxCols
End Sub

Private Sub botExcelEzeA_Click()
    L_TratarExcel "ESPIGON INTERNACIONAL A", "Informe por Grupo (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", EZEA, SprEzeA(0).MaxCols
End Sub

Private Sub botExcelEzeb_Click()
    L_TratarExcel "ESPIGON INTERNACIONAL B", "Informe por Grupo (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", EZEB, SprEzeB(0).MaxCols
End Sub

Private Sub L_TratarExcel(Titulo As String, subTit As String, Esp As String, CantCol As Integer)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim fila As Integer, filaTit, filaant
Dim i As Integer
Dim tit As Variant
Dim nombre As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

frmGLC.caption = Aplicacion.SeteoProceso(frmGLC.caption)

If nombre <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
 '   AppExcel.application.Visible = True
    ReDim titCol(CantCol)
        
    col = 1
    fila = 3
    
    Exl_PonerValor AppExcel, 1, 1, Titulo
    rango = Exl_rangos(1, 1, 1, CantCol)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, col, subTit
    rango = Exl_rangos(fila, fila, 1, CantCol)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    If optComi(1).Value Then
        Exl_PonerValor AppExcel, fila, col, "Comitente : I.O.S.C y Propios"
    ElseIf optComi(0).Value Then
        Exl_PonerValor AppExcel, fila, col, "Comitente : NO I.O.S.C "
    ElseIf optComi(3).Value Then
        Exl_PonerValor AppExcel, fila, col, "Comitente : TODOS "
    Else
        Exl_PonerValor AppExcel, fila, col, "Comitente : " & cboComi.Text
    End If
    Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
               
    fila = fila + 2
    
    Exl_PonerValor AppExcel, fila, col, "Locales : " & L_Locales(Esp)
        
    filaTit = fila
    Select Case Esp
        Case AERO
            
            For i = 1 To CantCol
                SprAep(0).GetText i, 0, tit
                titCol(i) = tit
            Next
       
           For i = 0 To 4
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, col, "Información sobre :" & L_NombreDato(i + 1)
               
               fila = fila + 2
               filaant = fila
               Exl_BajarGrillaExel SprAep(i), AppExcel, fila, col, titCol
               fila = fila + SprAep(i).MaxRows
               rango = Exl_rangos(fila, fila, col, SprAep(i).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               If i > 2 Then
                   rango = Exl_rangos(filaant + 1, fila, col + 1, SprAep(i).MaxCols)
                   Exl_Format AppExcel, rango
               End If

           Next
        Case EZEA
            For i = 1 To CantCol
                SprEzeA(0).GetText i, 0, tit
                titCol(i) = tit
            Next
       
           For i = 0 To 4
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, col, "Información sobre :" & L_NombreDato(i + 1)
               
               fila = fila + 2
               filaant = fila
               Exl_BajarGrillaExel SprEzeA(i), AppExcel, fila, col, titCol
               fila = fila + SprEzeA(i).MaxRows
               rango = Exl_rangos(fila, fila, col, SprEzeA(i).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               If i > 2 Then
                   rango = Exl_rangos(filaant + 1, fila, col + 1, SprEzeA(i).MaxCols)
                   Exl_Format AppExcel, rango
               End If

           Next
        
        Case EZEB
            For i = 1 To CantCol
                SprEzeB(0).GetText i, 0, tit
                titCol(i) = tit
            Next
       
           For i = 0 To 4
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, col, "Información sobre :" & L_NombreDato(i + 1)
               
               fila = fila + 2
               filaant = fila
               Exl_BajarGrillaExel SprEzeB(i), AppExcel, fila, col, titCol
               fila = fila + SprEzeB(i).MaxRows
               rango = Exl_rangos(fila, fila, col, SprEzeB(i).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               If i > 2 Then
                   rango = Exl_rangos(filaant + 1, fila, col + 1, SprEzeB(i).MaxCols)
                   Exl_Format AppExcel, rango
               End If

           Next
    
    End Select
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    'AppExcel.application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
        
    AppExcel.SaveAs nombre & ".xls"
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
            frmGrafEvoLocal.CargarGrafico EZEA, "Importes", SprEzeA(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico EZEA, "Tickets", SprEzeA(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico EZEA, "Pasajeros", SprEzeA(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select




End Sub

Private Sub botGrEvoAep_Click()
    
    Select Case TabAep.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico AERO, "Importes", SprAep(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico AERO, "Tickets", SprAep(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico AERO, "Pasajeros", SprAep(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select


End Sub

Private Sub botGrEvoEzeB_Click()
    
    Select Case TabEzeB.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico EZEB, "Importes", SprEzeB(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico EZEB, "Tickets", SprEzeB(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico EZEB, "Pasajeros", SprEzeB(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select


End Sub


Private Sub botGrTortaAEP_Click()
Dim col_datos As Collection

    Set col_datos = New Collection
    
    Select Case TabAep.Tab
        Case 0
            L_LLenarColeccion col_datos, SprAep(0)
            FrmGraficos.CargarGrafico AERO, "Importes", col_datos
        Case 1
            L_LLenarColeccion col_datos, SprAep(1)
            FrmGraficos.CargarGrafico AERO, "Tickets", col_datos
        Case 2
            L_LLenarColeccion col_datos, SprAep(2)
            FrmGraficos.CargarGrafico AERO, "Pasajeros", col_datos

    End Select

End Sub

Private Sub botGrTortaEA_Click()
Dim col_datos As Collection

    Set col_datos = New Collection
    
    Select Case TabEzeA.Tab
        Case 0
            L_LLenarColeccion col_datos, SprEzeA(0)
            FrmGraficos.CargarGrafico EZEA, "Importes", col_datos
        Case 1
            L_LLenarColeccion col_datos, SprEzeA(1)
            FrmGraficos.CargarGrafico EZEA, "Tickets", col_datos
        Case 2
            L_LLenarColeccion col_datos, SprEzeA(2)
            FrmGraficos.CargarGrafico EZEA, "Pasajeros", col_datos

    End Select
End Sub


Private Sub botGrTortaEB_Click()
Dim col_datos As Collection

    Set col_datos = New Collection
    
    Select Case TabEzeB.Tab
        Case 0
            L_LLenarColeccion col_datos, SprEzeB(0)
            FrmGraficos.CargarGrafico EZEB, "Importes", col_datos
        Case 1
            L_LLenarColeccion col_datos, SprEzeB(1)
            FrmGraficos.CargarGrafico EZEB, "Tickets", col_datos
        Case 2
            L_LLenarColeccion col_datos, SprEzeB(2)
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
        L_Porcentages sprEzeAPorc(0), SprEzeA(0), "FILAS"
        L_Porcentages sprEzeAPorc(1), SprEzeA(1), "FILAS"
        L_Porcentages sprEzeAPorc(2), SprEzeA(2), "FILAS"
        botPorc(0).Enabled = True
        frEzeAPorc0.Refresh
        frEzeAPorc1.Refresh
        frEzeAPorc2.Refresh
        L_SeteoEjecutar True, EZEA
    Case 2
        frEzeA0.Visible = False
        frEzeA1.Visible = False
        frEzeA2.Visible = False
        L_Porcentages sprEzeAPorc(0), SprEzeA(0), "COL"
        L_Porcentages sprEzeAPorc(1), SprEzeA(1), "COL"
        L_Porcentages sprEzeAPorc(2), SprEzeA(2), "COL"
        
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
        L_Porcentages sprAepPorc(0), SprAep(0), "FILAS"
        L_Porcentages sprAepPorc(1), SprAep(1), "FILAS"
        L_Porcentages sprAepPorc(2), SprAep(2), "FILAS"
        botPorcAep(0).Enabled = True
        frAepPorc0.Refresh
        frAepPorc1.Refresh
        frAepPorc2.Refresh
        L_SeteoEjecutar True, AERO
    Case 2
        frAep0.Visible = False
        frAep1.Visible = False
        frAep2.Visible = False
        L_Porcentages sprAepPorc(0), SprAep(0), "COL"
        L_Porcentages sprAepPorc(1), SprAep(1), "COL"
        L_Porcentages sprAepPorc(2), SprAep(2), "COL"
        
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
        L_Porcentages sprEzeBPorc(0), SprEzeB(0), "FILAS"
        L_Porcentages sprEzeBPorc(1), SprEzeB(1), "FILAS"
        L_Porcentages sprEzeBPorc(2), SprEzeB(2), "FILAS"
        botPorcB(0).Enabled = True
        frEzeBPorc0.Refresh
        frEzeBPorc1.Refresh
        frEzeBPorc2.Refresh
        L_SeteoEjecutar True, EZEB
    Case 2
        frEzeB0.Visible = False
        frEzeB1.Visible = False
        frEzeB2.Visible = False
        L_Porcentages sprEzeBPorc(0), SprEzeB(0), "COL"
        L_Porcentages sprEzeBPorc(1), SprEzeB(1), "COL"
        L_Porcentages sprEzeBPorc(2), SprEzeB(2), "COL"
        
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
Dim total As Variant
Dim TOTImp As Long
Dim totCant As Long

For i = 0 To 4
    SprEzeA(i).MaxRows = 0
Next

Do While Not RsDataEzeA.EOF
    fecha = Format$(RsDataEzeA!fch_ticket, FTOFECHA)
    
    SprEzeA(0).MaxRows = SprEzeA(0).MaxRows + 1
    SprEzeA(0).SetText 1, SprEzeA(0).MaxRows, Format$(fecha, "dd-mm-yy")
    
    If optComi(3).Value Then
        SprEzeA(1).MaxRows = SprEzeA(1).MaxRows + 1
        SprEzeA(3).MaxRows = SprEzeA(3).MaxRows + 1
    
        SprEzeA(1).SetText 1, SprEzeA(1).MaxRows, Format$(fecha, "dd-mm-yy")
        SprEzeA(3).SetText 1, SprEzeA(3).MaxRows, Format$(fecha, "dd-mm-yy")
    End If
    
    TOTImp = 0
    totCant = 0
    
    Do While fecha = Format$(RsDataEzeA!fch_ticket, FTOFECHA)
        Select Case RsDataEzeA.grupo_venta
            Case "A"
                SprEzeA(0).SetText 2, SprEzeA(0).MaxRows, str(RsDataEzeA!imp)
                If optComi(3).Value Then
                    SprEzeA(1).SetText 2, SprEzeA(1).MaxRows, str(RsDataEzeA!cant)
                    If RsDataEzeA!cant > 0 Then
                    SprEzeA(3).SetText 2, SprEzeA(3).MaxRows, str(RsDataEzeA!imp / RsDataEzeA!cant)
                    End If
                End If
                
            Case "B"
                SprEzeA(0).SetText 3, SprEzeA(0).MaxRows, str(RsDataEzeA!imp)
                If optComi(3).Value Then
                    SprEzeA(1).SetText 3, SprEzeA(1).MaxRows, str(RsDataEzeA!cant)
                    If RsDataEzeA!cant > 0 Then
                    SprEzeA(3).SetText 3, SprEzeA(3).MaxRows, str(RsDataEzeA!imp / RsDataEzeA!cant)
                    End If
                End If
                
            Case "C"
                SprEzeA(0).SetText 4, SprEzeA(0).MaxRows, str(RsDataEzeA!imp)
                If optComi(3).Value Then
                    SprEzeA(1).SetText 4, SprEzeA(1).MaxRows, str(RsDataEzeA!cant)
                    If RsDataEzeA!cant > 0 Then
                    SprEzeA(3).SetText 4, SprEzeA(3).MaxRows, str(RsDataEzeA!imp / RsDataEzeA!cant)
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
        SprEzeA(3).SetText 5, SprEzeA(3).MaxRows, Format$((TOTImp / totCant), "#.0")
    End If
Loop
For i = 0 To 1
    Spread_TotalesGrillas SprEzeA(i), SprEzeA(i).MaxCols - 1, 2
Next

End Sub

Private Sub L_LlenarGrillaINTAPax()
Dim fecha As String
Dim i As Integer
Dim total As Variant
Dim TOTImp As Long
Dim totCant As Long


Do While Not RsDataEzeAPax.EOF
    fecha = Format$(RsDataEzeAPax!fch_vta, FTOFECHA)
    
    SprEzeA(2).MaxRows = SprEzeA(2).MaxRows + 1
    SprEzeA(4).MaxRows = SprEzeA(4).MaxRows + 1
    
    SprEzeA(2).SetText 1, SprEzeA(2).MaxRows, Format$(fecha, "dd-mm-yy")
    SprEzeA(4).SetText 1, SprEzeA(4).MaxRows, Format$(fecha, "dd-mm-yy")
    
    TOTImp = 0
    totCant = 0
    
    Do While fecha = Format$(RsDataEzeAPax!fch_vta, FTOFECHA)
        Select Case RsDataEzeAPax.grupo_venta
            Case "A"
                SprEzeA(2).SetText 2, SprEzeA(2).MaxRows, str(RsDataEzeAPax!cant)
                If RsDataEzeAPax!cant > 0 Then
                SprEzeA(4).SetText 2, SprEzeA(4).MaxRows, str(RsDataEzeAPax!imp / RsDataEzeAPax!cant)
                End If
            Case "B"
                SprEzeA(2).SetText 3, SprEzeA(2).MaxRows, str(RsDataEzeAPax!cant)
                If RsDataEzeAPax!cant > 0 Then
                SprEzeA(4).SetText 3, SprEzeA(4).MaxRows, str(RsDataEzeAPax!imp / RsDataEzeAPax!cant)
                End If
            Case "C"
                SprEzeA(2).SetText 4, SprEzeA(2).MaxRows, str(RsDataEzeAPax!cant)
                If RsDataEzeAPax!cant > 0 Then
                SprEzeA(4).SetText 4, SprEzeA(4).MaxRows, str(RsDataEzeAPax!imp / RsDataEzeAPax!cant)
                End If
        End Select
        TOTImp = TOTImp + RsDataEzeAPax!imp
        totCant = totCant + RsDataEzeAPax!cant
        
        RsDataEzeAPax.MoveNext
        If RsDataEzeAPax.EOF Then
            Exit Do
        End If
    Loop
    SprEzeA(4).SetText 5, SprEzeA(4).MaxRows, Format$((TOTImp / totCant), "#.0")
Loop
    Spread_TotalesGrillas SprEzeA(2), SprEzeA(2).MaxCols - 1, 2
End Sub


Private Sub L_LlenarGrillaINTB()
Dim depn As String, Sdep As String, fecha As String
Dim i As Integer
Dim total As Variant
Dim TOTImp As Long
Dim totCant As Long


For i = 0 To 4
    SprEzeB(i).MaxRows = 0
Next

Do While Not RsDataEzeB.EOF
    fecha = Format$(RsDataEzeB!fch_ticket, FTOFECHA)
    
    SprEzeB(0).MaxRows = SprEzeB(0).MaxRows + 1
    SprEzeB(0).SetText 1, SprEzeB(0).MaxRows, Format$(fecha, "dd-mm-yy")
    
    If optComi(3).Value Then
        SprEzeB(1).MaxRows = SprEzeB(1).MaxRows + 1
        SprEzeB(3).MaxRows = SprEzeB(3).MaxRows + 1
            
        SprEzeB(1).SetText 1, SprEzeB(1).MaxRows, Format$(fecha, "dd-mm-yy")
        SprEzeB(3).SetText 1, SprEzeB(3).MaxRows, Format$(fecha, "dd-mm-yy")
    End If
    TOTImp = 0
    totCant = 0

    Do While fecha = Format$(RsDataEzeB!fch_ticket, FTOFECHA)
        Select Case RsDataEzeB.grupo_venta
            Case "A"
                SprEzeB(0).SetText 2, SprEzeB(0).MaxRows, str(RsDataEzeB!imp)
                If optComi(3).Value Then
                    SprEzeB(1).SetText 2, SprEzeB(1).MaxRows, str(RsDataEzeB!cant)
                    If RsDataEzeB!cant > 0 Then
                    SprEzeB(3).SetText 2, SprEzeB(3).MaxRows, str(RsDataEzeB!imp / RsDataEzeB!cant)
                    End If
                End If
            Case "B"
                SprEzeB(0).SetText 3, SprEzeB(0).MaxRows, str(RsDataEzeB!imp)
                If optComi(3).Value Then
                    SprEzeB(1).SetText 3, SprEzeB(1).MaxRows, str(RsDataEzeB!cant)
                    If RsDataEzeB!cant > 0 Then
                    SprEzeB(3).SetText 3, SprEzeB(3).MaxRows, str(RsDataEzeB!imp / RsDataEzeB!cant)
                    End If
                End If
            Case "C"
                SprEzeB(0).SetText 4, SprEzeB(0).MaxRows, str(RsDataEzeB!imp)
                If optComi(3).Value Then
                    SprEzeB(1).SetText 4, SprEzeB(1).MaxRows, str(RsDataEzeB!cant)
                    If RsDataEzeB!cant > 0 Then
                    SprEzeB(3).SetText 4, SprEzeB(3).MaxRows, str(RsDataEzeB!imp / RsDataEzeB!cant)
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
        SprEzeB(3).SetText 5, SprEzeB(3).MaxRows, Format$((TOTImp / totCant), "#.0")
    End If
Loop
For i = 0 To 1
    Spread_TotalesGrillas SprEzeB(i), SprEzeB(i).MaxCols - 1, 2
Next


End Sub
Private Sub L_LlenarGrillaINTBPax()
Dim fecha As String
Dim i As Integer
Dim total As Variant
Dim TOTImp As Long
Dim totCant As Long


Do While Not RsDataEzeBPax.EOF
    fecha = Format$(RsDataEzeBPax!fch_vta, FTOFECHA)
    
    SprEzeB(2).MaxRows = SprEzeB(2).MaxRows + 1
    SprEzeB(4).MaxRows = SprEzeB(4).MaxRows + 1
    
    SprEzeB(2).SetText 1, SprEzeB(2).MaxRows, Format$(fecha, "dd-mm-yy")
    SprEzeB(4).SetText 1, SprEzeB(4).MaxRows, Format$(fecha, "dd-mm-yy")
    
    TOTImp = 0
    totCant = 0
    
    Do While fecha = Format$(RsDataEzeBPax!fch_vta, FTOFECHA)
        Select Case RsDataEzeBPax.grupo_venta
            Case "A"
                SprEzeB(2).SetText 2, SprEzeB(2).MaxRows, str(RsDataEzeBPax!cant)
                If RsDataEzeBPax!cant > 0 Then
                SprEzeB(4).SetText 2, SprEzeB(4).MaxRows, str(RsDataEzeBPax!imp / RsDataEzeBPax!cant)
                End If
                
            Case "B"
                SprEzeB(2).SetText 3, SprEzeB(2).MaxRows, str(RsDataEzeBPax!cant)
                If RsDataEzeBPax!cant > 0 Then
                SprEzeB(4).SetText 3, SprEzeB(4).MaxRows, str(RsDataEzeBPax!imp / RsDataEzeBPax!cant)
                End If
                
            Case "C"
                SprEzeB(2).SetText 4, SprEzeB(2).MaxRows, str(RsDataEzeBPax!cant)
                If RsDataEzeBPax!cant > 0 Then
                SprEzeB(4).SetText 4, SprEzeB(4).MaxRows, str(RsDataEzeBPax!imp / RsDataEzeBPax!cant)
                End If
        End Select
        TOTImp = TOTImp + RsDataEzeBPax!imp
        totCant = totCant + RsDataEzeBPax!cant
        
        RsDataEzeBPax.MoveNext
        If RsDataEzeBPax.EOF Then
            Exit Do
        End If
    Loop
    SprEzeB(4).SetText 5, SprEzeB(4).MaxRows, Format$((TOTImp / totCant), "#.0")
Loop
    Spread_TotalesGrillas SprEzeB(2), SprEzeB(2).MaxCols - 1, 2
End Sub


Private Sub L_LlenarGrillaAEP()
Dim depn As String, Sdep As String, fecha As String
Dim i As Integer
Dim total As Variant
Dim TOTImp As Long
Dim totCant As Long


For i = 0 To 4
    SprAep(i).MaxRows = 0
Next

Do While Not RsDataAep.EOF
    fecha = Format$(RsDataAep!fch_ticket, FTOFECHA)
    
    SprAep(0).MaxRows = SprAep(0).MaxRows + 1
    SprAep(0).SetText 1, SprAep(0).MaxRows, Format$(fecha, "dd-mm-yy")
    
    If optComi(3).Value Then
        SprAep(1).MaxRows = SprAep(1).MaxRows + 1
        SprAep(3).MaxRows = SprAep(3).MaxRows + 1
        
        SprAep(1).SetText 1, SprAep(1).MaxRows, Format$(fecha, "dd-mm-yy")
        SprAep(3).SetText 1, SprAep(3).MaxRows, Format$(fecha, "dd-mm-yy")
    End If
    
    TOTImp = 0
    totCant = 0
    Do While fecha = Format$(RsDataAep!fch_ticket, FTOFECHA)
        Select Case RsDataAep.grupo_venta
            Case "A"
                SprAep(0).SetText 2, SprAep(0).MaxRows, str(RsDataAep!imp)
                If optComi(3).Value Then
                    SprAep(1).SetText 2, SprAep(1).MaxRows, str(RsDataAep!cant)
                    If RsDataAep!cant > 0 Then
                    SprAep(3).SetText 2, SprAep(3).MaxRows, str(RsDataAep!imp / RsDataAep!cant)
                    End If
                End If
            Case "B"
                SprAep(0).SetText 3, SprAep(0).MaxRows, str(RsDataAep!imp)
                If optComi(3).Value Then
                    SprAep(1).SetText 3, SprAep(1).MaxRows, str(RsDataAep!cant)
                    If RsDataAep!cant > 0 Then
                    SprAep(3).SetText 3, SprAep(3).MaxRows, str(RsDataAep!imp / RsDataAep!cant)
                    End If
                End If
            Case "C"
                SprAep(0).SetText 4, SprAep(0).MaxRows, str(RsDataAep!imp)
                If optComi(3).Value Then
                    SprAep(1).SetText 4, SprAep(1).MaxRows, str(RsDataAep!cant)
                    If RsDataAep!cant > 0 Then
                    SprAep(3).SetText 4, SprAep(3).MaxRows, str(RsDataAep!imp / RsDataAep!cant)
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
        SprAep(3).SetText 5, SprAep(3).MaxRows, Format$((TOTImp / totCant), "#.0")
    End If
Loop
For i = 0 To 1
    Spread_TotalesGrillas SprAep(i), SprAep(i).MaxCols - 1, 2
Next

End Sub
Private Sub L_LlenarGrillaAEPPx()
Dim fecha As String
Dim i As Integer
Dim total As Variant
Dim TOTImp As Long
Dim totCant As Long


Do While Not RsDataAepPax.EOF
    fecha = Format$(RsDataAepPax!fch_vta, FTOFECHA)
    
    SprAep(2).MaxRows = SprAep(2).MaxRows + 1
    SprAep(4).MaxRows = SprAep(4).MaxRows + 1
    
    SprAep(2).SetText 1, SprAep(2).MaxRows, Format$(fecha, "dd-mm-yy")
    SprAep(4).SetText 1, SprAep(4).MaxRows, Format$(fecha, "dd-mm-yy")
    
    TOTImp = 0
    totCant = 0
    
    Do While fecha = Format$(RsDataAepPax!fch_vta, FTOFECHA)
        Select Case RsDataAepPax.grupo_venta
            Case "A"
                SprAep(2).SetText 2, SprAep(2).MaxRows, str(RsDataAepPax!cant)
                If RsDataAepPax!cant > 0 Then
                SprAep(4).SetText 2, SprAep(4).MaxRows, str(RsDataAepPax!imp / RsDataAepPax!cant)
                End If
            Case "B"
                SprAep(2).SetText 3, SprAep(2).MaxRows, str(RsDataAepPax!cant)
                If RsDataAepPax!cant > 0 Then
                SprAep(4).SetText 3, SprAep(4).MaxRows, str(RsDataAepPax!imp / RsDataAepPax!cant)
                End If
            Case "C"
                SprAep(2).SetText 4, SprAep(2).MaxRows, str(RsDataAepPax!cant)
                If RsDataAepPax!cant > 0 Then
                SprAep(4).SetText 4, SprAep(4).MaxRows, str(RsDataAepPax!imp / RsDataAepPax!cant)
                End If
        End Select
        TOTImp = TOTImp + RsDataAepPax!imp
        totCant = totCant + RsDataAepPax!cant
        
        RsDataAepPax.MoveNext
        If RsDataAepPax.EOF Then
            Exit Do
        End If
    Loop
    SprAep(4).SetText 5, SprAep(4).MaxRows, Format$((TOTImp / totCant), "#.0")
Loop
    Spread_TotalesGrillas SprAep(2), SprAep(2).MaxCols - 1, 2
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

Private Sub L_LLenarColeccion(ByRef col As Collection, spr As control)
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
    
    col.Add cl_dato
    
Next

End Sub

