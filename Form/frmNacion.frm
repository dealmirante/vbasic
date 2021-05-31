VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.0#0"; "crystl32.ocx"
Begin VB.Form frmNacion 
   Caption         =   "PASAJEROS VIAJADOS POR NACIONALIDAD"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9195
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   8940
      Top             =   1965
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   3810
      Left            =   75
      TabIndex        =   2
      Top             =   1560
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   6720
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
      TabCaption(0)   =   "RESUMEN"
      TabPicture(0)   =   "frmNacion.frx":0000
      Tab(0).ControlCount=   4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "botExcel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TabTotal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "rptform"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "BotCristal"
      Tab(0).Control(3).Enabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmNacion.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TabEspA(0)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmNacion.frx":0038
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TabEspB"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmNacion.frx":0054
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TabAep"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "INTERIOR"
      TabPicture(4)   =   "frmNacion.frx":0070
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TabInt"
      Tab(4).Control(0).Enabled=   0   'False
      Begin VB.CommandButton BotCristal 
         Height          =   525
         Left            =   8070
         Picture         =   "frmNacion.frx":008C
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2430
         Width           =   720
      End
      Begin Crystal.CrystalReport rptform 
         Left            =   330
         Top             =   2850
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   327682
         PrintFileLinesPerPage=   60
      End
      Begin TabDlg.SSTab TabTotal 
         Height          =   2505
         Left            =   850
         TabIndex        =   23
         Top             =   550
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4419
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmNacion.frx":11C2
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprTot(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Argentinos"
         TabPicture(1)   =   "frmNacion.frx":11DE
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprTot(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Extranjeros"
         TabPicture(2)   =   "frmNacion.frx":11FA
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprTot(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprTot 
            Height          =   1575
            Index           =   2
            Left            =   -74595
            OleObjectBlob   =   "frmNacion.frx":1216
            TabIndex        =   24
            Top             =   255
            Width           =   6420
         End
         Begin FPSpread.vaSpread SprTot 
            Height          =   1575
            Index           =   1
            Left            =   -74595
            OleObjectBlob   =   "frmNacion.frx":17FB
            TabIndex        =   25
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprTot 
            Height          =   1575
            Index           =   0
            Left            =   405
            OleObjectBlob   =   "frmNacion.frx":1DE0
            TabIndex        =   26
            Top             =   255
            Width           =   6360
         End
      End
      Begin VB.CommandButton botExcel 
         Caption         =   "Excel"
         Height          =   510
         Left            =   8085
         Picture         =   "frmNacion.frx":23C5
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3045
         Width           =   705
      End
      Begin TabDlg.SSTab TabEspA 
         Height          =   2505
         Index           =   0
         Left            =   -74150
         TabIndex        =   27
         Top             =   550
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4419
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmNacion.frx":2957
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprEzeA(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Argentinos"
         TabPicture(1)   =   "frmNacion.frx":2973
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprEzeA(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Extranjeros"
         TabPicture(2)   =   "frmNacion.frx":298F
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprEzeA(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprEzeA 
            Height          =   1575
            Index           =   0
            Left            =   405
            OleObjectBlob   =   "frmNacion.frx":29AB
            TabIndex        =   28
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprTot 
            Height          =   1575
            Index           =   4
            Left            =   -74625
            OleObjectBlob   =   "frmNacion.frx":2F90
            TabIndex        =   29
            Top             =   225
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprTot 
            Height          =   1575
            Index           =   5
            Left            =   -74685
            OleObjectBlob   =   "frmNacion.frx":34B8
            TabIndex        =   30
            Top             =   285
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeA 
            Height          =   1575
            Index           =   1
            Left            =   -74595
            OleObjectBlob   =   "frmNacion.frx":39E0
            TabIndex        =   31
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeA 
            Height          =   1575
            Index           =   2
            Left            =   -74595
            OleObjectBlob   =   "frmNacion.frx":3FC5
            TabIndex        =   32
            Top             =   255
            Width           =   6360
         End
      End
      Begin TabDlg.SSTab TabEspB 
         Height          =   2505
         Left            =   -74150
         TabIndex        =   33
         Top             =   550
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4419
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmNacion.frx":45AA
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprEzeB(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Argentinos"
         TabPicture(1)   =   "frmNacion.frx":45C6
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprEzeB(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Extranjeros"
         TabPicture(2)   =   "frmNacion.frx":45E2
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprEzeB(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprEzeA 
            Height          =   1575
            Index           =   3
            Left            =   -74655
            OleObjectBlob   =   "frmNacion.frx":45FE
            TabIndex        =   34
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeA 
            Height          =   1575
            Index           =   4
            Left            =   -74565
            OleObjectBlob   =   "frmNacion.frx":4B26
            TabIndex        =   35
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprTot 
            Height          =   1575
            Index           =   3
            Left            =   -74685
            OleObjectBlob   =   "frmNacion.frx":504E
            TabIndex        =   36
            Top             =   285
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprTot 
            Height          =   1575
            Index           =   6
            Left            =   -74625
            OleObjectBlob   =   "frmNacion.frx":5576
            TabIndex        =   37
            Top             =   225
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeB 
            Height          =   1575
            Index           =   0
            Left            =   405
            OleObjectBlob   =   "frmNacion.frx":5A9E
            TabIndex        =   38
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeB 
            Height          =   1575
            Index           =   1
            Left            =   -74595
            OleObjectBlob   =   "frmNacion.frx":6091
            TabIndex        =   39
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeB 
            Height          =   1575
            Index           =   2
            Left            =   -74595
            OleObjectBlob   =   "frmNacion.frx":6676
            TabIndex        =   40
            Top             =   255
            Width           =   6360
         End
      End
      Begin TabDlg.SSTab TabAep 
         Height          =   2505
         Left            =   -74150
         TabIndex        =   41
         Top             =   550
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4419
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmNacion.frx":6C5B
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprAep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Argentinos"
         TabPicture(1)   =   "frmNacion.frx":6C77
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprAep(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Extranjeros"
         TabPicture(2)   =   "frmNacion.frx":6C93
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprAep(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprEzeB 
            Height          =   1575
            Index           =   3
            Left            =   -74625
            OleObjectBlob   =   "frmNacion.frx":6CAF
            TabIndex        =   42
            Top             =   240
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeB 
            Height          =   1575
            Index           =   4
            Left            =   -74640
            OleObjectBlob   =   "frmNacion.frx":71D7
            TabIndex        =   43
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprAep 
            Height          =   1575
            Index           =   0
            Left            =   420
            OleObjectBlob   =   "frmNacion.frx":76FF
            TabIndex        =   44
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprTot 
            Height          =   1575
            Index           =   7
            Left            =   -74625
            OleObjectBlob   =   "frmNacion.frx":7CF2
            TabIndex        =   45
            Top             =   225
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprTot 
            Height          =   1575
            Index           =   8
            Left            =   -74685
            OleObjectBlob   =   "frmNacion.frx":821A
            TabIndex        =   46
            Top             =   285
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeA 
            Height          =   1575
            Index           =   5
            Left            =   -74565
            OleObjectBlob   =   "frmNacion.frx":8742
            TabIndex        =   47
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeA 
            Height          =   1575
            Index           =   6
            Left            =   -74655
            OleObjectBlob   =   "frmNacion.frx":8C6A
            TabIndex        =   48
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprAep 
            Height          =   1575
            Index           =   1
            Left            =   -74595
            OleObjectBlob   =   "frmNacion.frx":9192
            TabIndex        =   49
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprAep 
            Height          =   1575
            Index           =   2
            Left            =   -74595
            OleObjectBlob   =   "frmNacion.frx":9785
            TabIndex        =   50
            Top             =   255
            Width           =   6360
         End
      End
      Begin TabDlg.SSTab TabInt 
         Height          =   2505
         Left            =   -74150
         TabIndex        =   51
         Top             =   555
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4419
         _Version        =   327680
         TabOrientation  =   1
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmNacion.frx":9D78
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SprInt(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Argentinos"
         TabPicture(1)   =   "frmNacion.frx":9D94
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SprInt(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Extranjeros"
         TabPicture(2)   =   "frmNacion.frx":9DB0
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprInt(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread SprAep 
            Height          =   1575
            Index           =   3
            Left            =   -74640
            OleObjectBlob   =   "frmNacion.frx":9DCC
            TabIndex        =   52
            Top             =   240
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprAep 
            Height          =   1575
            Index           =   4
            Left            =   -74655
            OleObjectBlob   =   "frmNacion.frx":A2F4
            TabIndex        =   53
            Top             =   240
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeA 
            Height          =   1575
            Index           =   7
            Left            =   -74655
            OleObjectBlob   =   "frmNacion.frx":A81C
            TabIndex        =   54
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeA 
            Height          =   1575
            Index           =   8
            Left            =   -74565
            OleObjectBlob   =   "frmNacion.frx":AD44
            TabIndex        =   55
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprTot 
            Height          =   1575
            Index           =   9
            Left            =   -74685
            OleObjectBlob   =   "frmNacion.frx":B26C
            TabIndex        =   56
            Top             =   285
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprTot 
            Height          =   1575
            Index           =   10
            Left            =   -74625
            OleObjectBlob   =   "frmNacion.frx":B794
            TabIndex        =   57
            Top             =   225
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprInt 
            Height          =   1575
            Index           =   0
            Left            =   405
            OleObjectBlob   =   "frmNacion.frx":BCBC
            TabIndex        =   58
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeB 
            Height          =   1575
            Index           =   5
            Left            =   -74640
            OleObjectBlob   =   "frmNacion.frx":C2AF
            TabIndex        =   59
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprEzeB 
            Height          =   1575
            Index           =   6
            Left            =   -74625
            OleObjectBlob   =   "frmNacion.frx":C7D7
            TabIndex        =   60
            Top             =   240
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprInt 
            Height          =   1575
            Index           =   1
            Left            =   -74595
            OleObjectBlob   =   "frmNacion.frx":CCFF
            TabIndex        =   61
            Top             =   255
            Width           =   6360
         End
         Begin FPSpread.vaSpread SprInt 
            Height          =   1575
            Index           =   2
            Left            =   -74595
            OleObjectBlob   =   "frmNacion.frx":D2F2
            TabIndex        =   62
            Top             =   255
            Width           =   6360
         End
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
         Left            =   150
         Picture         =   "frmNacion.frx":D8E5
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Height          =   405
         Index           =   1
         Left            =   135
         Picture         =   "frmNacion.frx":D9E7
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
         Picture         =   "frmNacion.frx":DAE9
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
            Left            =   4740
            TabIndex        =   15
            Top             =   210
            Width           =   3015
            Begin VB.CommandButton botHelpFDAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmNacion.frx":E30B
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmNacion.frx":E47D
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   600
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesdeAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   18
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
               TabIndex        =   19
               Top             =   585
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
               TabIndex        =   21
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
               TabIndex        =   20
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
            Left            =   270
            TabIndex        =   4
            Top             =   225
            Width           =   3015
            Begin VB.CommandButton botHelpFH 
               Height          =   345
               Left            =   2535
               Picture         =   "frmNacion.frx":E5EF
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2550
               Picture         =   "frmNacion.frx":E761
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
      Height          =   420
      Left            =   150
      TabIndex        =   22
      Top             =   5400
      Width           =   8730
   End
End
Attribute VB_Name = "frmNacion"
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
Dim UsoPorInte As Boolean
Dim UsoPorTotal As Boolean

'Para uso del Cristal Report

Dim lFecha As String
Dim lFechaH As String

Dim lCia As Integer
Dim lSdep As String
Dim lTipo As String
Dim lMonto As Double

Dim Info As String

Public Sub L_AplicaFormula()
Dim x As Integer

For x = 0 To 2
    sprEzeA(x).Col = 3
    sprEzeA(x).Row = 3
    sprEzeA(x).Formula = "C2-C1"
    

    sprEzeB(x).Col = 3
    sprEzeB(x).Row = 3
    sprEzeB(x).Formula = "C2-C1"
    
    sprAEP(x).Col = 3
    sprAEP(x).Row = 3
    sprAEP(x).Formula = "C2-C1"
    
'    SprInt(X).col = 3
'    SprInt(X).Row = 3
'    SprInt(X).Formula = "C2-C1"
    
    SprTot(x).Col = 3
    SprTot(x).Row = 3
    SprTot(x).Formula = "C2-C1"
    
Next x

End Sub


Private Sub MeImprimir()
Dim i%

On Error GoTo ErrPrint:

frmNacion.caption = Aplicacion.SeteoProceso(frmNacion.caption)

rptform.Connect = "ODBC;UID=" & Aplicacion.username & ";PWD=" & Aplicacion.password & ";DATABASE=INTER;"

rptform.ReportFileName = RutaReportes
rptform.ReportFileName = RutaReportes & "PaxNac.rpt"
rptform.Formulas(0) = "flaFDesdeUno = DATE(" & Year(mskFDesde.FormattedText) & "," & Month(mskFDesde.FormattedText) & "," & Day(mskFDesde.FormattedText) & ")"
rptform.Formulas(1) = "flaFHastaUno = DATE(" & Year(mskFHasta.FormattedText) & "," & Month(mskFHasta.FormattedText) & "," & Day(mskFHasta.FormattedText) & ")"
rptform.Formulas(2) = "flaFDesdeDos = DATE(" & Year(mskFDesdeAnt.FormattedText) & "," & Month(mskFDesdeAnt.FormattedText) & "," & Day(mskFDesdeAnt.FormattedText) & ")"
rptform.Formulas(3) = "flaFHastaDos = DATE(" & Year(mskFHastaAnt.FormattedText) & "," & Month(mskFHastaAnt.FormattedText) & "," & Day(mskFHastaAnt.FormattedText) & ")"

rptform.Destination = 1     ' 0 pantalla
                            '1 impresora

rptform.Action = True
DoEvents

frmNacion.caption = Aplicacion.SeteoFin

Exit Sub
ErrPrint:
     frmNacion.caption = Aplicacion.SeteoFin
     Exit Sub


End Sub


Public Sub L_BajarGrillaExel(spr As control, apl As Object, FilaInit As Integer, ColInit As Integer, titCol() As String)
Dim gr_fila As Integer
Dim gr_col As Integer, gr_aux_col As Integer
Dim dato As Variant
Dim rango As String



For gr_fila = 0 To spr.MaxRows
    gr_aux_col = 0
    For gr_col = 0 To spr.MaxCols
        spr.Col = gr_col
        If Not spr.ColHidden Then
            If gr_fila = 0 Then
                apl.Application.Cells(FilaInit + gr_fila, gr_aux_col + 1).Value = titCol(gr_col)
            Else
                spr.GetText gr_col, gr_fila, dato
                Exl_PonerValor apl, FilaInit + gr_fila, gr_aux_col + 1, dato
                If IsNumeric(dato) Then
                If dato < 0 Then
                    rango = Exl_rangos(FilaInit + gr_fila, FilaInit + gr_fila, gr_aux_col + 1, gr_aux_col + 1)
                    Exl_LetraColor apl, rango, Exl_Rojo
                End If
                End If
            End If
            gr_aux_col = gr_aux_col + 1
        Else
            
        End If
    Next
Next
rango = Exl_rangos(FilaInit, FilaInit, ColInit, spr.MaxCols)

Exl_Justificacion apl, rango, Exl_CentroCol, Exl_CentroVert, False

Exl_ColorInt apl, rango, Exl_Gris



rango = Exl_rangos(FilaInit, spr.MaxRows + FilaInit, ColInit, spr.MaxCols)

Exl_Lineas apl, rango, Exl_Linsimple

End Sub

Public Sub Func_CellHeight(spr As control)

Dim i As Integer
Dim x As Integer

For i = 0 To spr.MaxRows
  For x = 0 To 4
    spr.Col = x
    spr.Row = i
    spr.RowHeight(spr.Row) = 14
  Next x
Next

End Sub

Public Sub Func_FormatRange(ColDesde As Integer, ColHasta As Integer, RowDesde As Integer, RowHasta As Integer, spr As control, fto As String)

    'Select a block of cells
    spr.Col = ColDesde
    spr.Row = RowDesde
    spr.Col2 = ColHasta
    spr.Row2 = RowHasta
    spr.BlockMode = True

    'Define cells as type FLOAT
    spr.CellType = SS_CELL_TYPE_FLOAT
    spr.TypeHAlign = SS_CELL_H_ALIGN_RIGHT
    spr.TypeFloatMin = "-" & fto
    spr.TypeFloatMax = fto

    spr.TypeFloatDecimalPlaces = 0
    spr.TypeFloatCurrencyChar = Asc("%")
    spr.TypeFloatSepChar = Asc(",")
    spr.TypeFloatDecimalChar = Asc(".")
    spr.TypeFloatSeparator = True
    'Turn block mode off
    spr.BlockMode = False


End Sub


Public Sub Func_FormatRangePuntuacion(ColDesde As Integer, ColHasta As Integer, RowDesde As Integer, RowHasta As Integer, spr As control)



    'Select a block of cells
    spr.Col = ColDesde
    spr.Row = RowDesde
    spr.Col2 = ColHasta
    spr.Row2 = RowHasta
    spr.BlockMode = True

    'Define cells as type FLOAT
    spr.CellType = SS_CELL_TYPE_INTEGER
    spr.TypeHAlign = SS_CELL_H_ALIGN_RIGHT
    'spr.TypeFloatMin = "-99999999.99"
    'spr.TypeFloatMax = "99999999.99"

    'spr.TypeFloatDecimalPlaces = 2
    'spr.TypeFloatCurrencyChar = Asc("%")
    spr.TypeFloatSepChar = Asc(",")
    'spr.TypeFloatDecimalChar = Asc(".")

    'Turn block mode off
    spr.BlockMode = False


End Sub


Public Sub Func_FormatInteger(ColDesde As Integer, ColHasta As Integer, RowDesde As Integer, RowHasta As Integer, spr As control)

    'Select a block of cells
    spr.Col = ColDesde
    spr.Row = RowDesde
    spr.Col2 = ColHasta
    spr.Row2 = RowHasta
    spr.BlockMode = True

    'Define cells as type FLOAT
    spr.CellType = SS_CELL_TYPE_INTEGER       'SS_CELL_TYPE_FLOAT
    spr.TypeHAlign = SS_CELL_H_ALIGN_RIGHT
    spr.TypeFloatMin = "-99999999.99"
    spr.TypeFloatMax = "99999999.99"
    'spr.TypeFloatMoney = False
    'spr.TypeFloatSeparator = True
    'spr.TypeFloatDecimalPlaces = 2
    spr.TypeFloatCurrencyChar = Asc("%")
    spr.TypeFloatSepChar = Asc(",")
    spr.TypeFloatDecimalChar = Asc(".")

    'Turn block mode off
    spr.BlockMode = False




    'Select a single cell
'    SprEzeA(0).col = 2
'    SprEzeA(0).Row = 3

    'Define cells as type FLOAT
'    SprEzeA(0).CellType = SS_CELL_TYPE_FLOAT
'    SprEzeA(0).TypeHAlign = SS_CELL_H_ALIGN_RIGHT
'    SprEzeA(0).TypeFloatMin = "-99999999.00"
'    SprEzeA(0).TypeFloatMax = "99999999.00"
'    SprEzeA(0).TypeFloatMoney = False
'    SprEzeA(0).TypeFloatSeparator = False
'    SprEzeA(0).TypeFloatDecimalPlaces = 0
'    SprEzeA(0).TypeFloatCurrencyChar = Asc("$")
'    SprEzeA(0).TypeFloatSepChar = Asc(",")
'    SprEzeA(0).TypeFloatDecimalChar = Asc(".")

End Sub

Public Sub L_AltoCelda()
Dim i As Integer

For i = 0 To 2
  Func_CellHeight sprEzeA(i)
  Func_CellHeight sprEzeB(i)
  Func_CellHeight sprAEP(i)
'  Func_CellHeight SprInt(i)
  Func_CellHeight SprTot(i)
Next

End Sub

Private Function L_Armarcondicion(anio As Integer) As String
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant

    If anio = Year(mskFDesde.FormattedText) Then
        fechaDesde = mskFDesde.FormattedText
        fechaHasta = mskFHasta.FormattedText
    Else
        fechaDesde = mskFDesdeAnt.FormattedText
        fechaHasta = mskFHastaAnt.FormattedText
    End If


Cond = " WHERE fecha between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

L_Armarcondicion = Cond

End Function

Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant

Do While Not RsData.EOF
        Select Case RsData!cod_depn
            Case DSLoc(1).Dep
                Select Case RsData!NAC
                  Case "N"
                    sprAEP(1).MaxRows = sprAEP(1).MaxRows + 1
                    sprAEP(1).SetText 0, sprAEP(1).MaxRows, "Período Actual"
                    sprAEP(1).SetText 1, sprAEP(1).MaxRows, str(RsData!imp)
                    sprAEP(1).SetText 2, sprAEP(1).MaxRows, str(RsData!vol)
                    sprAEP(1).SetText 5, sprAEP(1).MaxRows, str(RsData!att)
                   Case "E"
                    sprAEP(2).MaxRows = sprAEP(2).MaxRows + 1
                    sprAEP(2).SetText 0, sprAEP(2).MaxRows, "Período Actual"
                    sprAEP(2).SetText 1, sprAEP(2).MaxRows, str(RsData!imp)
                    sprAEP(2).SetText 2, sprAEP(2).MaxRows, str(RsData!vol)
                    sprAEP(2).SetText 5, sprAEP(2).MaxRows, str(RsData!att)
                End Select
            
            Case DSLoc(2).Dep
                
                Select Case RsData!Cod_Sdep
                    Case DSLoc(2).Sdep
                      Select Case RsData!NAC
                        Case Is = "N"
                            sprEzeA(1).MaxRows = sprEzeA(1).MaxRows + 1
                            sprEzeA(1).SetText 0, sprEzeA(1).MaxRows, "Período Actual"
                            sprEzeA(1).SetText 1, sprEzeA(1).MaxRows, str(RsData!imp)
                            sprEzeA(1).SetText 2, sprEzeA(1).MaxRows, str(RsData!vol)
                            sprEzeA(1).SetText 5, sprEzeA(1).MaxRows, str(RsData!att)
                        Case Is = "E"
                            sprEzeA(2).MaxRows = sprEzeA(2).MaxRows + 1
                            sprEzeA(2).SetText 0, sprEzeA(2).MaxRows, "Período Actual"
                            sprEzeA(2).SetText 1, sprEzeA(2).MaxRows, str(RsData!imp)
                            sprEzeA(2).SetText 2, sprEzeA(2).MaxRows, str(RsData!vol)
                            sprEzeA(2).SetText 5, sprEzeA(2).MaxRows, str(RsData!att)
                
                      End Select
                      
                    Case DSLoc(3).Sdep
                      Select Case RsData!NAC
                        Case Is = "N"
                            sprEzeB(1).MaxRows = sprEzeB(1).MaxRows + 1
                            sprEzeB(1).SetText 0, sprEzeB(1).MaxRows, "Período Actual"
                            sprEzeB(1).SetText 1, sprEzeB(1).MaxRows, str(RsData!imp)
                            sprEzeB(1).SetText 2, sprEzeB(1).MaxRows, str(RsData!vol)
                            sprEzeB(1).SetText 5, sprEzeB(1).MaxRows, str(RsData!att)
                        Case Is = "E"
                            sprEzeB(2).MaxRows = sprEzeB(2).MaxRows + 1
                            sprEzeB(2).SetText 0, sprEzeB(2).MaxRows, "Período Actual"
                            sprEzeB(2).SetText 1, sprEzeB(2).MaxRows, str(RsData!imp)
                            sprEzeB(2).SetText 2, sprEzeB(2).MaxRows, str(RsData!vol)
                            sprEzeB(2).SetText 5, sprEzeB(2).MaxRows, str(RsData!att)
                      
                      End Select
                              
                End Select
            
            Case DSLoc(4).Dep  'Interior
               Select Case RsData!NAC
                  Case "N"
                  '          SprInt(1).MaxRows = SprInt(1).MaxRows + 1
                  '          SprInt(1).SetText 0, SprInt(1).MaxRows, "Período Actual"
                  '          SprInt(1).SetText 1, SprInt(1).MaxRows, str(RsData!imp)
                  '          SprInt(1).SetText 2, SprInt(1).MaxRows, str(RsData!vol)
                  '          SprInt(1).SetText 5, SprInt(1).MaxRows, str(RsData!att)
                        Case Is = "E"
                  '          SprInt(2).MaxRows = SprInt(2).MaxRows + 1
                  '          SprInt(2).SetText 0, SprInt(2).MaxRows, "Período Actual"
                  '          SprInt(2).SetText 1, SprInt(2).MaxRows, str(RsData!imp)
                  '          SprInt(2).SetText 2, SprInt(2).MaxRows, str(RsData!vol)
                  '          SprInt(2).SetText 5, SprInt(2).MaxRows, str(RsData!att)
           
              End Select
                
        End Select
                
        RsData.MoveNext

Loop


End Sub
Private Sub L_DecoEspigonAnt()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant

Do While Not RsDataAnt.EOF
        Select Case RsDataAnt!cod_depn
            Case DSLoc(1).Dep
                Select Case RsDataAnt!NAC
                  Case "N"
                    sprAEP(1).MaxRows = sprAEP(1).MaxRows + 1
                    sprAEP(1).SetText 0, sprAEP(1).MaxRows, "Período Anterior"
                    sprAEP(1).SetText 1, sprAEP(1).MaxRows, str(RsDataAnt!imp)
                    sprAEP(1).SetText 2, sprAEP(1).MaxRows, str(RsDataAnt!vol)
                    sprAEP(1).SetText 5, sprAEP(1).MaxRows, str(RsDataAnt!att)
                 Case "E"
                    sprAEP(2).MaxRows = sprAEP(2).MaxRows + 1
                    sprAEP(2).SetText 0, sprAEP(2).MaxRows, "Período Anterior"
                    sprAEP(2).SetText 1, sprAEP(2).MaxRows, str(RsDataAnt!imp)
                    sprAEP(2).SetText 2, sprAEP(2).MaxRows, str(RsDataAnt!vol)
                    sprAEP(2).SetText 5, sprAEP(2).MaxRows, str(RsDataAnt!att)
                End Select
            Case DSLoc(2).Dep
                
                Select Case RsDataAnt!Cod_Sdep
                    Case DSLoc(2).Sdep
                      Select Case RsDataAnt!NAC
                        Case Is = "N"
                            sprEzeA(1).MaxRows = sprEzeA(1).MaxRows + 1
                            sprEzeA(1).SetText 0, sprEzeA(1).MaxRows, "Período Anterior"
                            sprEzeA(1).SetText 1, sprEzeA(1).MaxRows, str(RsDataAnt!imp)
                            sprEzeA(1).SetText 2, sprEzeA(1).MaxRows, str(RsDataAnt!vol)
                            sprEzeA(1).SetText 5, sprEzeA(1).MaxRows, str(RsDataAnt!att)
                        Case Is = "E"
                            sprEzeA(2).MaxRows = sprEzeA(2).MaxRows + 1
                            sprEzeA(2).SetText 0, sprEzeA(2).MaxRows, "Período Anterior"
                            sprEzeA(2).SetText 1, sprEzeA(2).MaxRows, str(RsDataAnt!imp)
                            sprEzeA(2).SetText 2, sprEzeA(2).MaxRows, str(RsDataAnt!vol)
                            sprEzeA(2).SetText 5, sprEzeA(2).MaxRows, str(RsDataAnt!att)
                      End Select
                      
                    Case DSLoc(3).Sdep
                      Select Case RsDataAnt!NAC
                        Case Is = "N"
                            sprEzeB(1).MaxRows = sprEzeB(1).MaxRows + 1
                            sprEzeB(1).SetText 0, sprEzeB(1).MaxRows, "Período Anterior"
                            sprEzeB(1).SetText 1, sprEzeB(1).MaxRows, str(RsDataAnt!imp)
                            sprEzeB(1).SetText 2, sprEzeB(1).MaxRows, str(RsDataAnt!vol)
                            sprEzeB(1).SetText 5, sprEzeB(1).MaxRows, str(RsDataAnt!att)
                        Case Is = "E"
                            sprEzeB(2).MaxRows = sprEzeB(2).MaxRows + 1
                            sprEzeB(2).SetText 0, sprEzeB(2).MaxRows, "Período Anterior"
                            sprEzeB(2).SetText 1, sprEzeB(2).MaxRows, str(RsDataAnt!imp)
                            sprEzeB(2).SetText 2, sprEzeB(2).MaxRows, str(RsDataAnt!vol)
                            sprEzeB(2).SetText 5, sprEzeB(2).MaxRows, str(RsDataAnt!att)
                            
                      End Select
                              
                End Select
            
            Case DSLoc(4).Dep 'Interior
               Select Case RsDataAnt!NAC
                  Case "N"
                  '          SprInt(1).MaxRows = SprInt(1).MaxRows + 1
                  '          SprInt(1).SetText 0, SprInt(1).MaxRows, "Período Anterior"
                  '          SprInt(1).SetText 1, SprInt(1).MaxRows, str(RsDataAnt!imp)
                  '          SprInt(1).SetText 2, SprInt(1).MaxRows, str(RsDataAnt!vol)
                  '          SprInt(1).SetText 5, SprInt(1).MaxRows, str(RsDataAnt!att)
                        Case Is = "E"
                  '          SprInt(2).MaxRows = SprInt(2).MaxRows + 1
                  '          SprInt(2).SetText 0, SprInt(2).MaxRows, "Período Anterior"
                  '          SprInt(2).SetText 1, SprInt(2).MaxRows, str(RsDataAnt!imp)
                  '          SprInt(2).SetText 2, SprInt(2).MaxRows, str(RsDataAnt!vol)
                  '          SprInt(2).SetText 5, SprInt(2).MaxRows, str(RsDataAnt!att)
                  End Select
                
        End Select
                
        RsDataAnt.MoveNext

Loop

End Sub




Public Sub L_FormatearCelda()

Dim i As Integer

For i = 0 To 2

'Func_FormatRange 1, 4, 4, 4, SprEzeA(i), "99999999.99"
'Func_FormatRange 1, 4, 4, 4, SprEzeB(i), "99999999.99"
'Func_FormatRange 1, 4, 4, 4, SprAep(i), "99999999.99"
'Func_FormatRange 1, 4, 4, 4, SprInt(i), "99999999.99"
'Func_FormatRange 1, 4, 4, 4, SprTot(i), "99999999.99"

'Func_FormatInteger 2, 2, 1, 3, SprEzeA(i)
'Func_FormatInteger 2, 2, 1, 3, SprEzeB(i)
'Func_FormatInteger 2, 2, 1, 3, SprAep(i)
'Func_FormatInteger 2, 2, 1, 3, SprInt(i)
'Func_FormatInteger 2, 2, 1, 3, SprTot(i)

Func_FormatRange 2, 2, 1, 3, sprEzeA(i), "99999999.00"
Func_FormatRange 2, 2, 1, 3, sprEzeB(i), "99999999.00"
Func_FormatRange 2, 2, 1, 3, sprAEP(i), "99999999.00"
Func_FormatRange 2, 2, 1, 3, sprInt(i), "99999999.00"
Func_FormatRange 2, 2, 1, 3, SprTot(i), "99999999.00"

Next

End Sub

Private Sub L_LimpiarGrillas()
Dim i As Integer

For i = 0 To 3
    sprEzeA(i).MaxRows = 0
    sprEzeB(i).MaxRows = 0
    sprAEP(i).MaxRows = 0
    SprTot(i).MaxRows = 0
Next i
    labInfo.caption = ""
    Timer1.Enabled = False
End Sub

Private Sub L_LlenarGenerales()

L_Llenargrilla sprEzeA(1), sprEzeA(2), sprEzeA(0)
L_Llenargrilla sprEzeB(1), sprEzeB(2), sprEzeB(0)
L_Llenargrilla sprAEP(1), sprAEP(2), sprAEP(0)
'L_Llenargrilla SprInt(1), SprInt(2), SprInt(0)
 
End Sub

Private Sub L_LlenarResumen()
Dim i As Integer
For i = 0 To 2
   SprTot(i).MaxRows = 4
Next

L_LlenarGrillaResumen sprEzeA(2), sprEzeB(2), sprAEP(2), sprInt(2), SprTot(2)
L_LlenarGrillaResumen sprEzeA(1), sprEzeB(1), sprAEP(1), sprInt(1), SprTot(1)
L_Llenargrilla SprTot(1), SprTot(2), SprTot(0)

For i = 0 To 2
    L_DiferenciaNominal SprTot(i)
    L_DiferenciaPorcentual SprTot(i)
Next

End Sub
Private Sub L_LlenarDiferencias()

Dim i As Integer

For i = 1 To 2

    L_DiferenciaNominal sprEzeA(i)
    L_DiferenciaNominal sprEzeB(i)
    L_DiferenciaNominal sprAEP(i)
'    L_DiferenciaNominal SprInt(i)
    
    L_DiferenciaPorcentual sprEzeA(i)
    L_DiferenciaPorcentual sprEzeB(i)
    L_DiferenciaPorcentual sprAEP(i)
'    L_DiferenciaPorcentual SprInt(i)
    
Next

End Sub

Private Sub L_LlenarDiferenciasGral()

L_DiferenciaNominal sprEzeA(0)
L_DiferenciaNominal sprEzeB(0)
L_DiferenciaNominal sprAEP(0)
'L_DiferenciaNominal SprInt(0)
 
L_DiferenciaPorcentual sprEzeA(0)
L_DiferenciaPorcentual sprEzeB(0)
L_DiferenciaPorcentual sprAEP(0)
'L_DiferenciaPorcentual SprInt(0)
    


End Sub

Private Sub L_Llenargrilla(Spr1 As control, spr2 As control, sprGral As control)

Dim ValorGral As Variant
Dim valor1 As Variant
Dim Valor2 As Variant
Dim Col As Integer
Dim Row As Integer


Col = 1
Row = 1
valor1 = 0
Valor2 = 0


Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
ValorGral = valor1 + Valor2
sprGral.MaxRows = 2
sprGral.SetText 0, Row, "Período Anterior"
sprGral.SetText Col, Row, str(ValorGral)

Row = 2
Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
ValorGral = valor1 + Valor2
sprGral.SetText 0, Row, "Período Actual"
sprGral.SetText Col, Row, str(ValorGral)

Col = 2
Row = 1
Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
ValorGral = valor1 + Valor2
sprGral.SetText Col, Row, str(ValorGral)

Row = 2
Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
ValorGral = valor1 + Valor2
sprGral.SetText Col, Row, str(ValorGral)


Col = 5
Row = 1
Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
ValorGral = valor1 + Valor2
sprGral.SetText Col, Row, str(ValorGral)

Row = 2
Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
ValorGral = valor1 + Valor2
sprGral.SetText Col, Row, str(ValorGral)

End Sub
Private Sub L_LlenarGrillaResumen(Spr1 As control, spr2 As control, Spr3 As control, spr4 As control, SprResumen As control)

Dim ValorResumen As Variant
Dim valor1 As Variant
Dim Valor2 As Variant
Dim Valor3 As Variant
Dim Valor4 As Variant
Dim Col As Integer
Dim Row As Integer


Col = 1
Row = 1
valor1 = 0
Valor2 = 0

SprResumen.SetText 0, 1, "Período Anterior"

Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
Spr3.GetText Col, Row, Valor3
spr4.GetText Col, Row, Valor4
ValorResumen = valor1 + Valor2 + Valor3 + Valor4
SprResumen.SetText Col, Row, str(ValorResumen)

Row = 2

SprResumen.SetText 0, 2, "Período Actual"

Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
Spr3.GetText Col, Row, Valor3
spr4.GetText Col, Row, Valor4
ValorResumen = valor1 + Valor2 + Valor3 + Valor4
SprResumen.SetText Col, Row, str(ValorResumen)


Col = 2
Row = 1

Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
Spr3.GetText Col, Row, Valor3
spr4.GetText Col, Row, Valor4
ValorResumen = valor1 + Valor2 + Valor3 + Valor4
SprResumen.SetText Col, Row, str(ValorResumen)

Row = 2

Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
Spr3.GetText Col, Row, Valor3
spr4.GetText Col, Row, Valor4
ValorResumen = valor1 + Valor2 + Valor3 + Valor4
SprResumen.SetText Col, Row, str(ValorResumen)


Col = 5
Row = 1

Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
Spr3.GetText Col, Row, Valor3
spr4.GetText Col, Row, Valor4
ValorResumen = valor1 + Valor2 + Valor3 + Valor4
SprResumen.SetText Col, Row, str(ValorResumen)

Row = 2

Spr1.GetText Col, Row, valor1
spr2.GetText Col, Row, Valor2
Spr3.GetText Col, Row, Valor3
spr4.GetText Col, Row, Valor4
ValorResumen = valor1 + Valor2 + Valor3 + Valor4
SprResumen.SetText Col, Row, str(ValorResumen)

End Sub

Private Sub L_DiferenciaNominal(Spr1 As control)
Dim valor1 As Variant
Dim Valor2 As Variant
Dim DifNominal As Variant
Dim i As Integer

For i = 1 To 4

valor1 = 0
Valor2 = 0

    Spr1.GetText i, 1, valor1
    Spr1.GetText i, 2, Valor2
    DifNominal = Valor2 - valor1
    Spr1.MaxRows = 3
    Spr1.SetText i, 3, str(DifNominal)
    
Next

Spr1.SetText 0, 3, "Dif. Nominal"

End Sub

Private Sub L_DiferenciaPorcentual(Spr1 As control)

Dim valor1 As Variant
Dim Valor2 As Variant
Dim DifPorcentual As Variant
Dim i As Integer

For i = 1 To 4

valor1 = 0
Valor2 = 0

    Spr1.GetText i, 3, valor1
    Spr1.GetText i, 1, Valor2
    DifPorcentual = (valor1 * 100) / Valor2
    Spr1.MaxRows = 4
    Spr1.SetText i, 4, str(DifPorcentual)
    
Next

Spr1.SetText 0, 4, "Dif. Porcent. %"

End Sub
Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset

On Error GoTo ErrInd:

frmNacion.caption = Aplicacion.SeteoProceso(frmNacion.caption)


sql = " SELECT "
sql = sql & " cod_depn, "
sql = sql & " cod_sdep, "
sql = sql & " nac, "
sql = sql & " sum(imp)imp, "
sql = sql & " sum(vol) vol ,"
sql = sql & " sum(att) att "
sql = sql & "FROM estadis.VIEW_PAX_RESUMEN "
sql = sql & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)))
sql = sql & "group by cod_depn,cod_sdep,nac"
sql = sql & " order by cod_depn,cod_sdep,nac"


sqlX = " SELECT "
sqlX = sqlX & " cod_depn, "
sqlX = sqlX & " cod_sdep, "
sqlX = sqlX & " nac, "
sqlX = sqlX & " sum(imp)imp, "
sqlX = sqlX & " sum(vol) vol ,"
sqlX = sqlX & " sum(att) att "
sqlX = sqlX & "FROM estadis.VIEW_PAX_RESUMEN "
sqlX = sqlX & L_Armarcondicion(Year(CDate(mskFDesdeAnt.FormattedText)) - 1)
sqlX = sqlX & "group by cod_depn,cod_sdep,nac"
sqlX = sqlX & " order by cod_depn,cod_sdep,nac"



If Aplicacion.ObtenerRsDAO(sqlX, RsDataAnt) Then
    
    If Aplicacion.CantReg(RsDataAnt) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        L_DecoEspigonAnt
        If Aplicacion.ObtenerRsDAO(sql, RsData) Then
            L_DecoEspigon
        End If
        
        L_LlenarDiferencias
        L_LlenarGenerales
        L_LlenarDiferenciasGral
        L_LlenarResumen
        L_FormatearCelda
        L_AltoCelda
        L_AplicaFormula
        L_Resaltar
        
        Aplicacion.CerrarDAO RsData
    End If
    
    Aplicacion.CerrarDAO RsDataAnt

End If


ErrInd:
    frmNacion.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub


Private Sub L_Resaltar()
Dim fila As Integer
Dim i As Integer

For i = 0 To 2

For fila = 3 To SprTot(i).MaxRows
    Spread.spread_ResaltarCelda SprTot(i), 1, fila
    Spread.spread_ResaltarCelda SprTot(i), 2, fila
    Spread.spread_ResaltarCelda SprTot(i), 3, fila
    Spread.spread_ResaltarCelda SprTot(i), 4, fila
    
    Spread.spread_ResaltarCelda sprEzeA(i), 1, fila
    Spread.spread_ResaltarCelda sprEzeA(i), 2, fila
    Spread.spread_ResaltarCelda sprEzeA(i), 3, fila
    Spread.spread_ResaltarCelda sprEzeA(i), 4, fila

    Spread.spread_ResaltarCelda sprEzeB(i), 1, fila
    Spread.spread_ResaltarCelda sprEzeB(i), 2, fila
    Spread.spread_ResaltarCelda sprEzeB(i), 3, fila
    Spread.spread_ResaltarCelda sprEzeB(i), 4, fila

    Spread.spread_ResaltarCelda sprAEP(i), 1, fila
    Spread.spread_ResaltarCelda sprAEP(i), 2, fila
    Spread.spread_ResaltarCelda sprAEP(i), 3, fila
    Spread.spread_ResaltarCelda sprAEP(i), 4, fila


'    Spread.spread_ResaltarCelda SprInt(i), 1, fila
'    Spread.spread_ResaltarCelda SprInt(i), 2, fila
'    Spread.spread_ResaltarCelda SprInt(i), 3, fila
'    Spread.spread_ResaltarCelda SprInt(i), 4, fila

Next fila

Next i

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
        Case INTE
            UsoPorInte = True
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
            UsoPorTotal = False
        Case INTE
            UsoPorInte = False
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntB Or UsoPorAero Or UsoPorInte Or UsoPorTotal) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub







Private Sub BotCristal_Click()
 MeImprimir
End Sub

Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        L_Refrescar
    Case 1

        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        tabEspigon.Enabled = False
        L_LimpiarGrillas
        
    Case 2
        Unload Me
End Select
End Sub



Private Sub botExcel_Click()

L_TratarExcel SprTot(0), "Análisis de Ventas por Pasajeros Viajados por Nacionalidad ", "", "", Exl_Blanco

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
tabEspigon.TabVisible(4) = False


End Sub

Private Sub L_LLenarColeccion(ByRef Col As Collection, spr As control)
Dim cl_dato As CLlgi
Dim valor As Variant
Dim item As Variant

spr.Row = spr.ActiveRow
spr.GetText 1, spr.Row, item

Set cl_dato = New CLlgi
    
spr.GetText 2, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
Col.Add cl_dato
    
Set cl_dato = New CLlgi
spr.GetText 3, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
Col.Add cl_dato

Set cl_dato = New CLlgi
spr.GetText 6, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
Col.Add cl_dato

End Sub


Private Sub L_TratarExcel(spr As control, titulo As String, subTit As String, Info As String, color As Integer)
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

frmNacion.caption = Aplicacion.SeteoProceso(frmNacion.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    AppExcel.Application.Visible = True
    
    ReDim titCol(spr.MaxCols)
    Col = 1
    fila = 3
    
    For i = 0 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 1, 1, titulo
    rango = Exl_rangos(1, 1, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, color
    
    Exl_PonerValor AppExcel, fila, Col, "Período Actual : " & Format$(mskFDesde.FormattedText, FTOFECHA) & " al " & Format$(mskFHasta.FormattedText, FTOFECHA)
    rango = Exl_rangos(fila, fila, 1, 3)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    fila = fila + 1
    
    Exl_PonerValor AppExcel, fila, Col, "Período Anterior : " & Format$(mskFDesdeAnt.FormattedText, FTOFECHA) & " al " & Format$(mskFHastaAnt.FormattedText, FTOFECHA)
    rango = Exl_rangos(fila, fila, 1, 3)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    fila = fila + 4
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia General"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel spr, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + spr.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
    
    '------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia Argentinos"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel SprTot(1), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + SprTot(1).MaxRows, Col + 1, SprTot(1).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + SprTot(1).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, SprTot(1).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - SprTot(1).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
    '-----------------------------------
   
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia Extranjeros"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel SprTot(2), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + SprTot(2).MaxRows, Col + 1, SprTot(2).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + SprTot(2).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, SprTot(2).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - SprTot(2).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A General"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeA(0), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeA(0).MaxRows, Col + 1, sprEzeA(0).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeA(0).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeA(0).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeA(0).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A Argentinos"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeA(1), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeA(1).MaxRows, Col + 1, sprEzeA(1).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeA(1).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeA(1).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeA(1).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
       
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A Argentinos"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeA(2), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeA(2).MaxRows, Col + 1, sprEzeA(2).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeA(2).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeA(2).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeA(2).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
   
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B General"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeB(0), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeB(0).MaxRows, Col + 1, sprEzeB(0).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeB(0).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeB(0).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeB(0).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B Argentinos"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeB(1), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeB(1).MaxRows, Col + 1, sprEzeB(1).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeB(1).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeB(1).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeB(1).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
       
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B Extranjeros"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeB(2), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeB(2).MaxRows, Col + 1, sprEzeB(2).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeB(2).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeB(2).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeB(2).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque General"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprAEP(0), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprAEP(0).MaxRows, Col + 1, sprAEP(0).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprAEP(0).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprAEP(0).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprAEP(0).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque Argentinos"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprAEP(1), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprAEP(1).MaxRows, Col + 1, sprAEP(1).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprAEP(1).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprAEP(1).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprAEP(1).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
       
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque Extranjeros"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprAEP(2), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprAEP(2).MaxRows, Col + 1, sprAEP(2).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprAEP(2).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprAEP(2).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprAEP(2).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
'    fila = fila + 3

'    Exl_PonerValor AppExcel, fila, col, "Total Interior General"
'    rango = Exl_rangos(fila, fila, 1, 2)
'    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
'    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
'    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
'    fila = fila + 2
    
'    L_BajarGrillaExel SprInt(0), AppExcel, fila, col, titCol
'    rango = Exl_rangos(fila + 1, fila + SprInt(0).MaxRows, col + 1, SprInt(0).MaxCols)
'    Exl_Format AppExcel, rango
'    fila = fila + SprInt(0).MaxRows + 1
'    rango = Exl_rangos(fila, fila, col, SprInt(0).MaxCols)
'    Exl_ColorInt AppExcel, rango, color
    
'    rango = Exl_rangos(fila - SprInt(0).MaxRows, fila - 1, 5, 5)
'    Exl_ColorInt AppExcel, rango, Exl_Gris
    
'    Exl_AnchoCol AppExcel, 1, 1, 20
'    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
'    fila = fila + 3
    
'    Exl_PonerValor AppExcel, fila, col, "Total Interior Argentinos"
'    rango = Exl_rangos(fila, fila, 1, 2)
'    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
'    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
'    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
'    fila = fila + 2
    
'    L_BajarGrillaExel SprInt(1), AppExcel, fila, col, titCol
'    rango = Exl_rangos(fila + 1, fila + SprInt(1).MaxRows, col + 1, SprInt(1).MaxCols)
'    Exl_Format AppExcel, rango
'    fila = fila + SprInt(1).MaxRows + 1
'    rango = Exl_rangos(fila, fila, col, SprInt(1).MaxCols)
'    Exl_ColorInt AppExcel, rango, color
    
'    rango = Exl_rangos(fila - SprInt(1).MaxRows, fila - 1, 5, 5)
'    Exl_ColorInt AppExcel, rango, Exl_Gris
    
'    Exl_AnchoCol AppExcel, 1, 1, 20
'    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
       
'    fila = fila + 3
    
'    Exl_PonerValor AppExcel, fila, col, "Total Interior Extranjeros"
'    rango = Exl_rangos(fila, fila, 1, 2)
'    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
'    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
'    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
'    fila = fila + 2
    
'    L_BajarGrillaExel SprInt(2), AppExcel, fila, col, titCol
'    rango = Exl_rangos(fila + 1, fila + SprInt(2).MaxRows, col + 1, SprInt(2).MaxCols)
'    Exl_Format AppExcel, rango
'    fila = fila + SprInt(2).MaxRows + 1
'    rango = Exl_rangos(fila, fila, col, SprInt(2).MaxCols)
'    Exl_ColorInt AppExcel, rango, color
    
'    rango = Exl_rangos(fila - SprInt(2).MaxRows, fila - 1, 5, 5)
'    Exl_ColorInt AppExcel, rango, Exl_Gris
    
'    Exl_AnchoCol AppExcel, 1, 1, 20
'    Exl_AnchoCol AppExcel, 2, 3, 17
   
    '-------------------------------
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True

    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
   AppExcel.SaveAs NOMBRE & ".xls"
   Set AppExcel = Nothing
End If

ErrorExl:

    frmNacion.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6300
Me.Width = 9300

mskFDesde.Text = "01-" & Format$(Month(Date), "0#") & "-" & Format$(Year(Date), "####")
mskFHasta.Text = Format$(Date - 1, FTOFECHA)

mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio, FTOFECHA)

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
    '    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
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
ElseIf CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Or Year(mskFHasta.FormattedText) <> Year(mskFDesde.FormattedText) Then
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

'Private Sub optFechas_Click(Index As Integer)
'Select Case Index
'    Case 0
'        mskFHasta.Visible = False
'        botHelpFH.Visible = False
'        Label1(0).caption = "Fecha Tope"
'        Label1(1).Visible = False
'        frPerAnt.Visible = False
'    Case 1
'        mskFHasta.Visible = True
'        botHelpFH.Visible = True
'        Label1(0).caption = "Fecha Desde"
'        Label1(1).Visible = True
'        frPerAnt.Visible = True
'End Select
'End Sub

Private Sub tabEspigon_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    Select Case tabEspigon.Tab
        Case 0
            SprTot(0).SetFocus
        Case 1
            sprEzeA(0).SetFocus
        Case 2
            sprEzeB(0).SetFocus
        Case 3
            sprAEP(0).SetFocus
    End Select
    
    
ErrT:
    Exit Sub



End Sub

Private Sub Timer1_Timer()
    labInfo.Visible = Not labInfo.Visible
End Sub


