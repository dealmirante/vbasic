VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.0#0"; "crystl32.ocx"
Begin VB.Form frmNacion_ES 
   Caption         =   "PASAJEROS VIAJADOS POR NACIONALIDAD - Entradas y Salidas"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   9195
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   8940
      Top             =   1965
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
         Picture         =   "frmNacion_ES.frx":0000
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
         Picture         =   "frmNacion_ES.frx":0102
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
         Picture         =   "frmNacion_ES.frx":0204
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
            TabIndex        =   14
            Top             =   210
            Width           =   3015
            Begin VB.CommandButton botHelpFDAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmNacion_ES.frx":0A26
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmNacion_ES.frx":0B98
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
            Left            =   270
            TabIndex        =   4
            Top             =   225
            Width           =   3015
            Begin VB.CommandButton botHelpFH 
               Height          =   345
               Left            =   2535
               Picture         =   "frmNacion_ES.frx":0D0A
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2550
               Picture         =   "frmNacion_ES.frx":0E7C
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
   Begin TabDlg.SSTab tabEspigon 
      Height          =   4395
      Left            =   90
      TabIndex        =   2
      Top             =   1695
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   7752
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
      TabPicture(0)   =   "frmNacion_ES.frx":0FEE
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rptform"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TabResumen_ES"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmNacion_ES.frx":100A
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabEZEA"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmNacion_ES.frx":1026
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TabEZEB"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmNacion_ES.frx":1042
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TabAEP_ES"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "INTERIOR"
      TabPicture(4)   =   "frmNacion_ES.frx":105E
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TabINT_ES"
      Tab(4).Control(0).Enabled=   0   'False
      Begin TabDlg.SSTab TabINT_ES 
         Height          =   3300
         Left            =   -74600
         TabIndex        =   60
         Top             =   600
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5821
         _Version        =   327680
         TabOrientation  =   3
         TabHeight       =   520
         ForeColor       =   9926553
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "LLEGADA"
         TabPicture(0)   =   "frmNacion_ES.frx":107A
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "TabINT(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "SALIDA"
         TabPicture(1)   =   "frmNacion_ES.frx":1096
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TabINT(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "TOTAL"
         TabPicture(2)   =   "frmNacion_ES.frx":10B2
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "TabINT(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin TabDlg.SSTab TabINT 
            Height          =   2500
            Index           =   0
            Left            =   500
            TabIndex        =   61
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":10CE
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprInt(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":10EA
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprInt(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":1106
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprInt(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprInt 
               Height          =   1500
               Index           =   1
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":1122
               TabIndex        =   64
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprInt 
               Height          =   1500
               Index           =   2
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":1715
               TabIndex        =   63
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprInt 
               Height          =   1500
               Index           =   0
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":1D08
               TabIndex        =   62
               Top             =   300
               Width           =   6400
            End
         End
         Begin TabDlg.SSTab TabINT 
            Height          =   2500
            Index           =   1
            Left            =   -74500
            TabIndex        =   65
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":22FB
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprInt(3)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":2317
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprInt(4)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":2333
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprInt(5)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprInt 
               Height          =   1500
               Index           =   5
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":234F
               TabIndex        =   68
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprInt 
               Height          =   1500
               Index           =   4
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":2942
               TabIndex        =   67
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprInt 
               Height          =   1500
               Index           =   3
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":2F35
               TabIndex        =   66
               Top             =   300
               Width           =   6400
            End
         End
         Begin TabDlg.SSTab TabINT 
            Height          =   2500
            Index           =   2
            Left            =   -74500
            TabIndex        =   75
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":3528
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprInt(6)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":3544
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprInt(7)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":3560
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprInt(8)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprInt 
               Height          =   1500
               Index           =   6
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":357C
               TabIndex        =   88
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprInt 
               Height          =   1500
               Index           =   7
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":3B6F
               TabIndex        =   89
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprInt 
               Height          =   1500
               Index           =   8
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":4162
               TabIndex        =   90
               Top             =   300
               Width           =   6400
            End
         End
      End
      Begin TabDlg.SSTab TabAEP_ES 
         Height          =   3300
         Left            =   -74600
         TabIndex        =   51
         Top             =   600
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5821
         _Version        =   327680
         TabOrientation  =   3
         Tab             =   2
         TabHeight       =   706
         ForeColor       =   9926553
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "LLEGADA"
         TabPicture(0)   =   "frmNacion_ES.frx":4755
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "TabAEP(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "SALIDA"
         TabPicture(1)   =   "frmNacion_ES.frx":4771
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TabAEP(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "TOTAL"
         TabPicture(2)   =   "frmNacion_ES.frx":478D
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "TabAEP(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin TabDlg.SSTab TabAEP 
            Height          =   2500
            Index           =   0
            Left            =   -74500
            TabIndex        =   52
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":47A9
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprAep(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":47C5
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprAep(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":47E1
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprAep(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprAep 
               Height          =   1500
               Index           =   0
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":47FD
               TabIndex        =   55
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprAep 
               Height          =   1500
               Index           =   1
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":4DF0
               TabIndex        =   54
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprAep 
               Height          =   1500
               Index           =   2
               Left            =   -74700
               OleObjectBlob   =   "frmNacion_ES.frx":53E3
               TabIndex        =   53
               Top             =   300
               Width           =   6400
            End
         End
         Begin TabDlg.SSTab TabAEP 
            Height          =   2500
            Index           =   1
            Left            =   -74800
            TabIndex        =   56
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":59D6
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprAep(3)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":59F2
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprAep(4)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":5A0E
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprAep(5)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprAep 
               Height          =   1500
               Index           =   5
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":5A2A
               TabIndex        =   59
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprAep 
               Height          =   1500
               Index           =   4
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":601D
               TabIndex        =   58
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprAep 
               Height          =   1500
               Index           =   3
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":6610
               TabIndex        =   57
               Top             =   300
               Width           =   6400
            End
         End
         Begin TabDlg.SSTab TabAEP 
            Height          =   2500
            Index           =   2
            Left            =   200
            TabIndex        =   74
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":6C03
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprAep(6)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":6C1F
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprAep(7)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":6C3B
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprAep(8)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprAep 
               Height          =   1500
               Index           =   6
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":6C57
               TabIndex        =   85
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprAep 
               Height          =   1500
               Index           =   7
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":724A
               TabIndex        =   86
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprAep 
               Height          =   1500
               Index           =   8
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":783D
               TabIndex        =   87
               Top             =   300
               Width           =   6400
            End
         End
      End
      Begin TabDlg.SSTab TabEZEB 
         Height          =   3300
         Left            =   -74600
         TabIndex        =   42
         Top             =   600
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5821
         _Version        =   327680
         TabOrientation  =   3
         Tab             =   2
         TabHeight       =   520
         ForeColor       =   9926553
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "LLEGADA"
         TabPicture(0)   =   "frmNacion_ES.frx":7E30
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "TabEspB(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "SALIDA"
         TabPicture(1)   =   "frmNacion_ES.frx":7E4C
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TabEspB(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "TOTAL"
         TabPicture(2)   =   "frmNacion_ES.frx":7E68
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "TabEspB(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin TabDlg.SSTab TabEspB 
            Height          =   2500
            Index           =   0
            Left            =   -74500
            TabIndex        =   43
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":7E84
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprEzeB(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":7EA0
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprEzeB(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":7EBC
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprEzeB(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   1500
               Index           =   0
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":7ED8
               TabIndex        =   46
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   1500
               Index           =   1
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":84CB
               TabIndex        =   45
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   1500
               Index           =   2
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":8AB0
               TabIndex        =   44
               Top             =   300
               Width           =   6400
            End
         End
         Begin TabDlg.SSTab TabEspB 
            Height          =   2500
            Index           =   1
            Left            =   -74500
            TabIndex        =   47
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":9095
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprEzeB(3)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":90B1
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprEzeB(4)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":90CD
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprEzeB(5)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   1500
               Index           =   5
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":90E9
               TabIndex        =   50
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   1500
               Index           =   4
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":96DC
               TabIndex        =   49
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   1500
               Index           =   3
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":9CCF
               TabIndex        =   48
               Top             =   300
               Width           =   6400
            End
         End
         Begin TabDlg.SSTab TabEspB 
            Height          =   2500
            Index           =   2
            Left            =   500
            TabIndex        =   73
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":A2C2
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprEzeB(6)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":A2DE
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprEzeB(7)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":A2FA
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprEzeB(8)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   1500
               Index           =   6
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":A316
               TabIndex        =   82
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   1500
               Index           =   7
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":A909
               TabIndex        =   83
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   1500
               Index           =   8
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":AEFC
               TabIndex        =   84
               Top             =   300
               Width           =   6400
            End
         End
      End
      Begin TabDlg.SSTab tabEZEA 
         Height          =   3300
         Left            =   -74600
         TabIndex        =   34
         Top             =   600
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5821
         _Version        =   327680
         TabOrientation  =   3
         Tab             =   2
         TabHeight       =   520
         ForeColor       =   9926553
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "LLEGADA"
         TabPicture(0)   =   "frmNacion_ES.frx":B4EF
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "TabEspA(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "SALIDA"
         TabPicture(1)   =   "frmNacion_ES.frx":B50B
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TabEspA(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "TOTAL"
         TabPicture(2)   =   "frmNacion_ES.frx":B527
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "TabEspA(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin TabDlg.SSTab TabEspA 
            Height          =   2500
            Index           =   1
            Left            =   -74500
            TabIndex        =   35
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":B543
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprEzeA(3)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":B55F
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprEzeA(4)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":B57B
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprEzeA(5)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   1500
               Index           =   3
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":B597
               TabIndex        =   69
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   1500
               Index           =   5
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":BB7C
               TabIndex        =   41
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   1500
               Index           =   4
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":C161
               TabIndex        =   40
               Top             =   300
               Width           =   6400
            End
         End
         Begin TabDlg.SSTab TabEspA 
            Height          =   2500
            Index           =   0
            Left            =   -74500
            TabIndex        =   36
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":C746
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprEzeA(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":C762
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprEzeA(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":C77E
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprEzeA(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   1500
               Index           =   0
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":C79A
               TabIndex        =   39
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   1500
               Index           =   1
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":CD7F
               TabIndex        =   38
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   1500
               Index           =   2
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":D364
               TabIndex        =   37
               Top             =   300
               Width           =   6400
            End
         End
         Begin TabDlg.SSTab TabEspA 
            Height          =   2500
            Index           =   2
            Left            =   500
            TabIndex        =   72
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":D949
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprEzeA(6)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":D965
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprEzeA(7)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":D981
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprEzeA(8)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   1500
               Index           =   6
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":D99D
               TabIndex        =   79
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   1500
               Index           =   7
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":DF82
               TabIndex        =   80
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   1500
               Index           =   8
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":E567
               TabIndex        =   81
               Top             =   300
               Width           =   6400
            End
         End
      End
      Begin TabDlg.SSTab TabResumen_ES 
         Height          =   3300
         Left            =   400
         TabIndex        =   22
         Top             =   600
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5821
         _Version        =   327680
         TabOrientation  =   3
         Tab             =   2
         TabHeight       =   520
         ForeColor       =   9926553
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "LLEGADA"
         TabPicture(0)   =   "frmNacion_ES.frx":EB4C
         Tab(0).ControlCount=   3
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "TabTotal(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "botCristal(0)"
         Tab(0).Control(1).Enabled=   -1  'True
         Tab(0).Control(2)=   "botExcel(0)"
         Tab(0).Control(2).Enabled=   -1  'True
         TabCaption(1)   =   "SALIDA"
         TabPicture(1)   =   "frmNacion_ES.frx":EB68
         Tab(1).ControlCount=   3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TabTotal(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "botCristal(1)"
         Tab(1).Control(1).Enabled=   -1  'True
         Tab(1).Control(2)=   "botExcel(3)"
         Tab(1).Control(2).Enabled=   -1  'True
         TabCaption(2)   =   "TOTAL"
         TabPicture(2)   =   "frmNacion_ES.frx":EB84
         Tab(2).ControlCount=   3
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "TabTotal(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "botExcel(6)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "botCristal(2)"
         Tab(2).Control(2).Enabled=   0   'False
         Begin VB.CommandButton botCristal 
            Height          =   525
            Index           =   2
            Left            =   6975
            Picture         =   "frmNacion_ES.frx":EBA0
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   510
            Index           =   6
            Left            =   6990
            Picture         =   "frmNacion_ES.frx":FCD6
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   900
            Width           =   705
         End
         Begin TabDlg.SSTab TabTotal 
            Height          =   2500
            Index           =   2
            Left            =   200
            TabIndex        =   71
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
            _ExtentY        =   4419
            _Version        =   327680
            TabOrientation  =   1
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
            TabCaption(0)   =   "General"
            TabPicture(0)   =   "frmNacion_ES.frx":10268
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprTot(6)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":10284
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprTot(7)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":102A0
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprTot(8)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprTot 
               Height          =   1500
               Index           =   6
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":102BC
               TabIndex        =   76
               Top             =   300
               Width           =   6405
            End
            Begin FPSpread.vaSpread SprTot 
               Height          =   1500
               Index           =   7
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":108A1
               TabIndex        =   77
               Top             =   300
               Width           =   6405
            End
            Begin FPSpread.vaSpread SprTot 
               Height          =   1500
               Index           =   8
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":10E86
               TabIndex        =   78
               Top             =   300
               Width           =   6405
            End
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   3
            Left            =   -68000
            Picture         =   "frmNacion_ES.frx":1146B
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   915
            Width           =   700
         End
         Begin VB.CommandButton botCristal 
            Height          =   500
            Index           =   1
            Left            =   -68000
            Picture         =   "frmNacion_ES.frx":119FD
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   300
            Width           =   700
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   0
            Left            =   -68000
            Picture         =   "frmNacion_ES.frx":12B33
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   930
            Width           =   700
         End
         Begin VB.CommandButton botCristal 
            Height          =   500
            Index           =   0
            Left            =   -68000
            Picture         =   "frmNacion_ES.frx":130C5
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   315
            Width           =   700
         End
         Begin TabDlg.SSTab TabTotal 
            Height          =   2500
            Index           =   0
            Left            =   -74800
            TabIndex        =   23
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":141FB
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprTot(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":14217
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprTot(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":14233
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprTot(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprTot 
               Height          =   1500
               Index           =   2
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":1424F
               TabIndex        =   26
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprTot 
               Height          =   1500
               Index           =   1
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":14834
               TabIndex        =   25
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprTot 
               Height          =   1500
               Index           =   0
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":14E19
               TabIndex        =   24
               Top             =   300
               Width           =   6405
            End
         End
         Begin TabDlg.SSTab TabTotal 
            Height          =   2500
            Index           =   1
            Left            =   -74800
            TabIndex        =   29
            Top             =   300
            Width           =   6700
            _ExtentX        =   11827
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
            TabPicture(0)   =   "frmNacion_ES.frx":153FE
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprTot(3)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Argentinos"
            TabPicture(1)   =   "frmNacion_ES.frx":1541A
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprTot(4)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Extranjeros"
            TabPicture(2)   =   "frmNacion_ES.frx":15436
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprTot(5)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprTot 
               Height          =   1500
               Index           =   3
               Left            =   150
               OleObjectBlob   =   "frmNacion_ES.frx":15452
               TabIndex        =   70
               Top             =   300
               Width           =   6405
            End
            Begin FPSpread.vaSpread SprTot 
               Height          =   1500
               Index           =   5
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":15A37
               TabIndex        =   33
               Top             =   300
               Width           =   6400
            End
            Begin FPSpread.vaSpread SprTot 
               Height          =   1500
               Index           =   4
               Left            =   -74850
               OleObjectBlob   =   "frmNacion_ES.frx":1601C
               TabIndex        =   32
               Top             =   300
               Width           =   6400
            End
         End
      End
      Begin Crystal.CrystalReport rptform 
         Left            =   165
         Top             =   2265
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   327682
         PrintFileLinesPerPage=   60
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
      Left            =   180
      TabIndex        =   21
      Top             =   6240
      Width           =   8730
   End
End
Attribute VB_Name = "frmNacion_ES"
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

For x = 0 To 8
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


Private Sub L_PonerNombres()
Dim i As Integer

For i = 0 To 8
    sprEzeA(i).MaxRows = 4
    sprEzeB(i).MaxRows = 4
    sprAEP(i).MaxRows = 4
    SprTot(i).MaxRows = 4
    
    sprEzeA(i).SetText 0, 1, "Período Anterior"
    sprEzeA(i).SetText 0, 2, "Período Actual"
    sprEzeA(i).SetText 0, 3, "Dif. Nominal"
    sprEzeA(i).SetText 0, 4, "Dif. Porcentual"
    
    sprEzeB(i).SetText 0, 1, "Período Anterior"
    sprEzeB(i).SetText 0, 2, "Período Actual"
    sprEzeB(i).SetText 0, 3, "Dif. Nominal"
    sprEzeB(i).SetText 0, 4, "Dif. Porcentual"
    
    sprAEP(i).SetText 0, 1, "Período Anterior"
    sprAEP(i).SetText 0, 2, "Período Actual"
    sprAEP(i).SetText 0, 3, "Dif. Nominal"
    sprAEP(i).SetText 0, 4, "Dif. Porcentual"
    
    SprTot(i).SetText 0, 1, "Período Anterior"
    SprTot(i).SetText 0, 2, "Período Actual"
    SprTot(i).SetText 0, 3, "Dif. Nominal"
    SprTot(i).SetText 0, 4, "Dif. Porcentual"
    
Next i

L_AltoCelda

End Sub

Private Sub MeImprimir(tipo As String)
Dim i%

'On Error GoTo ErrPrint:

frmNacion_ES.caption = Aplicacion.SeteoProceso(frmNacion_ES.caption)

rptform.Connect = "ODBC;UID=" & Aplicacion.username & ";PWD=" & Aplicacion.password & ";DATABASE=INTER;"

rptform.ReportFileName = RutaReportes
rptform.ReportFileName = RutaReportes & "NacTES.rpt"
rptform.Formulas(0) = "flaFDesdeUno = DATE(" & Year(mskFDesde.FormattedText) & "," & Month(mskFDesde.FormattedText) & "," & Day(mskFDesde.FormattedText) & ")"
rptform.Formulas(1) = "flaFHastaUno = DATE(" & Year(mskFHasta.FormattedText) & "," & Month(mskFHasta.FormattedText) & "," & Day(mskFHasta.FormattedText) & ")"
rptform.Formulas(2) = "flaFDesdeDos = DATE(" & Year(mskFDesdeAnt.FormattedText) & "," & Month(mskFDesdeAnt.FormattedText) & "," & Day(mskFDesdeAnt.FormattedText) & ")"
rptform.Formulas(3) = "flaFHastaDos = DATE(" & Year(mskFHastaAnt.FormattedText) & "," & Month(mskFHastaAnt.FormattedText) & "," & Day(mskFHastaAnt.FormattedText) & ")"

Select Case tipo
   Case "T"
       rptform.ParameterFields(0) = "PrTipo;T ;true"
   Case "L"
       rptform.ParameterFields(0) = "PrTipo;L ;true"
   Case "S"
       rptform.ParameterFields(0) = "PrTipo;S ;true"
End Select

rptform.Destination = 0     ' 0 pantalla
                            '1 impresora

rptform.Action = True

DoEvents

frmNacion_ES.caption = Aplicacion.SeteoFin

Exit Sub
'ErrPrint:
'     frmNacion_ES.caption = Aplicacion.SeteoFin
'     Exit Sub


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

For i = 0 To 8
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
Dim titulo As String

Do While Not RsData.EOF
        Select Case RsData!cod_depn
            Case DSLoc(1).Dep
                Select Case RsData!NAC
                  Case "N"
                   If RsData!tipo = "L" Then
                    sprAEP(1).SetText 1, 2, str(RsData!imp)
                    sprAEP(1).SetText 2, 2, str(RsData!vol)
                    sprAEP(1).SetText 5, 2, str(RsData!att)
                   Else
                    sprAEP(4).SetText 1, 2, str(RsData!imp)
                    sprAEP(4).SetText 2, 2, str(RsData!vol)
                    sprAEP(4).SetText 5, 2, str(RsData!att)
                   End If
                   Case "E"
                   If RsData!tipo = "L" Then
                    sprAEP(2).SetText 1, 2, str(RsData!imp)
                    sprAEP(2).SetText 2, 2, str(RsData!vol)
                    sprAEP(2).SetText 5, 2, str(RsData!att)
                   Else
                    sprAEP(5).SetText 1, 2, str(RsData!imp)
                    sprAEP(5).SetText 2, 2, str(RsData!vol)
                    sprAEP(5).SetText 5, 2, str(RsData!att)
                   End If
                End Select
            
            Case DSLoc(2).Dep
                
                Select Case RsData!Cod_Sdep
                    Case DSLoc(2).Sdep
                      Select Case RsData!NAC
                        Case Is = "N"
                          If RsData!tipo = "L" Then
                            sprEzeA(1).SetText 1, 2, str(RsData!imp)
                            sprEzeA(1).SetText 2, 2, str(RsData!vol)
                            sprEzeA(1).SetText 5, 2, str(RsData!att)
                           Else
                            sprEzeA(4).SetText 1, 2, str(RsData!imp)
                            sprEzeA(4).SetText 2, 2, str(RsData!vol)
                            sprEzeA(4).SetText 5, 2, str(RsData!att)
                           End If
                        Case Is = "E"
                          If RsData!tipo = "L" Then
                            sprEzeA(2).SetText 1, 2, str(RsData!imp)
                            sprEzeA(2).SetText 2, 2, str(RsData!vol)
                            sprEzeA(2).SetText 5, 2, str(RsData!att)
                          Else
                            sprEzeA(5).SetText 1, 2, str(RsData!imp)
                            sprEzeA(5).SetText 2, 2, str(RsData!vol)
                            sprEzeA(5).SetText 5, 2, str(RsData!att)
                          End If
                
                      End Select
                      
                    Case DSLoc(3).Sdep
                      Select Case RsData!NAC
                        Case Is = "N"
                          If RsData!tipo = "L" Then
                            sprEzeB(1).SetText 1, 2, str(RsData!imp)
                            sprEzeB(1).SetText 2, 2, str(RsData!vol)
                            sprEzeB(1).SetText 5, 2, str(RsData!att)
                          Else
                            sprEzeB(4).SetText 1, 2, str(RsData!imp)
                            sprEzeB(4).SetText 2, 2, str(RsData!vol)
                            sprEzeB(4).SetText 5, 2, str(RsData!att)
                          End If
                        Case Is = "E"
                          If RsData!tipo = "L" Then
                            sprEzeB(2).SetText 1, 2, str(RsData!imp)
                            sprEzeB(2).SetText 2, 2, str(RsData!vol)
                            sprEzeB(2).SetText 5, 2, str(RsData!att)
                           Else
                            sprEzeB(5).SetText 1, 2, str(RsData!imp)
                            sprEzeB(5).SetText 2, 2, str(RsData!vol)
                            sprEzeB(5).SetText 5, 2, str(RsData!att)
                           End If
                      
                      End Select
                              
                End Select
            
            Case DSLoc(4).Dep  'Interior
               Select Case RsData!NAC
                  Case "N"
                  '          SprInt(1).SetText 1, SprInt(1).MaxRows, str(RsData!imp)
                  '          SprInt(1).SetText 2, SprInt(1).MaxRows, str(RsData!vol)
                  '          SprInt(1).SetText 5, SprInt(1).MaxRows, str(RsData!att)
                        Case Is = "E"
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
                   If RsDataAnt!tipo = "L" Then
                    'SprAep(1).MaxRows = SprAep(1).MaxRows + 1
                    'SprAep(1).SetText 0, SprAep(1).MaxRows, "Período Anterior"
                    sprAEP(1).SetText 1, 1, str(RsDataAnt!imp)
                    sprAEP(1).SetText 2, 1, str(RsDataAnt!vol)
                    sprAEP(1).SetText 5, 1, str(RsDataAnt!att)
                  Else
                    'SprAep(4).MaxRows = SprAep(4).MaxRows + 1
                    'SprAep(4).SetText 0, SprAep(4).MaxRows, "Período Anterior"
                    sprAEP(4).SetText 1, 1, str(RsDataAnt!imp)
                    sprAEP(4).SetText 2, 1, str(RsDataAnt!vol)
                    sprAEP(4).SetText 5, 1, str(RsDataAnt!att)
                  End If
                 Case "E"
                   If RsDataAnt!tipo = "L" Then
                    'SprAep(2).MaxRows = SprAep(2).MaxRows + 1
                    'SprAep(2).SetText 0, SprAep(2).MaxRows, "Período Anterior"
                    sprAEP(2).SetText 1, 1, str(RsDataAnt!imp)
                    sprAEP(2).SetText 2, 1, str(RsDataAnt!vol)
                    sprAEP(2).SetText 5, 1, str(RsDataAnt!att)
                  Else
                    'SprAep(5).MaxRows = SprAep(5).MaxRows + 1
                    'SprAep(5).SetText 0, SprAep(5).MaxRows, "Período Anterior"
                    sprAEP(5).SetText 1, 1, str(RsDataAnt!imp)
                    sprAEP(5).SetText 2, 1, str(RsDataAnt!vol)
                    sprAEP(5).SetText 5, 1, str(RsDataAnt!att)
                  End If
                End Select
            Case DSLoc(2).Dep
                
                Select Case RsDataAnt!Cod_Sdep
                    Case DSLoc(2).Sdep
                      Select Case RsDataAnt!NAC
                        Case Is = "N"
                          If RsDataAnt!tipo = "L" Then
                            'SprEzeA(1).MaxRows = SprEzeA(1).MaxRows + 1
                            'SprEzeA(1).SetText 0, SprEzeA(1).MaxRows, "Período Anterior"
                            sprEzeA(1).SetText 1, 1, str(RsDataAnt!imp)
                            sprEzeA(1).SetText 2, 1, str(RsDataAnt!vol)
                            sprEzeA(1).SetText 5, 1, str(RsDataAnt!att)
                           Else
                            'SprEzeA(4).MaxRows = SprEzeA(4).MaxRows + 1
                            'SprEzeA(4).SetText 0, SprEzeA(4).MaxRows, "Período Anterior"
                            sprEzeA(4).SetText 1, 1, str(RsDataAnt!imp)
                            sprEzeA(4).SetText 2, 1, str(RsDataAnt!vol)
                            sprEzeA(4).SetText 5, 1, str(RsDataAnt!att)
                         End If
                        Case Is = "E"
                          If RsDataAnt!tipo = "L" Then
                            'SprEzeA(2).MaxRows = SprEzeA(2).MaxRows + 1
                            'SprEzeA(2).SetText 0, SprEzeA(2).MaxRows, "Período Anterior"
                            sprEzeA(2).SetText 1, 1, str(RsDataAnt!imp)
                            sprEzeA(2).SetText 2, 1, str(RsDataAnt!vol)
                            sprEzeA(2).SetText 5, 1, str(RsDataAnt!att)
                           Else
                            'SprEzeA(5).MaxRows = SprEzeA(5).MaxRows + 1
                            'SprEzeA(5).SetText 0, SprEzeA(5).MaxRows, "Período Anterior"
                            sprEzeA(5).SetText 1, 1, str(RsDataAnt!imp)
                            sprEzeA(5).SetText 2, 1, str(RsDataAnt!vol)
                            sprEzeA(5).SetText 5, 1, str(RsDataAnt!att)
                          End If
                      End Select
                      
                    Case DSLoc(3).Sdep
                      Select Case RsDataAnt!NAC
                        Case Is = "N"
                          If RsDataAnt!tipo = "L" Then
                            'SprEzeB(1).MaxRows = SprEzeB(1).MaxRows + 1
                            'SprEzeB(1).SetText 0, SprEzeB(1).MaxRows, "Período Anterior"
                            sprEzeB(1).SetText 1, 1, str(RsDataAnt!imp)
                            sprEzeB(1).SetText 2, 1, str(RsDataAnt!vol)
                            sprEzeB(1).SetText 5, 1, str(RsDataAnt!att)
                           Else
                            'SprEzeB(4).MaxRows = SprEzeB(4).MaxRows + 1
                            'SprEzeB(4).SetText 0, SprEzeB(4).MaxRows, "Período Anterior"
                            sprEzeB(4).SetText 1, 1, str(RsDataAnt!imp)
                            sprEzeB(4).SetText 2, 1, str(RsDataAnt!vol)
                            sprEzeB(4).SetText 5, 1, str(RsDataAnt!att)
                         End If
                        Case Is = "E"
                          If RsDataAnt!tipo = "L" Then
                            'SprEzeB(2).MaxRows = SprEzeB(2).MaxRows + 1
                            'SprEzeB(2).SetText 0, SprEzeB(2).MaxRows, "Período Anterior"
                            sprEzeB(2).SetText 1, 1, str(RsDataAnt!imp)
                            sprEzeB(2).SetText 2, 1, str(RsDataAnt!vol)
                            sprEzeB(2).SetText 5, 1, str(RsDataAnt!att)
                           Else
                            'SprEzeB(5).MaxRows = SprEzeB(5).MaxRows + 1
                            'SprEzeB(5).SetText 0, SprEzeB(5).MaxRows, "Período Anterior"
                            sprEzeB(5).SetText 1, 1, str(RsDataAnt!imp)
                            sprEzeB(5).SetText 2, 1, str(RsDataAnt!vol)
                            sprEzeB(5).SetText 5, 1, str(RsDataAnt!att)
                          End If
                            
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

For i = 0 To 8

Func_FormatRange 2, 2, 1, 3, sprEzeA(i), "99999999.00"
Func_FormatRange 2, 2, 1, 3, sprEzeB(i), "99999999.00"
Func_FormatRange 2, 2, 1, 3, sprAEP(i), "99999999.00"
Func_FormatRange 2, 2, 1, 3, sprInt(i), "99999999.00"
Func_FormatRange 2, 2, 1, 3, SprTot(i), "99999999.00"

Next

End Sub

Private Sub L_LimpiarGrillas()
Dim i As Integer

For i = 0 To 8
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

L_Llenargrilla sprEzeA(4), sprEzeA(5), sprEzeA(3)
L_Llenargrilla sprEzeB(4), sprEzeB(5), sprEzeB(3)
L_Llenargrilla sprAEP(4), sprAEP(5), sprAEP(3)
'L_Llenargrilla SprInt(4), SprInt(5), SprInt(3)
 
End Sub

Private Sub L_LlenarResumen()
Dim i As Integer
For i = 0 To 5
   SprTot(i).MaxRows = 4
Next

L_LlenarGrillaResumen sprEzeA(2), sprEzeB(2), sprAEP(2), sprInt(2), SprTot(2)
L_LlenarGrillaResumen sprEzeA(1), sprEzeB(1), sprAEP(1), sprInt(1), SprTot(1)
L_LlenarGrillaResumen sprEzeA(4), sprEzeB(4), sprAEP(4), sprInt(4), SprTot(4)
L_LlenarGrillaResumen sprEzeA(5), sprEzeB(5), sprAEP(5), sprInt(5), SprTot(5)
L_Llenargrilla SprTot(1), SprTot(2), SprTot(0)
L_Llenargrilla SprTot(4), SprTot(5), SprTot(3)

For i = 0 To 5
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

For i = 4 To 5

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

Private Sub L_LlenarTabTotal()

Dim i As Integer

'Llena las grillas de Argentinos y Extranjeros

L_Llenargrilla sprEzeA(1), sprEzeA(4), sprEzeA(7)
L_Llenargrilla sprEzeB(1), sprEzeB(4), sprEzeB(7)
L_Llenargrilla sprAEP(1), sprAEP(4), sprAEP(7)
'L_Llenargrilla SprInt(1), SprInt(4), SprInt(7)

L_Llenargrilla sprEzeA(2), sprEzeA(5), sprEzeA(8)
L_Llenargrilla sprEzeB(2), sprEzeB(5), sprEzeB(8)
L_Llenargrilla sprAEP(2), sprAEP(5), sprAEP(8)
'L_Llenargrilla SprInt(2), SprInt(5), SprInt(8)

'Llena las grillas de General

L_Llenargrilla sprEzeA(7), sprEzeA(8), sprEzeA(6)
L_Llenargrilla sprEzeB(7), sprEzeB(8), sprEzeB(6)
L_Llenargrilla sprAEP(7), sprAEP(8), sprAEP(6)
'L_Llenargrilla SprInt(7), SprInt(8), SprInt(6)



For i = 6 To 8    '7
    L_DiferenciaNominal sprEzeA(i)
    L_DiferenciaNominal sprEzeB(i)
    L_DiferenciaNominal sprAEP(i)
'    L_DiferenciaNominal SprInt(i)
    
    L_DiferenciaPorcentual sprEzeA(i)
    L_DiferenciaPorcentual sprEzeB(i)
    L_DiferenciaPorcentual sprAEP(i)
'    L_DiferenciaPorcentual SprInt(i)
Next


'Llena las Grillas de Resumen

For i = 6 To 8
   SprTot(i).MaxRows = 4
Next

L_LlenarGrillaResumen sprEzeA(7), sprEzeB(7), sprAEP(7), sprInt(7), SprTot(7)
L_LlenarGrillaResumen sprEzeA(8), sprEzeB(8), sprAEP(8), sprInt(8), SprTot(8)

L_Llenargrilla SprTot(7), SprTot(8), SprTot(6)

For i = 6 To 8
    L_DiferenciaNominal SprTot(i)
    L_DiferenciaPorcentual SprTot(i)
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
    
L_DiferenciaNominal sprEzeA(3)
L_DiferenciaNominal sprEzeB(3)
L_DiferenciaNominal sprAEP(3)
'L_DiferenciaNominal SprInt(3)
 
L_DiferenciaPorcentual sprEzeA(3)
L_DiferenciaPorcentual sprEzeB(3)
L_DiferenciaPorcentual sprAEP(3)
'L_DiferenciaPorcentual SprInt(3)

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
    
    If valor1 = 0 Or Valor2 = 0 Then
       DifPorcentual = 0
     Else
       DifPorcentual = (valor1 * 100) / Valor2
    End If
    
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

frmNacion_ES.caption = Aplicacion.SeteoProceso(frmNacion_ES.caption)


sql = " SELECT "
sql = sql & " cod_depn, "
sql = sql & " cod_sdep, "
sql = sql & " nac, "
sql = sql & " tipo, "
sql = sql & " sum(imp)imp, "
sql = sql & " sum(vol) vol ,"
sql = sql & " sum(att) att "
sql = sql & "FROM estadis.VIEW_PAX_RESUMEN "
sql = sql & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)))
sql = sql & "group by cod_depn,cod_sdep,nac,tipo"
sql = sql & " order by cod_depn,cod_sdep,nac,tipo"


sqlX = " SELECT "
sqlX = sqlX & " cod_depn, "
sqlX = sqlX & " cod_sdep, "
sqlX = sqlX & " nac, "
sqlX = sqlX & " tipo, "
sqlX = sqlX & " sum(imp)imp, "
sqlX = sqlX & " sum(vol) vol ,"
sqlX = sqlX & " sum(att) att "
sqlX = sqlX & "FROM estadis.VIEW_PAX_RESUMEN "
sqlX = sqlX & L_Armarcondicion(Year(CDate(mskFDesdeAnt.FormattedText)) - 1)
sqlX = sqlX & "group by cod_depn,cod_sdep,nac,tipo"
sqlX = sqlX & " order by cod_depn,cod_sdep,nac,tipo"


L_LimpiarGrillas
L_ControlaVolados
L_PonerNombres

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
        L_LlenarTabTotal
        L_FormatearCelda
        L_AltoCelda
        L_AplicaFormula
        L_Resaltar
        
        
        Aplicacion.CerrarDAO RsData
        Aplicacion.CerrarDAO RsDataAnt
      
      
      Else
        If Aplicacion.ObtenerRsDAO(sql, RsData) Then
            L_DecoEspigon
        End If
        
        L_LlenarDiferencias
        L_LlenarGenerales
        L_LlenarDiferenciasGral
        L_LlenarResumen
        L_LlenarTabTotal
        L_FormatearCelda
        L_AltoCelda
        L_AplicaFormula
        L_Resaltar
        
        Aplicacion.CerrarDAO RsData
    End If
    
'    Aplicacion.CerrarDAO RsDataAnt

End If


ErrInd:
    frmNacion_ES.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub


Private Sub L_ControlaVolados()
Dim FH As String
Dim sql As String
Dim rs_pax As Recordset

sql = " SELECT MIN(fch_actualizado) fch FROM estadis.control_carga "
sql = sql & "  WHERE cod_sdep <> 'INT' "

If Aplicacion.ObtenerRsDAO(sql, rs_pax) Then
   If CDate(mskFHasta.FormattedText) <= CDate(rs_pax!fch) Then
      labInfo.caption = ""
      Timer1.Enabled = False
    Else
      labInfo.caption = "Datos de Pax transitados cargados al " & rs_pax!fch
      Timer1.Enabled = True
   End If

End If

End Sub

Private Sub L_Resaltar()
Dim fila As Integer
Dim i As Integer

For i = 0 To 8

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







Private Sub BotCristal_Click(Index As Integer)

Select Case Index
   Case 0
      MeImprimir ("L")
   Case 1
      MeImprimir ("S")
   Case 2
      MeImprimir ("T")
End Select

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



Private Sub botExcel_Click(Index As Integer)
   
Select Case Index
  Case 0
        L_TratarExcel_Llegadas SprTot(0), "Análisis de Ventas por Pasajeros Viajados por Nacionalidad ", "", "", Exl_Blanco
  Case 3
        L_TratarExcel_Salidas SprTot(3), "Análisis de Ventas por Pasajeros Viajados por Nacionalidad ", "", "", Exl_Blanco
  Case 6
        L_TratarExcel_Todos SprTot(6), "Análisis de Ventas por Pasajeros Viajados por Nacionalidad ", "", "", Exl_Blanco
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


Private Sub L_TratarExcel_Llegadas(spr As control, titulo As String, subTit As String, Info As String, color As Integer)
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

frmNacion_ES.caption = Aplicacion.SeteoProceso(frmNacion_ES.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.application.Visible = true
    
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
    
     fila = fila + 2
    
    Exl_PonerValor AppExcel, fila, Col, "Tipo de Vuelos : LLEGADAS"
    rango = Exl_rangos(fila, fila, 1, 3)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    
    
    fila = fila + 2
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia - Llegadas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia Argentinos - Llegadas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia Extranjeros - Llegadas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A General - Llegadas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A Argentinos - Llegadas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A Argentinos - Llegadas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B General - Llegadas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B Argentinos - Llegadas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B Extranjeros - Llegadas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque General - Llegadas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque Argentinos - Llegadas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque Extranjeros - Llegadas"
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

    frmNacion_ES.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub L_TratarExcel_Salidas(spr As control, titulo As String, subTit As String, Info As String, color As Integer)
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

frmNacion_ES.caption = Aplicacion.SeteoProceso(frmNacion_ES.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.application.Visible = true
    
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
    
     fila = fila + 2
    
    Exl_PonerValor AppExcel, fila, Col, "Tipo de Vuelos : SALIDAS"
    rango = Exl_rangos(fila, fila, 1, 3)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    
    
    fila = fila + 2
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia - Salidas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia Argentinos - Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel SprTot(4), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + SprTot(4).MaxRows, Col + 1, SprTot(4).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + SprTot(4).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, SprTot(4).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - SprTot(4).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
    '-----------------------------------
   
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia Extranjeros - Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel SprTot(5), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + SprTot(5).MaxRows, Col + 1, SprTot(5).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + SprTot(5).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, SprTot(5).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - SprTot(5).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A General - Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeA(3), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeA(3).MaxRows, Col + 1, sprEzeA(3).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeA(3).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeA(3).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeA(3).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A Argentinos - Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeA(4), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeA(4).MaxRows, Col + 1, sprEzeA(4).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeA(4).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeA(4).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeA(4).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
       
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A Argentinos - Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeA(5), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeA(5).MaxRows, Col + 1, sprEzeA(5).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeA(5).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeA(5).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeA(5).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
   
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B General - Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeB(3), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeB(3).MaxRows, Col + 1, sprEzeB(3).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeB(3).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeB(3).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeB(3).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B Argentinos - Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeB(4), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeB(4).MaxRows, Col + 1, sprEzeB(4).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeB(4).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeB(4).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeB(4).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
       
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B Extranjeros - Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeB(5), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeB(5).MaxRows, Col + 1, sprEzeB(5).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeB(5).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeB(5).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeB(5).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque General - Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprAEP(3), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprAEP(3).MaxRows, Col + 1, sprAEP(3).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprAEP(3).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprAEP(3).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprAEP(3).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque Argentinos - Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprAEP(4), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprAEP(4).MaxRows, Col + 1, sprAEP(4).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprAEP(4).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprAEP(4).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprAEP(4).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
       
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque Extranjeros - Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprAEP(5), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprAEP(5).MaxRows, Col + 1, sprAEP(5).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprAEP(5).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprAEP(5).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprAEP(5).MaxRows, fila - 1, 5, 5)
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

    frmNacion_ES.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub L_TratarExcel_Todos(spr As control, titulo As String, subTit As String, Info As String, color As Integer)
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

frmNacion_ES.caption = Aplicacion.SeteoProceso(frmNacion_ES.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.application.Visible = true
    
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
    
     fila = fila + 2
    
    Exl_PonerValor AppExcel, fila, Col, "Tipo de Vuelos : LLEGADAS / SALIDAS"
    rango = Exl_rangos(fila, fila, 1, 3)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    
    
    fila = fila + 2
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia - Llegadas / Salidas"
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
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia Argentinos - Llegadas / Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel SprTot(7), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + SprTot(7).MaxRows, Col + 1, SprTot(7).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + SprTot(7).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, SprTot(7).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - SprTot(7).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
    '-----------------------------------
   
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compañia Extranjeros - Llegadas / Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel SprTot(8), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + SprTot(8).MaxRows, Col + 1, SprTot(8).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + SprTot(8).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, SprTot(8).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - SprTot(8).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A General - Llegadas / Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeA(6), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeA(6).MaxRows, Col + 1, sprEzeA(6).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeA(6).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeA(6).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeA(6).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A Argentinos - Llegadas / Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeA(7), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeA(7).MaxRows, Col + 1, sprEzeA(7).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeA(7).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeA(7).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeA(7).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
       
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional A Extranjeros - Llegadas / Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeA(8), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeA(8).MaxRows, Col + 1, sprEzeA(8).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeA(8).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeA(8).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeA(8).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
   
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B General - Llegadas / Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeB(6), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeB(6).MaxRows, Col + 1, sprEzeB(6).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeB(6).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeB(6).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeB(6).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B Argentinos - Llegadas / Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeB(7), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeB(7).MaxRows, Col + 1, sprEzeB(7).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeB(7).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeB(7).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeB(7).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
       
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Internacional B Extranjeros - Llegadas / Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprEzeB(8), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeB(8).MaxRows, Col + 1, sprEzeB(8).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprEzeB(8).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprEzeB(8).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprEzeB(8).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque General - Llegadas / Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprAEP(6), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprAEP(6).MaxRows, Col + 1, sprAEP(6).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprAEP(6).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprAEP(6).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprAEP(6).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque Argentinos - Llegadas / Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprAEP(7), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprAEP(7).MaxRows, Col + 1, sprAEP(7).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprAEP(7).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprAEP(7).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprAEP(7).MaxRows, fila - 1, 5, 5)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_AnchoCol AppExcel, 1, 1, 20
    Exl_AnchoCol AppExcel, 2, 3, 17
   
   '---------------------------------
       
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Total Aeroparque Extranjeros - Llegadas / Salidas"
    rango = Exl_rangos(fila, fila, 1, 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    
    L_BajarGrillaExel sprAEP(8), AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprAEP(8).MaxRows, Col + 1, sprAEP(8).MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + sprAEP(8).MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, sprAEP(8).MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - sprAEP(8).MaxRows, fila - 1, 5, 5)
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

    frmNacion_ES.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub

Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 7000
Me.Width = 9300

mskFDesde.Text = "01-" & Format$(Month(Date), "0#") & "-" & Format$(Year(Date), "####")
mskFHasta.Text = Format$(Date - 1, FTOFECHA)

mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio, FTOFECHA)

'L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmNacion_ES"

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


