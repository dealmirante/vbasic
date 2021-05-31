VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConsultaConcurso 
   Caption         =   "Seguimiento de unidades vendidas por Producto"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
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
      Left            =   8460
      Picture         =   "frmConsultaConcurso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
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
      Left            =   8460
      Picture         =   "frmConsultaConcurso.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   14
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
      Left            =   8460
      Picture         =   "frmConsultaConcurso.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   30
      Width           =   570
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   4545
      Left            =   120
      TabIndex        =   1
      Top             =   1185
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   8017
      _Version        =   327680
      Tabs            =   5
      TabsPerRow      =   5
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
      TabPicture(0)   =   "frmConsultaConcurso.frx":0A26
      Tab(0).ControlCount=   25
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "labGr(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "labGr(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "labGr(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label7"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "labImp"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "labGr(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label5(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "frA"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtPers(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtPers(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtPers(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtPers(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "botExel"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtGr(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtGr(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtGr(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtGr(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtPers(3)"
      Tab(0).Control(24).Enabled=   0   'False
      TabCaption(1)   =   "EZE"
      TabPicture(1)   =   "frmConsultaConcurso.frx":0A42
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabEzeA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "botExcelEzeA"
      Tab(1).Control(1).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmConsultaConcurso.frx":0A5E
      Tab(2).ControlCount=   3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "botExcelEzeB"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "tabEzeB"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label9(4)"
      Tab(2).Control(2).Enabled=   0   'False
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmConsultaConcurso.frx":0A7A
      Tab(3).ControlCount=   2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "botExcelAep"
      Tab(3).Control(0).Enabled=   -1  'True
      Tab(3).Control(1)=   "tabAep"
      Tab(3).Control(1).Enabled=   0   'False
      TabCaption(4)   =   "INT"
      TabPicture(4)   =   "frmConsultaConcurso.frx":0A96
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "tabInt"
      Tab(4).Control(0).Enabled=   0   'False
      Begin VB.TextBox txtPers 
         Height          =   330
         Index           =   3
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   3225
         Width           =   1665
      End
      Begin VB.TextBox txtGr 
         Height          =   285
         Index           =   3
         Left            =   3315
         TabIndex        =   97
         Top             =   3255
         Visible         =   0   'False
         Width           =   315
      End
      Begin TabDlg.SSTab tabInt 
         Height          =   3555
         Left            =   -74235
         TabIndex        =   85
         Top             =   555
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   6271
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabHeight       =   520
         BackColor       =   12632256
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
         TabPicture(0)   =   "frmConsultaConcurso.frx":0AB2
         Tab(0).ControlCount=   10
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lab1(6)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lab1(7)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label9(7)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "labPorcT(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label9(8)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "labPorcM(3)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "labPorcN(3)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtT(3)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtM(3)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "sprInt(0)"
         Tab(0).Control(9).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmConsultaConcurso.frx":0ACE
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprInt(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprInt 
            Height          =   2925
            Index           =   1
            Left            =   -74460
            OleObjectBlob   =   "frmConsultaConcurso.frx":0AEA
            TabIndex        =   87
            Top             =   135
            Width           =   6165
         End
         Begin FPSpread.vaSpread sprInt 
            Height          =   2190
            Index           =   0
            Left            =   450
            OleObjectBlob   =   "frmConsultaConcurso.frx":1021
            TabIndex        =   86
            Top             =   150
            Width           =   6360
         End
         Begin VB.TextBox txtM 
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
            Height          =   315
            Index           =   3
            Left            =   3090
            Locked          =   -1  'True
            TabIndex        =   91
            Top             =   2400
            Width           =   1560
         End
         Begin VB.TextBox txtT 
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
            Height          =   315
            Index           =   3
            Left            =   3090
            Locked          =   -1  'True
            TabIndex        =   90
            Top             =   2805
            Width           =   1545
         End
         Begin VB.Label labPorcN 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   3
            Left            =   6270
            TabIndex        =   106
            Top             =   2820
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label labPorcM 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   3
            Left            =   4725
            TabIndex        =   95
            Top             =   2445
            Width           =   870
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   5670
            TabIndex        =   94
            Top             =   2430
            Width           =   315
         End
         Begin VB.Label labPorcT 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   3
            Left            =   4740
            TabIndex        =   93
            Top             =   2835
            Width           =   870
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   5670
            TabIndex        =   92
            Top             =   2805
            Width           =   315
         End
         Begin VB.Label lab1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tot Rubro - T. Mañana (C)"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   465
            TabIndex        =   89
            Top             =   2400
            Width           =   2535
         End
         Begin VB.Label lab1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tot Rubro - T. Tarde/Noche (M)"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   465
            TabIndex        =   88
            Top             =   2805
            Width           =   2535
         End
      End
      Begin VB.TextBox txtGr 
         Height          =   285
         Index           =   2
         Left            =   3315
         TabIndex        =   84
         Top             =   2685
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtGr 
         Height          =   285
         Index           =   1
         Left            =   3300
         TabIndex        =   83
         Top             =   2460
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtGr 
         Height          =   285
         Index           =   0
         Left            =   3300
         TabIndex        =   82
         Top             =   2115
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton botExel 
         Height          =   705
         Left            =   7380
         Picture         =   "frmConsultaConcurso.frx":1556
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2430
         Width           =   810
      End
      Begin VB.TextBox txtPers 
         Height          =   330
         Index           =   4
         Left            =   5250
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   3645
         Width           =   1665
      End
      Begin VB.TextBox txtPers 
         Height          =   330
         Index           =   2
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   2715
         Width           =   1665
      End
      Begin VB.TextBox txtPers 
         Height          =   330
         Index           =   1
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2490
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox txtPers 
         Height          =   330
         Index           =   0
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2100
         Width           =   1665
      End
      Begin VB.Frame frA 
         Height          =   1110
         Left            =   315
         TabIndex        =   31
         Top             =   315
         Width           =   8130
         Begin VB.CommandButton botBases 
            Caption         =   "Productos"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   6540
            TabIndex        =   51
            Top             =   285
            Width           =   1005
         End
         Begin VB.TextBox txtImp 
            Height          =   330
            Left            =   4020
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   645
            Width           =   1665
         End
         Begin VB.TextBox txtCant 
            Height          =   330
            Left            =   4035
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   240
            Width           =   1665
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Importe Total a Distribuir"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   1485
            TabIndex        =   33
            Top             =   645
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total de Unidades Vendidas"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   1500
            TabIndex        =   32
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.CommandButton botExcelAep 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67125
         Picture         =   "frmConsultaConcurso.frx":1AE8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3885
         Width           =   765
      End
      Begin VB.CommandButton botExcelEzeB 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67125
         Picture         =   "frmConsultaConcurso.frx":207A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3900
         Width           =   765
      End
      Begin VB.CommandButton botExcelEzeA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67155
         Picture         =   "frmConsultaConcurso.frx":260C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3885
         Width           =   765
      End
      Begin TabDlg.SSTab tabEzeA 
         Height          =   3555
         Left            =   -74235
         TabIndex        =   2
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6271
         _Version        =   327680
         TabOrientation  =   1
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
         TabPicture(0)   =   "frmConsultaConcurso.frx":2B9E
         Tab(0).ControlCount=   13
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lab1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lab1(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "labPorcM(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label9(0)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label9(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "labPorcT(0)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label9(9)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "labPorcN(0)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lab1(8)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtM(0)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtT(0)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtN(0)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "sprEzeA(0)"
         Tab(0).Control(12).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmConsultaConcurso.frx":2BBA
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprEzeA(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "General"
         TabPicture(2)   =   "frmConsultaConcurso.frx":2BD6
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprRes(0)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprRes 
            Height          =   2970
            Index           =   0
            Left            =   -74835
            OleObjectBlob   =   "frmConsultaConcurso.frx":2BF2
            TabIndex        =   53
            Top             =   165
            Width           =   8310
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2940
            Index           =   1
            Left            =   -74475
            OleObjectBlob   =   "frmConsultaConcurso.frx":3455
            TabIndex        =   26
            Top             =   150
            Width           =   6180
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   1800
            Index           =   0
            Left            =   405
            OleObjectBlob   =   "frmConsultaConcurso.frx":398C
            TabIndex        =   25
            Top             =   165
            Width           =   6375
         End
         Begin VB.TextBox txtN 
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
            Height          =   315
            Index           =   0
            Left            =   2835
            Locked          =   -1  'True
            TabIndex        =   100
            Top             =   2820
            Width           =   1470
         End
         Begin VB.TextBox txtT 
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
            Height          =   315
            Index           =   0
            Left            =   2820
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   2430
            Width           =   1470
         End
         Begin VB.TextBox txtM 
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
            Height          =   315
            Index           =   0
            Left            =   2820
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   2040
            Width           =   1470
         End
         Begin VB.Label lab1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tot Rubro - T. Noche"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   420
            TabIndex        =   103
            Top             =   2820
            Width           =   2220
         End
         Begin VB.Label labPorcN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   0
            Left            =   4425
            TabIndex        =   102
            Top             =   2850
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   5295
            TabIndex        =   101
            Top             =   2790
            Width           =   315
         End
         Begin VB.Label labPorcT 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   0
            Left            =   4410
            TabIndex        =   76
            Top             =   2520
            Width           =   870
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   5280
            TabIndex        =   71
            Top             =   2400
            Width           =   315
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   5310
            TabIndex        =   70
            Top             =   2085
            Width           =   315
         End
         Begin VB.Label labPorcM 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   0
            Left            =   4410
            TabIndex        =   69
            Top             =   2100
            Width           =   855
         End
         Begin VB.Label lab1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tot Rubro - T. Tarde"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   405
            TabIndex        =   58
            Top             =   2430
            Width           =   2220
         End
         Begin VB.Label lab1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tot Rubro - T. Mañana"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   405
            TabIndex        =   56
            Top             =   2040
            Width           =   2220
         End
      End
      Begin TabDlg.SSTab tabAep 
         Height          =   3525
         Left            =   -74265
         TabIndex        =   3
         Top             =   525
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   6218
         _Version        =   327680
         TabOrientation  =   1
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
         TabPicture(0)   =   "frmConsultaConcurso.frx":3EA0
         Tab(0).ControlCount=   10
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lab1(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lab1(5)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "labPorcM(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label9(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "labPorcT(2)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label9(6)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "labPorcN(2)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtM(2)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtT(2)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "sprAep(0)"
         Tab(0).Control(9).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmConsultaConcurso.frx":3EBC
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprAep(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "General"
         TabPicture(2)   =   "frmConsultaConcurso.frx":3ED8
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprRes(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprRes 
            Height          =   2985
            Index           =   2
            Left            =   -74835
            OleObjectBlob   =   "frmConsultaConcurso.frx":3EF4
            TabIndex        =   55
            Top             =   150
            Width           =   8220
         End
         Begin FPSpread.vaSpread sprAep 
            Height          =   2880
            Index           =   1
            Left            =   -74445
            OleObjectBlob   =   "frmConsultaConcurso.frx":47FD
            TabIndex        =   30
            Top             =   150
            Width           =   6210
         End
         Begin FPSpread.vaSpread sprAep 
            Height          =   2205
            Index           =   0
            Left            =   480
            OleObjectBlob   =   "frmConsultaConcurso.frx":4D26
            TabIndex        =   29
            Top             =   135
            Width           =   6420
         End
         Begin VB.TextBox txtT 
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
            Height          =   315
            Index           =   2
            Left            =   2805
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   2850
            Width           =   1545
         End
         Begin VB.TextBox txtM 
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
            Height          =   315
            Index           =   2
            Left            =   2805
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   2445
            Width           =   1560
         End
         Begin VB.Label labPorcN 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   2
            Left            =   6060
            TabIndex        =   105
            Top             =   2865
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   5385
            TabIndex        =   81
            Top             =   2850
            Width           =   315
         End
         Begin VB.Label labPorcT 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   2
            Left            =   4455
            TabIndex        =   80
            Top             =   2880
            Width           =   870
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   5385
            TabIndex        =   75
            Top             =   2475
            Width           =   315
         End
         Begin VB.Label labPorcM 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   2
            Left            =   4440
            TabIndex        =   74
            Top             =   2490
            Width           =   870
         End
         Begin VB.Label lab1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tot Rubro - T. Tarde/Noche"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   480
            TabIndex        =   67
            Top             =   2820
            Width           =   2220
         End
         Begin VB.Label lab1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tot Rubro - T. Mañana"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   480
            TabIndex        =   63
            Top             =   2430
            Width           =   2220
         End
      End
      Begin TabDlg.SSTab tabEzeB 
         Height          =   3540
         Left            =   -74295
         TabIndex        =   16
         Top             =   495
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   6244
         _Version        =   327680
         TabOrientation  =   1
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
         TabPicture(0)   =   "frmConsultaConcurso.frx":525D
         Tab(0).ControlCount=   10
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lab1(2)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lab1(4)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "labPorcM(1)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label9(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "labPorcT(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label9(5)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "labPorcN(1)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtM(1)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtT(1)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "sprEzeB(0)"
         Tab(0).Control(9).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmConsultaConcurso.frx":5279
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprEzeB(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "General"
         TabPicture(2)   =   "frmConsultaConcurso.frx":5295
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprRes(1)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2910
            Index           =   1
            Left            =   -74460
            OleObjectBlob   =   "frmConsultaConcurso.frx":52B1
            TabIndex        =   28
            Top             =   180
            Width           =   6135
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2205
            Index           =   0
            Left            =   405
            OleObjectBlob   =   "frmConsultaConcurso.frx":57E8
            TabIndex        =   27
            Top             =   135
            Width           =   6345
         End
         Begin FPSpread.vaSpread sprRes 
            Height          =   2985
            Index           =   1
            Left            =   -74880
            OleObjectBlob   =   "frmConsultaConcurso.frx":5D1F
            TabIndex        =   54
            Top             =   165
            Width           =   8340
         End
         Begin VB.TextBox txtT 
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
            Height          =   315
            Index           =   1
            Left            =   2715
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   2850
            Width           =   1560
         End
         Begin VB.TextBox txtM 
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
            Height          =   315
            Index           =   1
            Left            =   2730
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   2430
            Width           =   1560
         End
         Begin VB.Label labPorcN 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   1
            Left            =   5940
            TabIndex        =   104
            Top             =   2850
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   5325
            TabIndex        =   79
            Top             =   2820
            Width           =   315
         End
         Begin VB.Label labPorcT 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   1
            Left            =   4335
            TabIndex        =   78
            Top             =   2850
            Width           =   885
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   5340
            TabIndex        =   73
            Top             =   2460
            Width           =   315
         End
         Begin VB.Label labPorcM 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   1
            Left            =   4335
            TabIndex        =   72
            Top             =   2460
            Width           =   915
         End
         Begin VB.Label lab1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tot Rubro - T. Tarde/Noche"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   375
            TabIndex        =   65
            Top             =   2835
            Width           =   2220
         End
         Begin VB.Label lab1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tot Rubro - T. Mañana"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   375
            TabIndex        =   61
            Top             =   2445
            Width           =   2205
         End
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INTERIOR"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   870
         TabIndex        =   98
         Top             =   3225
         Width           =   2355
      End
      Begin VB.Label labGr 
         Alignment       =   1  'Right Justify
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   3615
         TabIndex        =   96
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -68910
         TabIndex        =   77
         Top             =   3315
         Width           =   315
      End
      Begin VB.Label labImp 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         Height          =   390
         Left            =   3660
         TabIndex        =   50
         Top             =   4050
         Width           =   2265
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A cada persona le corresponde $ "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   315
         TabIndex        =   49
         Top             =   4035
         Width           =   8115
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   855
         TabIndex        =   48
         Top             =   3645
         Width           =   4185
      End
      Begin VB.Line Line1 
         X1              =   3630
         X2              =   8085
         Y1              =   3585
         Y2              =   3585
      End
      Begin VB.Label Label6 
         Caption         =   "Cantidad de Personas"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   5040
         TabIndex        =   41
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label labGr 
         Alignment       =   1  'Right Justify
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   40
         Top             =   2745
         Width           =   1575
      End
      Begin VB.Label labGr 
         Alignment       =   1  'Right Justify
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   3585
         TabIndex        =   39
         Top             =   2505
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label labGr 
         Alignment       =   1  'Right Justify
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   3585
         TabIndex        =   38
         Top             =   2130
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AEROPARQUE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   870
         TabIndex        =   37
         Top             =   2700
         Width           =   2340
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INTERNACIONAL  'B'"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   36
         Top             =   2475
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EZEIZA"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   35
         Top             =   2100
         Width           =   2340
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GANADORES"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   315
         TabIndex        =   34
         Top             =   1485
         Width           =   8115
      End
   End
   Begin VB.Frame frdatos 
      Height          =   1170
      Left            =   15
      TabIndex        =   0
      Top             =   -45
      Width           =   8385
      Begin VB.ComboBox cboConcurso 
         Height          =   315
         Left            =   4230
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   225
         Width           =   2475
      End
      Begin VB.TextBox txtConc 
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   6720
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   21
         Top             =   210
         Width           =   720
      End
      Begin VB.ComboBox cboProv 
         Height          =   315
         Left            =   4215
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   705
         Width           =   2475
      End
      Begin VB.TextBox txtProv 
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   6720
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   19
         Top             =   705
         Width           =   735
      End
      Begin VB.ListBox LstConcurso 
         Height          =   255
         Left            =   8025
         TabIndex        =   18
         Top             =   300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox lstProv 
         Height          =   255
         Left            =   7995
         TabIndex        =   17
         Top             =   705
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   2535
         Picture         =   "frmConsultaConcurso.frx":6582
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   690
         Width           =   375
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2550
         Picture         =   "frmConsultaConcurso.frx":66F4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1380
         TabIndex        =   9
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
         Left            =   1380
         TabIndex        =   10
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
      Begin VB.Label labRubro 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7680
         TabIndex        =   59
         Top             =   465
         Width           =   540
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Concurso"
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
         Left            =   3135
         TabIndex        =   24
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
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
         Left            =   3135
         TabIndex        =   23
         Top             =   720
         Width           =   1050
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
         Left            =   180
         TabIndex        =   12
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
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmConsultaConcurso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset

Dim RsDataEzeB  As Recordset
Dim RsDataAep  As Recordset
Dim RsDataEzeA  As Recordset

Dim FVDesde As String
Dim FVHasta As String

Dim TTUnid As Long
Dim TTDist  As Single

'Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas
                       'con porcentajes
'Dim UsoPorIntB As Boolean
'Dim UsoPorAero As Boolean
'Dim UsoPorTotal As Boolean

Dim col_Prod As Collection

Dim valorOrden(1 To 3) As Single
Dim grupoOrden(1 To 3) As Integer

Dim Porc As Single
Private Function L_Armarcondicion(Esp As String) As String
Dim Cond As String, condLoc As String
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant, i

fechaDesde = mskFDesde.FormattedText
fechaHasta = mskFHasta.FormattedText

Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

Cond = Cond & " and Cod_prod IN ( " & L_TodosLosProd & " )"


L_Armarcondicion = Cond

End Function

Private Sub L_BuscarElRubro()
Dim SQL As String
Dim rs As Recordset

SQL = "SELECT cod_rubr FROM estadis.Concurso_h WHERE id_concurso = '" & txtConc.Text & "' "

If Aplicacion.ObtenerRsDAO(SQL, rs) Then
    If Aplicacion.CantReg(rs) > 0 Then
        labRubro.caption = rs!cod_rubr
    End If
    Aplicacion.CerrarDAO rs
End If


End Sub

Private Function L_CalcularImporteDist() As Single
Dim SQL As String
Dim rs As Recordset

SQL = " SELECT ESTADIS.Calculo_plus( '" & txtConc.Text & "','" & txtProv.Text & "', " _
& func_ToDate(mskFDesde.FormattedText) & ", " & func_ToDate(mskFHasta.FormattedText) & " )" _
& " as plus FROM Dual "

If Aplicacion.ObtenerDyDAO(SQL, rs) Then
    L_CalcularImporteDist = TTDist + rs!plus
    
    Aplicacion.CerrarDAO rs
End If

End Function

Private Function L_CalcularImporteDistas() As Single
Dim SQL As String
Dim rs As Recordset
Dim prov As String
Dim unidades As Long

SQL = " SELECT cod_prov, limite_rango, plus " _
& " FROM concurso_plus WHERE id_concurso = '" & txtConc.Text & "' " _
& " ORDER BY cod_prov, limite_rango desc "

If Aplicacion.ObtenerRsDAO(SQL, rs) Then
    prov = rs!cod_prov
    
    Do While Not rs.EOF
        
    Loop
    Aplicacion.CerrarDAO rs
End If

End Function

Private Sub L_CalculoOrdenGanadores(VA As Single, VB As Single, VC As Single)

If Val(VA) > Val(VB) And Val(VA) > Val(VC) And Val(VB) >= Val(VC) Then
'(A,B,C)
    valorOrden(1) = VA
    valorOrden(2) = VB
    valorOrden(3) = VC
    grupoOrden(1) = 3
    grupoOrden(2) = 5
    grupoOrden(3) = 7

ElseIf Val(VA) > Val(VB) And Val(VA) > Val(VC) And Val(VC) >= Val(VB) Then
'(A,C,B)
    valorOrden(1) = VA
    valorOrden(2) = VC
    valorOrden(3) = VB
    grupoOrden(1) = 3
    grupoOrden(2) = 7
    grupoOrden(3) = 5
ElseIf Val(VB) > Val(VC) And Val(VA) >= Val(VC) Then
'(B,A,C)
    valorOrden(1) = VB
    valorOrden(2) = VA
    valorOrden(3) = VC
    grupoOrden(1) = 5
    grupoOrden(2) = 3
    grupoOrden(3) = 7
ElseIf Val(VB) > Val(VC) And Val(VC) >= Val(VA) Then
'(B,C,A)
    valorOrden(1) = VB
    valorOrden(2) = VC
    valorOrden(3) = VA
    grupoOrden(1) = 5
    grupoOrden(2) = 7
    grupoOrden(3) = 3
ElseIf Val(VA) >= Val(VB) Then
'(C,A,B)
    valorOrden(1) = VC
    valorOrden(2) = VA
    valorOrden(3) = VB
    grupoOrden(1) = 7
    grupoOrden(2) = 3
    grupoOrden(3) = 5
ElseIf Val(VB) >= Val(VA) Then
'(C,B,A)
    valorOrden(1) = VC
    valorOrden(2) = VB
    valorOrden(3) = VA
    grupoOrden(1) = 7
    grupoOrden(2) = 5
    grupoOrden(3) = 3
End If

End Sub

Private Sub L_LimpiarGrillas()
Dim i
sprRes(0).MaxRows = 2
sprRes(1).MaxRows = 2
sprRes(2).MaxRows = 2
For i = 0 To 1
    sprEzeA(i).MaxRows = 0
    sprEzeB(i).MaxRows = 0
    sprAep(i).MaxRows = 0
    SprInt(i).MaxRows = 0
Next
TTDist = 0
TTUnid = 0
txtCant.Text = ""
txtImp.Text = ""

For i = 0 To 3
    labGr(i).caption = ""
    txtGr(i).Text = ""
    labPorcM(i).caption = ""
    labPorcT(i).caption = ""
    txtM(i).Text = ""
    txtT(i).Text = ""
Next

txtPers(0).Text = ""
txtPers(1).Text = ""
txtPers(2).Text = ""
txtPers(3).Text = ""
txtPers(4).Text = ""

labImp.caption = ""

End Sub

Private Sub L_LlenarTotRubro()
Dim SQL As String
Dim rs As Recordset
Dim valor As Variant

If labRubro.caption <> "DIS" Then
    SQL = "SELECT cod_sdep,DECODE(gRUPO_VENTA,'A','A','B') turno, sum(importe) Imp " _
    & " FROM estadis.venta_plg " _
    & " WHERE FCH_TICKET Between " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText) _
    & " And cod_sdep in('AEP') And cod_rubr = '" & labRubro.caption & "'" & " GROUP BY cod_sdep,DECODE(gRUPO_VENTA,'A','A','B') " _
    & " UNION " _
    & " SELECT 'INT' cod_sdep,DECODE(cod_sdep,'CORD','A','B')  turno, sum(importe) Imp " _
    & " FROM estadis.venta_plg " _
    & " WHERE FCH_TICKET Between " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText) _
    & " And cod_depn = 'INT'  and cod_sdep in ('MEND','CORD')And cod_rubr = '" & labRubro.caption & "'" & " GROUP BY cod_sdep,DECODE(cod_sdep,'CORD','A','B') " _
    & "  UNION " _
    & " SELECT  'INTA' cod_sdep,GRUPO_VENTA  turno, sum(importe) Imp " _
    & " FROM estadis.venta_plg " _
    & " WHERE FCH_TICKET Between " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText) _
    & " And cod_depn = 'EZE'  And cod_rubr = '" & labRubro.caption & "'" & " GROUP BY grupo_venta  "
Else
    SQL = "SELECT cod_sdep,DECODE(gRUPO_VENTA,'A','A','B') turno, sum(importe) Imp " _
    & " FROM estadis.venta_plg " _
    & " WHERE FCH_TICKET Between " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText) _
    & " And cod_depn = 'AEP' GROUP BY cod_sdep,DECODE(gRUPO_VENTA,'A','A','B') " _
    & " UNION " _
    & " SELECT 'INT' cod_sdep,DECODE(cod_sdep,'CORD','A','B') turno, sum(importe) Imp " _
    & " FROM estadis.venta_plg " _
    & " WHERE FCH_TICKET Between " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText) _
    & " And cod_depn = 'INT' and cod_sdep in ('MEND','CORD') GROUP BY cod_sdep,DECODE(cod_sdep,'CORD','A','B')  " _
    & "  UNION " _
    & " SELECT  'INTA' cod_sdep,GRUPO_VENTA  turno, sum(importe) Imp " _
    & " WHERE FCH_TICKET Between " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText) _
    & " And cod_depn = 'EZE' GROUP BY gRUPO_VENTA " _

End If

If Aplicacion.ObtenerRsDAO(SQL, rs) Then
    
    Do While Not rs.EOF
        Select Case rs!cod_sdep
            Case "AEP"
                'Select Case rs!turno
                '    Case "A"
                '        txtM(2).Text = Format(rs!imp, "#,#00.0")
                '        SprAep(0).GetText 2, SprAep(0).MaxRows, valor
                '        If rs!imp > 0 Then
                '            labPorcM(2).caption = Format(Val(valor) / rs!imp * 100, "#0.00")
                '        End If
                '
                '    Case "B"
                '        txtT(2).Text = Format(rs!imp, "#,#00.0")
                '        SprAep(0).GetText 3, SprAep(0).MaxRows, valor
                '        If rs!imp > 0 Then
                '            labPorcT(2).caption = Format(Val(valor) / rs!imp * 100, "#0.00")
                '        End If
                '
                'End Select
            Case "INTA"
                Select Case rs!turno
                    Case "A"
                        txtM(0).Text = Format(rs!imp, "#,#00.0")
                        sprEzeA(0).GetText 2, sprEzeA(0).MaxRows, valor
                        If rs!imp > 0 Then
                            labPorcM(0).caption = Format(Val(valor) / rs!imp * 100, "#0.00")
                        End If
                    Case "B"
                        txtT(0).Text = Format(rs!imp, "#,#00.0")
                        sprEzeA(0).GetText 3, sprEzeA(0).MaxRows, valor
                        If rs!imp > 0 Then
                            labPorcT(0).caption = Format(Val(valor) / rs!imp * 100, "#0.00")
                        End If
                    Case "C"
                        txtN(0).Text = Format(rs!imp, "#,#00.0")
                        sprEzeA(0).GetText 4, sprEzeA(0).MaxRows, valor
                        If rs!imp > 0 Then
                            labPorcN(0).caption = Format(Val(valor) / rs!imp * 100, "#0.00")
                        End If
                        
                End Select
            Case "INTB"
                Select Case rs!turno
                    Case "A"
                        txtM(1).Text = Format(rs!imp, "#,#00.0")
                        sprEzeB(0).GetText 2, sprEzeB(0).MaxRows, valor
                        If rs!imp > 0 Then
                            labPorcM(1).caption = Format(Val(valor) / rs!imp * 100, "#0.00")
                        End If
                    Case "B"
                        txtT(1).Text = Format(rs!imp, "#,#00.0")
                        sprEzeB(0).GetText 3, sprEzeB(0).MaxRows, valor
                        If rs!imp > 0 Then
                            labPorcT(1).caption = Format(Val(valor) / rs!imp * 100, "#0.00")
                        End If
                End Select
            Case "INT"
                Select Case rs!turno
                    Case "A"
                        txtM(3).Text = Format(rs!imp, "#,#00.0")
                        SprInt(0).GetText 2, SprInt(0).MaxRows, valor
                        If rs!imp > 0 Then
                            labPorcM(3).caption = Format(Val(valor) / rs!imp * 100, "#0.00")
                        End If
                        
                    Case "B"
                        txtT(3).Text = Format(rs!imp, "#,#00.0")
                        SprInt(0).GetText 3, SprInt(0).MaxRows, valor
                        If rs!imp > 0 Then
                            labPorcT(3).caption = Format(Val(valor) / rs!imp * 100, "#0.00")
                        End If
                
                End Select
                        
        End Select
        rs.MoveNext
    Loop
    Aplicacion.CerrarDAO rs
End If

'Calculo especial para AEP
SQL = " Select grupo turno,  sum(porc) IMP From estadis.porcentaje_dgrl " _
      & " WHERE anio = " & Year(mskFDesde.FormattedText) _
      & " And Mes = " & Month(mskFDesde.FormattedText) & " And Tipo_porc = 'I' and nivel = 0 " _
      & " and depn = 'AEP' and sdep = 'AEP'" _
      & " and dia Between " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText) _
      & " GROUP BY grupo "
If Aplicacion.ObtenerRsDAO(SQL, rs) Then
    
    Do While Not rs.EOF
       Select Case rs!turno
           Case "A"
              txtM(2).Text = Format(rs!imp, "#,#00.0")
              sprAep(0).GetText 2, sprAep(0).MaxRows, valor
              If rs!imp > 0 Then
                  labPorcM(2).caption = Format(Val(valor) / rs!imp * 100, "#0.00")
              End If
       
           Case "B"
              txtT(2).Text = Format(rs!imp, "#,#00.0")
               sprAep(0).GetText 3, sprAep(0).MaxRows, valor
               If rs!imp > 0 Then
                   labPorcT(2).caption = Format(Val(valor) / rs!imp * 100, "#0.00")
               End If
        
        End Select
        rs.MoveNext
    Loop
    
End If
End Sub

Private Function L_Mayor25porc(Esp As Integer, col As Integer) As Boolean
Dim i
Dim valor As Variant
Dim Result As Boolean

Result = True
For i = 3 To sprRes(Esp).MaxRows
    sprRes(Esp).GetText col, i, valor
    
    If Val(valor) < 25 Then
        Result = False
        Exit For
    End If
Next

L_Mayor25porc = Result

End Function

Private Sub L_Porcentages(ByRef sprPorc As control, sprDato As control, Orientacion As String)
Dim i As Integer, j As Integer
Dim valor As Variant
Dim tot As Variant, Prod As Variant
Dim Result As Double

On Error GoTo ErrPorc:

sprPorc.MaxRows = 0

Select Case Orientacion
    Case "FILAS"
        For i = 1 To sprDato.MaxRows
            sprDato.GetText sprDato.MaxCols, i, tot
            sprDato.GetText 1, i, Prod
            Result = 0
            sprPorc.MaxRows = i
            
            sprPorc.SetText 1, i, str(Prod)
            sprPorc.SetText sprPorc.MaxCols, i, "100"
            
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
            sprDato.GetText 1, i, Prod
            Result = 0
            sprPorc.MaxRows = i
        
            sprPorc.SetText 1, i, Prod
            
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

Private Sub L_Refrescar()
Dim SQL As String
Dim sqlX As String
Dim rs As Recordset
Dim i

On Error GoTo ErrGLC:

frmConsultaConcurso.caption = Aplicacion.SeteoProceso(frmConsultaConcurso.caption)
 
SQL = " SELECT cod_depn,cod_sdep,cod_local,DECODE(gRUPO_VENTA,'A','A','B') GRUPO_VENTA, " _
& " Sum (Importe) Importe, " _
& " Sum (cantidad) Unidades, " _
& " Sum(cantidad * cd.paga) Paga " _
& " FROM ESTADIS.CONCURSO_D CD, ESTADIS.VENTA_PLG V " _
& " WHERE CD.ID_CONCURSO = '" & txtConc.Text & "'" _
& " And V.FCH_TICKET Between " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText) _
& " AND V.COD_PROD = CD.COD_PROD AND V.COD_SDEP IN ('AEP') "
'& " AND CH.ID_CONCURSO=CD.ID_CONCURSO "
If txtProv.Text <> "" Then
    SQL = SQL & " And cd.cod_prov = '" & txtProv.Text & "' "
End If
SQL = SQL & " group by cod_depn,cod_sdep,cod_local,DECODE(gRUPO_VENTA,'A','A','B') " _
& "  "

SQL = SQL & " UNION "


SQL = SQL & " SELECT cod_depn,cod_depn COD_SDEP,cod_local,DECODE(cod_sdep,'CORD','A','B')  GRUPO_VENTA, " _
& " Sum (Importe) Importe, " _
& " Sum (cantidad) Unidades, " _
& " Sum(cantidad * cd.paga) Paga " _
& " FROM ESTADIS.CONCURSO_D CD, ESTADIS.VENTA_PLG V " _
& " WHERE CD.ID_CONCURSO = '" & txtConc.Text & "'" _
& " And V.FCH_TICKET Between " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText) _
& " AND V.COD_PROD = CD.COD_PROD AND V.COD_DEPN = 'INT' and V.COD_SDEP IN ('CORD','MEND') "
'& " AND CH.ID_CONCURSO=CD.ID_CONCURSO "
If txtProv.Text <> "" Then
    SQL = SQL & " And cd.cod_prov = '" & txtProv.Text & "' "
End If
SQL = SQL & " group by cod_depn,cod_depN,cod_local,DECODE(cod_sdep,'CORD','A','B') " _
& "  "


SQL = SQL & " UNION "


SQL = SQL & " SELECT cod_depn,'INTA' cod_sdep,cod_local,grupo_venta, " _
& " Sum (Importe) Importe, " _
& " Sum (cantidad) Unidades, " _
& " Sum(cantidad * cd.paga) Paga " _
& " FROM ESTADIS.CONCURSO_D CD, ESTADIS.VENTA_PLG V " _
& " WHERE CD.ID_CONCURSO = '" & txtConc.Text & "'" _
& " And V.FCH_TICKET Between " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText) _
& " AND V.COD_PROD = CD.COD_PROD AND V.COD_DEPN = 'EZE' "
'& " AND CH.ID_CONCURSO=CD.ID_CONCURSO "
If txtProv.Text <> "" Then
    SQL = SQL & " And cd.cod_prov = '" & txtProv.Text & "' "
End If
SQL = SQL & " group by cod_depn,cod_sdep,cod_local,grupo_venta "

If Aplicacion.ObtenerRsDAO(SQL, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        L_Llenargrilla
        L_LlenarTotRubro
        resumen
    End If

    Aplicacion.CerrarDAO RsData
End If

ErrGLC:
    frmConsultaConcurso.caption = Aplicacion.SeteoFin
    Exit Sub
    

End Sub

Private Sub L_SeteosGral()
        
    frdatos.Enabled = False
    botEjecutar(0).Enabled = False
    tabEspigon.Enabled = True
    frA.Refresh

End Sub

Private Sub L_TodosLosConcursos()
Dim SQL As String
Dim rs As Recordset
Dim Conc As String
Dim totGr_A As Single, totGr_B As Single, totGr_C As Single
Dim totGrB_A As Single, totGrB_B As Single, totGrB_C As Single
Dim totGrE_A As Single, totGrE_B As Single, totGrE_C As Single

SQL = "select H.descrip,C.cod_sdep,C.grupo,sum(C.importe) TotGr " _
& " from estadis.resumen_concurso C, estadis.concurso_h H " _
& " WHERE Fecha Between " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText) _
& " And H.id_concurso = C.id_concurso And activo = 'S' " _
& " group by H.descrip,C.cod_sdep,C.grupo "

If Aplicacion.ObtenerRsDAO(SQL, rs) Then
    
    Do While Not rs.EOF
        Conc = rs!Descrip
        totGr_A = 0
        totGr_B = 0
        totGr_C = 0
        totGrB_A = 0
        totGrB_B = 0
        totGrB_C = 0
        totGrE_A = 0
        totGrE_B = 0
        totGrE_C = 0
        
        sprRes(0).MaxRows = sprRes(0).MaxRows + 1
        sprRes(0).SetText 1, sprRes(0).MaxRows, Trim(rs!Descrip)
        sprRes(1).MaxRows = sprRes(1).MaxRows + 1
        sprRes(1).SetText 1, sprRes(1).MaxRows, Trim(rs!Descrip)
        sprRes(2).MaxRows = sprRes(2).MaxRows + 1
        sprRes(2).SetText 1, sprRes(2).MaxRows, Trim(rs!Descrip)
        
        Do While Conc = rs!Descrip
            Select Case rs!cod_sdep
                Case "AEP"
                    Select Case rs!Grupo
                        Case "A"
                            sprRes(2).SetText 2, sprRes(2).MaxRows, str(rs!totgr)
                            totGrE_A = rs!totgr
                        Case "B"
                            sprRes(2).SetText 4, sprRes(2).MaxRows, str(rs!totgr)
                            totGrE_B = rs!totgr
                        Case "C"
                            sprRes(2).SetText 6, sprRes(2).MaxRows, str(rs!totgr)
                            totGrE_C = rs!totgr
                    End Select
                    
                Case "INTA"
                    Select Case rs!Grupo
                        Case "A"
                            sprRes(0).SetText 2, sprRes(0).MaxRows, str(rs!totgr)
                            totGr_A = rs!totgr
                        Case "B"
                            sprRes(0).SetText 4, sprRes(0).MaxRows, str(rs!totgr)
                            totGr_B = rs!totgr
                        Case "C"
                            sprRes(0).SetText 6, sprRes(0).MaxRows, str(rs!totgr)
                            totGr_C = rs!totgr
                    End Select
                Case "INTB"
                    Select Case rs!Grupo
                        Case "A"
                            sprRes(1).SetText 2, sprRes(1).MaxRows, str(rs!totgr)
                            totGrB_A = rs!totgr
                        Case "B"
                            sprRes(1).SetText 4, sprRes(1).MaxRows, str(rs!totgr)
                            totGrB_B = rs!totgr
                        Case "C"
                            sprRes(1).SetText 6, sprRes(1).MaxRows, str(rs!totgr)
                            totGrB_C = rs!totgr
                    End Select
                
            End Select
            
            rs.MoveNext
            If rs.EOF Then
                Exit Do
            End If
        Loop
        sprRes(0).SetText 3, sprRes(0).MaxRows, Format$(100 * totGr_A / (totGr_A + totGr_B + totGr_C), "##.##")
        sprRes(0).SetText 5, sprRes(0).MaxRows, Format$(100 * totGr_B / (totGr_A + totGr_B + totGr_C), "##.##")
        sprRes(0).SetText 7, sprRes(0).MaxRows, Format$(100 * totGr_C / (totGr_A + totGr_B + totGr_C), "##.##")
        sprRes(1).SetText 3, sprRes(1).MaxRows, Format$(100 * totGrB_A / (totGrB_A + totGrB_B + totGrB_C), "##.##")
        sprRes(1).SetText 5, sprRes(1).MaxRows, Format$(100 * totGrB_B / (totGrB_A + totGrB_B + totGrB_C), "##.##")
        sprRes(1).SetText 7, sprRes(1).MaxRows, Format$(100 * totGrB_C / (totGrB_A + totGrB_B + totGrB_C), "##.##")
        sprRes(2).SetText 3, sprRes(2).MaxRows, Format$(100 * totGrE_A / (totGrE_A + totGrE_B + totGrE_C), "##.##")
        sprRes(2).SetText 5, sprRes(2).MaxRows, Format$(100 * totGrE_B / (totGrE_A + totGrE_B + totGrE_C), "##.##")
        sprRes(2).SetText 7, sprRes(2).MaxRows, Format$(100 * totGrE_C / (totGrE_A + totGrE_B + totGrE_C), "##.##")
        
    Loop
    Aplicacion.CerrarDAO rs
End If

End Sub

Private Function L_TodosLosProd() As String
Dim cl As CLProdMon
Dim Prod As String

Prod = ""

For Each cl In col_Prod
    Prod = Prod & ", " & cl.Cod_prod
Next

If Prod = "" Then
    L_TodosLosProd = "-1"
Else
    L_TodosLosProd = Right(Prod, Len(Prod) - 2)
End If


End Function

Private Sub PorSiVuelveACambiar()
'Inter A-----------------------------------------
'sprEzeA(0).GetText 2, sprEzeA(0).MaxRows, valorGA
'sprEzeA(0).GetText 3, sprEzeA(0).MaxRows, valorGB
'SprEzeA(0).GetText 4, SprEzeA(0).MaxRows, valorGC

'valorGA = Val(labPorcM(0).caption)
'valorGB = Val(labPorcT(0).caption)

'valorOrden(1) = 0
'valorOrden(2) = 0
'valorOrden(3) = 0

'L_CalculoOrdenGanadores Val(valorGA), Val(valorGB), 0

'For i = 1 To 3
'    If L_Mayor25porc(0, grupoOrden(i)) Then
'        txtGr(0).Text = Chr(64 + Int(grupoOrden(1) / 2))
'        labGr(0).caption = L_Transf(txtGr(0).Text)
'        Exit For
'    End If
'Next

'Inter B -----------------------------------------
'SprEzeB(0).GetText 2, SprEzeB(0).MaxRows, valorGA
'SprEzeB(0).GetText 3, SprEzeB(0).MaxRows, valorGB
'SprEzeB(0).GetText 4, SprEzeB(0).MaxRows, valorGC

'valorGA = Val(labPorcM(1).caption)
'valorGB = Val(labPorcT(1).caption)

'valorOrden(1) = 0
'valorOrden(2) = 0
'valorOrden(3) = 0

'L_CalculoOrdenGanadores Val(valorGA), Val(valorGB), 0

'For i = 1 To 3
'    If L_Mayor25porc(1, grupoOrden(i)) Then
'        txtGr(1).Text = Chr(64 + Int(grupoOrden(1) / 2))
'        labGr(1).caption = L_Transf(txtGr(1).Text)
'        Exit For
'    End If
'Next


'Aeroparque -----------------------------------
'SprAep(0).GetText 2, SprAep(0).MaxRows, valorGA
'SprAep(0).GetText 3, SprAep(0).MaxRows, valorGB
'SprAep(0).GetText 4, SprAep(0).MaxRows, valorGC

'valorGA = Val(labPorcM(2).caption)
'valorGB = Val(labPorcT(2).caption)

'valorOrden(1) = 0
'valorOrden(2) = 0
'valorOrden(3) = 0

'L_CalculoOrdenGanadores Val(valorGA), Val(valorGB), 0

'For i = 1 To 3
'    If L_Mayor25porc(2, grupoOrden(i)) Then
'        txtGr(2).Text = Chr(64 + Int(grupoOrden(1) / 2))
'        labGr(2).caption = L_Transf(txtGr(2).Text)
'        Exit For
'    End If
'Next
'Interior -----------------------------------
'valorGA = Val(labPorcM(3).caption)
'valorGB = Val(labPorcT(3).caption)

'valorOrden(1) = 0
'valorOrden(2) = 0
'valorOrden(3) = 0

'L_CalculoOrdenGanadores Val(valorGA), Val(valorGB), 0

End Sub

Private Sub resumen()
Dim valorGA As Variant
Dim valorGB As Variant
Dim valorGC As Variant
Dim SQL As String
Dim rs As Recordset

Dim i As Integer

'L_TodosLosConcursos

txtCant.Text = TTUnid
txtImp.Text = L_CalcularImporteDist


For i = 0 To 3
    valorGA = Val(labPorcM(i).caption)
    valorGB = Val(labPorcT(i).caption)
    valorGC = Val(labPorcN(i).caption)
    
    valorOrden(1) = 0
    valorOrden(2) = 0
    valorOrden(3) = 0

    L_CalculoOrdenGanadores Val(valorGA), Val(valorGB), Val(valorGC)

    txtGr(i).Text = Chr(64 + Int(grupoOrden(1) / 2))
    labGr(i).caption = L_Transf(txtGr(i).Text, i)

Next


'& "  And( (cod_Sdep = 'INTB' And DECODE(gRUPO,'A','A','B') = '" & txtGr(1).Text & "' ) or"
'Calculo de cantidad de personal
SQL = " SELECT cod_sdep, DECODE(gRUPO,'A','A','B'), count(legajo) CantPers " _
& " FROM estadis.persona_equipos PE, estadis.equipos E " _
& " where E.cod_equipo = PE.cod_equipo " _
& "  And (cod_depn = 'AEP' And DECODE(gRUPO,'A','A','B') = '" & txtGr(2).Text & "' )  " _
& " Group By cod_sdep,DECODE(gRUPO,'A','A','B') " _
& " UNION " _
& " SELECT 'INT' cod_sdep, decode(cod_sdep,'CORD','A','B') grupo, count(legajo) CantPers " _
& " FROM estadis.persona_equipos PE, estadis.equipos E " _
& " where E.cod_equipo = PE.cod_equipo " _
& " And cod_Sdep in ('CORD','MEND') And decode(cod_sdep,'CORD','A','B') = '" & txtGr(3).Text & "'  " _
& " Group By cod_sdep,decode(cod_sdep,'CORD','A','B') " _
& " UNION " _
& " SELECT 'INTA' cod_sdep, grupo, count(legajo) CantPers " _
& " FROM estadis.persona_equipos PE, estadis.equipos E " _
& " where E.cod_equipo = PE.cod_equipo " _
& " And cod_depn = 'EZE' And GRUPO = '" & txtGr(0).Text & "'  " _
& " Group By grupo "

'& " And(  (cod_Sdep = 'INTA' And DECODE(gRUPO,'A','A','B') = '" & txtGr(0).Text & "' ) " _

txtPers(0).Text = 0
txtPers(1).Text = 0
txtPers(2).Text = 0
If Aplicacion.ObtenerDyDAO(SQL, rs) Then
    
    Do While Not rs.EOF
        Select Case rs!cod_sdep
            Case "INTA"
                txtPers(0).Text = rs!cantpers
            Case "INTB"
                txtPers(1).Text = rs!cantpers
            Case "AEP"
                txtPers(2).Text = rs!cantpers
            Case "INT"
                txtPers(3).Text = rs!cantpers
        End Select
        rs.MoveNext
    Loop
    Aplicacion.CerrarDAO rs
End If
txtPers(4).Text = Val(txtPers(0).Text) + Val(txtPers(1).Text) + Val(txtPers(2).Text) + Val(txtPers(3).Text)

If Val(txtPers(4).Text) > 0 Then
labImp.caption = Format$((Val(txtImp.Text) / Val(txtPers(4).Text)), "#,##0.00")
End If

End Sub

Private Function L_Transf(letra As String, i As Integer) As String

If i = 0 Then
 Select Case letra
 Case "A"
    L_Transf = "Mañana"
 Case "B"
    L_Transf = "Tarde"
 Case "C"
    L_Transf = "Noche"
 End Select

Else
 
 Select Case letra
 Case "A"
    L_Transf = "Mañana"
 Case "B"
    L_Transf = "Tarde/Noche"
 End Select
End If
End Function



Private Sub botBases_Click()

If txtConc.Text <> "" Then
    FrmAdmConcurso.ConsultaBases txtConc.Text
End If

End Sub

Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        Select Case Aplicacion.Perfil
            Case "TODOS"
                If cboConcurso.Text <> "" And mskFDesde.Text <> "" And mskFHasta.Text <> "" Then
                    L_Refrescar
                    L_SeteosGral
                Else
                    MsgBox "Faltan algunos datos", vbOKOnly + vbExclamation, "ATENCION"
                End If
        End Select
    Case 1
        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        'tabEspigon.Enabled = False
        L_LimpiarGrillas
    Case 2
        Unload Me
End Select
End Sub

Private Sub botExcelAep_Click()
Select Case TabAep.Tab
    Case 0
    Case 1
End Select

End Sub

Private Sub botExel_Click()


Porc = frmPorc.TomarPorc

If Porc <> -1 Then
    L_TratarExcel
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


Private Sub cboConcurso_Click()
Dim SQL As String
Dim rs As Recordset

txtConc.Text = LstConcurso.List(cboConcurso.ListIndex)

If txtConc.Text <> "" Then
    SQL = " SELECT fch_vdesde,fch_vhasta FROM estadis.concurso_h " _
    & " WHERE id_concurso = '" & txtConc.Text & "' "
    Call Aplicacion.ObtenerDyDAO(SQL, rs)
    FVDesde = Format$(rs!fch_vdesde, FTOFECHA)
    FVHasta = Format$(rs!fch_vhasta, FTOFECHA)
    
    mskFDesde.Text = FVDesde
    mskFHasta.Text = FVHasta

End If
SQL = " SELECT P.cod_prov,P.descrip FROM estadis.concurso_PROV C,baires.proveedor P " _
& " WHERE P.cod_prov=C.cod_prov And C.id_concurso = '" & txtConc.Text & "' " _
& " UNION " _
& " SELECT CP.cod_prov, 'DISCONTINUADOS' AS Descrip " _
& " FROM estadis.concurso_prov CP " _
& " WHERE  CP.cod_prov ='N' " _
& " and CP.id_concurso = '" & txtConc.Text & "' "


FuncCbos_LlenarCboLst cboProv, lstProv, SQL

L_BuscarElRubro

End Sub


Private Sub cboConcurso_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    cboConcurso.ListIndex = -1
    txtProv.Text = ""
    FVDesde = ""
    FVHasta = ""

    mskFDesde.Text = ""
    mskFHasta.Text = ""
ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub cboProv_Click()
txtProv.Text = lstProv.List(cboProv.ListIndex)
End Sub


Private Sub cboProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    cboProv.ListIndex = -1
ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabEspigon
End Sub

Private Sub L_TratarExcel()
Dim AppExcel As Object
Dim rango As String
Dim col As Integer
Dim fila As Integer
Dim nombre As String
Dim SQL As String
Dim rs As Recordset

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

frmConsultaConcurso.caption = Aplicacion.SeteoProceso(frmConsultaConcurso.caption)

If nombre <> "" Then
    '& "     OR(cod_Sdep = 'INTB' And DECODE(gRUPO,'A','A','B')  = '" & txtGr(1).Text & "' )"
    SQL = " SELECT legajo,time " _
    & " FROM estadis.persona_equipos PE, estadis.equipos E " _
    & " where E.cod_equipo = PE.cod_equipo " _
    & " And(  (cod_depn = 'EZE' And gRUPO = '" & txtGr(0).Text & "' ) " _
    & "     OR(cod_depn = 'AEP' And DECODE(gRUPO,'A','A','B') = '" & txtGr(2).Text & "' ) " _
    & "     OR(cod_depn = 'INT' And DECODE(gRUPO,'A','A','B') = '" & txtGr(3).Text & "') ) "
        
    If Aplicacion.ObtenerRsDAO(SQL, rs) Then
         If rs.RecordCount > 0 Then
            Set AppExcel = CreateObject("excel.sheet")
            fila = 1
            Do While Not rs.EOF
    
                'AppExcel.application.Visible = True
                
                Exl.Exl_PonerValor AppExcel, fila, 1, str(rs!Legajo)
                Exl.Exl_PonerValor AppExcel, fila, 2, "112"
                If rs!Time = FULL Then
                    Exl.Exl_PonerValor AppExcel, fila, 4, Format((Val(labImp.caption) * Porc) / 100, "###.00")
                Else
                    Exl.Exl_PonerValor AppExcel, fila, 4, Format(((Val(labImp.caption) / 2) * Porc) / 100, "###.00")
                End If
                
                fila = fila + 1
                rs.MoveNext
            Loop
            AppExcel.SaveAs nombre & ".xls"
            Set AppExcel = Nothing
         Else
            MsgBox "No hay personal asignado a los grupos", vbExclamation + vbOKOnly, "ATENCION"
         End If
         Aplicacion.CerrarDAO rs
    End If
End If

ErrorExl:

    frmConsultaConcurso.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub Form_Load()
Dim i
Dim SQL As String

Me.Left = 50
Me.Top = 100
Me.Height = 6500
Me.Width = 9300


SQL = " SELECT id_concurso,descrip FROM estadis.concurso_H  WHERE activo = 'S' "
SQL = SQL & " ORDER BY id_concurso "

FuncCbos_LlenarCboLst cboConcurso, LstConcurso, SQL

'UsoPorIntA = False
'UsoPorIntB = False
'UsoPorTotal = False
'UsoPorAero = False

Set col_Prod = New Collection
L_LimpiarGrillas

TabEzeA.TabVisible(2) = False
TabEzeB.TabVisible(2) = False
TabAep.TabVisible(2) = False

tabEspigon.TabVisible(2) = False

frmPrincipal.lstForms.AddItem "frmProd"

End Sub

Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmProd"
End Sub


Private Sub mskFDesde_LostFocus()
    
If mskFDesde.Text <> "" Then
If cboConcurso.Text <> "" Then
    If IsDate(mskFDesde.FormattedText) Then
        If Not (CDate(mskFDesde.FormattedText) >= CDate(FVDesde) _
        And CDate(mskFDesde.FormattedText) <= CDate(FVHasta)) Then
            mskFDesde.Text = FVDesde
        End If
     Else
        mskFDesde.Text = FVDesde
     End If
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)

    If mskFHasta.Text <> "" Then
        If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Or Year(mskFHasta.FormattedText) <> Year(mskFDesde.FormattedText) Then
            mskFHasta.Text = mskFDesde.Text
        End If
    End If

Else
    MsgBox "Primero debe elegir un Concurso ", vbOKOnly + vbInformation, "ATENCION"
    mskFDesde.Text = ""
End If
Else
mskFHasta.Text = ""
End If

'        mskFDesde.Text = Date
'    End If

End Sub


Private Sub mskFHasta_LostFocus()

If mskFHasta.Text <> "" Then
If cboConcurso.Text <> "" Then
    If IsDate(mskFHasta.FormattedText) Then
        If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) _
        Or CDate(mskFHasta.FormattedText) > CDate(FVHasta) Then
                mskFHasta.Text = FVHasta
        End If
    Else
        mskFHasta.Text = FVHasta
    End If
    mskFHasta.Text = Format$(mskFHasta.FormattedText, FTOFECHA)
Else
    MsgBox "Primero debe elegir un Concurso ", vbOKOnly + vbInformation, "ATENCION"
    mskFHasta.Text = ""
End If
End If


'        mskFHasta.Text = mskFDesde.Text

End Sub


Private Sub L_Llenargrilla()
Dim codProd As String
Dim i As Integer
Dim cLocal As String

For i = 0 To 1
    sprEzeA(i).MaxRows = 0
    sprEzeB(i).MaxRows = 0
    sprAep(i).MaxRows = 0
Next

Do While Not RsData.EOF
    
    Select Case RsData!cod_sdep
        Case "INTA"
            sprEzeA(0).MaxRows = sprEzeA(0).MaxRows + 1
            sprEzeA(1).MaxRows = sprEzeA(1).MaxRows + 1
            cLocal = RsData!cod_local
            sprEzeA(0).SetText 1, sprEzeA(0).MaxRows, cLocal
            sprEzeA(1).SetText 1, sprEzeA(1).MaxRows, cLocal
            Do While cLocal = RsData!cod_local
                TTDist = TTDist + RsData!paga
                TTUnid = TTUnid + RsData!unidades
                
                Select Case RsData.grupo_venta
                    Case "A"
                        sprEzeA(0).SetText 2, sprEzeA(0).MaxRows, str(RsData!Importe)
                        sprEzeA(1).SetText 2, sprEzeA(1).MaxRows, str(RsData!unidades)
                    Case "B"
                        sprEzeA(0).SetText 3, sprEzeA(0).MaxRows, str(RsData!Importe)
                        sprEzeA(1).SetText 3, sprEzeA(1).MaxRows, str(RsData!unidades)
                    Case "C"
                        sprEzeA(0).SetText 4, sprEzeA(0).MaxRows, str(RsData!Importe)
                        sprEzeA(1).SetText 4, sprEzeA(1).MaxRows, str(RsData!unidades)
                End Select
                
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If
            Loop
        Case "INTB"
            sprEzeB(0).MaxRows = sprEzeB(0).MaxRows + 1
            sprEzeB(1).MaxRows = sprEzeB(1).MaxRows + 1
            cLocal = RsData!cod_local
            sprEzeB(0).SetText 1, sprEzeB(0).MaxRows, cLocal
            sprEzeB(1).SetText 1, sprEzeB(1).MaxRows, cLocal
            Do While cLocal = RsData!cod_local
                TTDist = TTDist + RsData!paga
                TTUnid = TTUnid + RsData!unidades
                
                Select Case RsData.grupo_venta
                    Case "A"
                        sprEzeB(0).SetText 2, sprEzeB(0).MaxRows, str(RsData!Importe)
                        sprEzeB(1).SetText 2, sprEzeB(1).MaxRows, str(RsData!unidades)
                    Case "B"
                        sprEzeB(0).SetText 3, sprEzeB(0).MaxRows, str(RsData!Importe)
                        sprEzeB(1).SetText 3, sprEzeB(1).MaxRows, str(RsData!unidades)
                    Case "C"
                        sprEzeB(0).SetText 4, sprEzeB(0).MaxRows, str(RsData!Importe)
                        sprEzeB(1).SetText 4, sprEzeB(1).MaxRows, str(RsData!unidades)
                End Select
                
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If
            Loop
        
        Case "AEP"
            sprAep(0).MaxRows = sprAep(0).MaxRows + 1
            sprAep(1).MaxRows = sprAep(1).MaxRows + 1
            cLocal = RsData!cod_local
            sprAep(0).SetText 1, sprAep(0).MaxRows, cLocal
            sprAep(1).SetText 1, sprAep(1).MaxRows, cLocal
            Do While cLocal = RsData!cod_local
                TTDist = TTDist + RsData!paga
                TTUnid = TTUnid + RsData!unidades
                Select Case RsData.grupo_venta
                    Case "A"
                        sprAep(0).SetText 2, sprAep(0).MaxRows, str(RsData!Importe)
                        sprAep(1).SetText 2, sprAep(1).MaxRows, str(RsData!unidades)
                    Case "B"
                        sprAep(0).SetText 3, sprAep(0).MaxRows, str(RsData!Importe)
                        sprAep(1).SetText 3, sprAep(1).MaxRows, str(RsData!unidades)
                    Case "C"
                        sprAep(0).SetText 4, sprAep(0).MaxRows, str(RsData!Importe)
                        sprAep(1).SetText 4, sprAep(1).MaxRows, str(RsData!unidades)
                End Select
                
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If
            Loop
        Case "INT"
            SprInt(0).MaxRows = SprInt(0).MaxRows + 1
            SprInt(1).MaxRows = SprInt(1).MaxRows + 1
            cLocal = RsData!cod_local
            SprInt(0).SetText 1, SprInt(0).MaxRows, cLocal
            SprInt(1).SetText 1, SprInt(1).MaxRows, cLocal
            Do While cLocal = RsData!cod_local
                TTDist = TTDist + RsData!paga
                TTUnid = TTUnid + RsData!unidades
                Select Case RsData.grupo_venta
                    Case "A"
                        SprInt(0).SetText 2, SprInt(0).MaxRows, str(RsData!Importe)
                        SprInt(1).SetText 2, SprInt(1).MaxRows, str(RsData!unidades)
                    Case "B"
                        SprInt(0).SetText 3, SprInt(0).MaxRows, str(RsData!Importe)
                        SprInt(1).SetText 3, SprInt(1).MaxRows, str(RsData!unidades)
                    Case "C"
                        SprInt(0).SetText 4, SprInt(0).MaxRows, str(RsData!Importe)
                        SprInt(1).SetText 4, SprInt(1).MaxRows, str(RsData!unidades)
                End Select
                
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If
            Loop
         Case Else
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If
         
    End Select
    
        
Loop

For i = 0 To 1
    Spread_TotalesGrillas sprEzeA(i), sprEzeA(i).MaxCols - 1, 2
    Spread_TotalesGrillas sprEzeB(i), sprEzeB(i).MaxCols - 1, 2
    Spread_TotalesGrillas sprAep(i), sprAep(i).MaxCols - 1, 2
    Spread_TotalesGrillas SprInt(i), SprInt(i).MaxCols - 1, 2
Next

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

