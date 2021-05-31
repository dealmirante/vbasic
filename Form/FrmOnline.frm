VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form FrmOnline 
   Caption         =   "Ventas Online por Rubro"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   11175
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   615
      Top             =   5370
   End
   Begin VB.Frame Frame1 
      Height          =   930
      Left            =   2190
      TabIndex        =   18
      Top             =   -45
      Width           =   7350
      Begin ComctlLib.Slider sldTmp 
         Height          =   210
         Left            =   4830
         TabIndex        =   24
         Top             =   390
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   370
         _Version        =   327680
         OLEDropMode     =   1
         LargeChange     =   1
         Min             =   1
         SelStart        =   5
         Value           =   5
      End
      Begin VB.CommandButton botSemaforo 
         Height          =   615
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   180
         Width           =   660
      End
      Begin VB.TextBox txtTmpC 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3705
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   540
         Width           =   1080
      End
      Begin VB.TextBox txtTmpA 
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
         Height          =   285
         Left            =   3705
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label labFecha 
         Alignment       =   2  'Center
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
         Height          =   525
         Left            =   75
         TabIndex        =   34
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label Label5 
         Caption         =   "minutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5445
         TabIndex        =   27
         Top             =   645
         Width           =   735
      End
      Begin VB.Label labMin 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "5"
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
         Height          =   270
         Left            =   4995
         TabIndex        =   26
         Top             =   600
         Width           =   330
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "consulta cada..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4905
         TabIndex        =   25
         Top             =   180
         Width           =   1440
      End
      Begin VB.Label Label2 
         Caption         =   "Tiempo Consulta"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2370
         TabIndex        =   20
         Top             =   585
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Hora Actual "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2385
         TabIndex        =   19
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame3 
      Height          =   945
      Left            =   120
      TabIndex        =   9
      Top             =   -60
      Width           =   1995
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
         Height          =   495
         Index           =   2
         Left            =   1155
         Picture         =   "FrmOnline.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   225
         Width           =   615
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
         Height          =   495
         Index           =   0
         Left            =   330
         Picture         =   "FrmOnline.frx":0822
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   210
         Width           =   615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4590
      Left            =   135
      TabIndex        =   0
      Top             =   900
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   8096
      _Version        =   327680
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "VENTAS POR RUBRO"
      TabPicture(0)   =   "FrmOnline.frx":0924
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TabEspigon"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "VENTAS POR LOCAL"
      TabPicture(1)   =   "FrmOnline.frx":0940
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TabEspigonVtas"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "% MONEDAS"
      TabPicture(2)   =   "FrmOnline.frx":095C
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab7"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "COMPARATIVOS"
      TabPicture(3)   =   "FrmOnline.frx":0978
      Tab(3).ControlCount=   7
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label10"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label9"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label8"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label7"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label4"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label11"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame2"
      Tab(3).Control(6).Enabled=   0   'False
      Begin TabDlg.SSTab SSTab7 
         Height          =   3675
         Left            =   -73725
         TabIndex        =   95
         Top             =   495
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   6482
         _Version        =   327680
         TabHeight       =   520
         ForeColor       =   255
         TabCaption(0)   =   "TOTAL"
         TabPicture(0)   =   "FrmOnline.frx":0994
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprPago(2)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "LLEADAS"
         TabPicture(1)   =   "FrmOnline.frx":09B0
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprPago(0)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "SALIDAS"
         TabPicture(2)   =   "FrmOnline.frx":09CC
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprPago(1)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprPago 
            Height          =   2955
            Index           =   2
            Left            =   405
            OleObjectBlob   =   "FrmOnline.frx":09E8
            TabIndex        =   96
            Top             =   450
            Width           =   6765
         End
         Begin FPSpread.vaSpread sprPago 
            Height          =   2955
            Index           =   0
            Left            =   -74595
            OleObjectBlob   =   "FrmOnline.frx":15ED
            TabIndex        =   97
            Top             =   450
            Width           =   6750
         End
         Begin FPSpread.vaSpread sprPago 
            Height          =   2955
            Index           =   1
            Left            =   -74595
            OleObjectBlob   =   "FrmOnline.frx":21F1
            TabIndex        =   98
            Top             =   450
            Width           =   6750
         End
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   3000
         Left            =   120
         TabIndex        =   50
         Top             =   915
         Width           =   10665
         Begin MSMask.MaskEdBox mskAntHora 
            Height          =   345
            Index           =   0
            Left            =   990
            TabIndex        =   51
            Top             =   840
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskAntHora 
            Height          =   345
            Index           =   1
            Left            =   2490
            TabIndex        =   52
            Top             =   840
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskAntHora 
            Height          =   345
            Index           =   3
            Left            =   5790
            TabIndex        =   53
            Top             =   840
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskActual 
            Height          =   345
            Index           =   0
            Left            =   990
            TabIndex        =   54
            Top             =   225
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskActual 
            Height          =   360
            Index           =   1
            Left            =   2490
            TabIndex        =   55
            Top             =   210
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskActual 
            Height          =   345
            Index           =   3
            Left            =   5790
            TabIndex        =   56
            Top             =   225
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskEstim 
            Height          =   345
            Index           =   0
            Left            =   990
            TabIndex        =   57
            Top             =   2235
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskEstim 
            Height          =   345
            Index           =   1
            Left            =   2490
            TabIndex        =   58
            Top             =   2235
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskEstim 
            Height          =   345
            Index           =   3
            Left            =   5805
            TabIndex        =   59
            Top             =   2235
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskActual 
            Height          =   345
            Index           =   4
            Left            =   9090
            TabIndex        =   60
            Top             =   225
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskAntHora 
            Height          =   345
            Index           =   4
            Left            =   9105
            TabIndex        =   61
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskEstim 
            Height          =   345
            Index           =   4
            Left            =   9090
            TabIndex        =   62
            Top             =   2235
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskAnt 
            Height          =   345
            Index           =   0
            Left            =   990
            TabIndex        =   63
            Top             =   1530
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskAnt 
            Height          =   345
            Index           =   1
            Left            =   2490
            TabIndex        =   64
            Top             =   1530
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskAnt 
            Height          =   345
            Index           =   3
            Left            =   5820
            TabIndex        =   65
            Top             =   1530
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskAnt 
            Height          =   345
            Index           =   4
            Left            =   9090
            TabIndex        =   66
            Top             =   1530
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskAntHora 
            Height          =   345
            Index           =   2
            Left            =   4140
            TabIndex        =   67
            Top             =   840
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskActual 
            Height          =   360
            Index           =   2
            Left            =   4125
            TabIndex        =   68
            Top             =   225
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   635
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskEstim 
            Height          =   345
            Index           =   2
            Left            =   4125
            TabIndex        =   69
            Top             =   2235
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskAnt 
            Height          =   345
            Index           =   2
            Left            =   4140
            TabIndex        =   70
            Top             =   1530
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskActual 
            Height          =   345
            Index           =   5
            Left            =   7485
            TabIndex        =   109
            Top             =   225
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskAntHora 
            Height          =   345
            Index           =   5
            Left            =   7515
            TabIndex        =   110
            Top             =   855
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskAnt 
            Height          =   345
            Index           =   5
            Left            =   7485
            TabIndex        =   112
            Top             =   1545
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskEstim 
            Height          =   345
            Index           =   5
            Left            =   7485
            TabIndex        =   114
            Top             =   2265
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.0"
            PromptChar      =   " "
         End
         Begin VB.Label labEstim 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   8445
            TabIndex        =   115
            Top             =   2310
            Width           =   510
         End
         Begin VB.Label labAnt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   8460
            TabIndex        =   113
            Top             =   1620
            Width           =   495
         End
         Begin VB.Label labAntHora 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   8475
            TabIndex        =   111
            Top             =   930
            Width           =   495
         End
         Begin VB.Label Label6 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   2
            Left            =   90
            TabIndex        =   89
            Top             =   1530
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   " Estimado"
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
            Left            =   30
            TabIndex        =   88
            Top             =   2250
            Width           =   1170
         End
         Begin VB.Label Label6 
            Caption         =   " Actual "
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
            Left            =   30
            TabIndex        =   87
            Top             =   225
            Width           =   1170
         End
         Begin VB.Label Label6 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   3
            Left            =   120
            TabIndex        =   86
            Top             =   825
            Width           =   960
         End
         Begin VB.Label labAntHora 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   1920
            TabIndex        =   85
            Top             =   885
            Width           =   465
         End
         Begin VB.Label labAntHora 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   3510
            TabIndex        =   84
            Top             =   930
            Width           =   540
         End
         Begin VB.Label labAntHora 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   5250
            TabIndex        =   83
            Top             =   885
            Width           =   480
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   840
            X2              =   10455
            Y1              =   1395
            Y2              =   1395
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   825
            X2              =   10455
            Y1              =   705
            Y2              =   705
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   840
            X2              =   10410
            Y1              =   2070
            Y2              =   2070
         End
         Begin VB.Label labEstim 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   1905
            TabIndex        =   82
            Top             =   2250
            Width           =   540
         End
         Begin VB.Label labEstim 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   3495
            TabIndex        =   81
            Top             =   2280
            Width           =   525
         End
         Begin VB.Label labEstim 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   5250
            TabIndex        =   80
            Top             =   2265
            Width           =   510
         End
         Begin VB.Label labAntHora 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   6885
            TabIndex        =   79
            Top             =   900
            Width           =   555
         End
         Begin VB.Label labEstim 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   6915
            TabIndex        =   78
            Top             =   2265
            Width           =   525
         End
         Begin VB.Label labAnt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   6900
            TabIndex        =   77
            Top             =   1590
            Width           =   525
         End
         Begin VB.Label labAnt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   5265
            TabIndex        =   76
            Top             =   1575
            Width           =   510
         End
         Begin VB.Label labAnt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   3480
            TabIndex        =   75
            Top             =   1620
            Width           =   570
         End
         Begin VB.Label labAnt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   1905
            TabIndex        =   74
            Top             =   1575
            Width           =   540
         End
         Begin VB.Label labAnt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   10065
            TabIndex        =   73
            Top             =   1605
            Width           =   495
         End
         Begin VB.Label labEstim 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   10050
            TabIndex        =   72
            Top             =   2280
            Width           =   510
         End
         Begin VB.Label labAntHora 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   10065
            TabIndex        =   71
            Top             =   915
            Width           =   495
         End
      End
      Begin TabDlg.SSTab TabEspigonVtas 
         Height          =   3960
         Left            =   -74535
         TabIndex        =   17
         Top             =   450
         Width           =   9780
         _ExtentX        =   17251
         _ExtentY        =   6985
         _Version        =   327680
         Tabs            =   6
         Tab             =   5
         TabsPerRow      =   6
         TabHeight       =   520
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
         TabPicture(0)   =   "FrmOnline.frx":2DF5
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "SprTotalBsAs"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "botExcel(7)"
         Tab(0).Control(1).Enabled=   -1  'True
         TabCaption(1)   =   "EZE - A Llegadas"
         TabPicture(1)   =   "FrmOnline.frx":2E11
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "botExcel(4)"
         Tab(1).Control(0).Enabled=   -1  'True
         Tab(1).Control(1)=   "SprVtasEzeA"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "EZE - A Salidas"
         TabPicture(2)   =   "FrmOnline.frx":2E2D
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SprVtasEzeAS"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "botExcel(5)"
         Tab(2).Control(1).Enabled=   -1  'True
         TabCaption(3)   =   "EZE - B"
         TabPicture(3)   =   "FrmOnline.frx":2E49
         Tab(3).ControlCount=   2
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "botExcel(6)"
         Tab(3).Control(0).Enabled=   -1  'True
         Tab(3).Control(1)=   "SprVtasEzeB"
         Tab(3).Control(1).Enabled=   0   'False
         TabCaption(4)   =   "AEP"
         TabPicture(4)   =   "FrmOnline.frx":2E65
         Tab(4).ControlCount=   2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "botExcel(9)"
         Tab(4).Control(0).Enabled=   -1  'True
         Tab(4).Control(1)=   "SprVtasAep"
         Tab(4).Control(1).Enabled=   0   'False
         TabCaption(5)   =   "INTERIOR"
         TabPicture(5)   =   "FrmOnline.frx":2E81
         Tab(5).ControlCount=   1
         Tab(5).ControlEnabled=   -1  'True
         Tab(5).Control(0)=   "sprVtasInt"
         Tab(5).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprVtasInt 
            Height          =   2955
            Left            =   495
            OleObjectBlob   =   "FrmOnline.frx":2E9D
            TabIndex        =   107
            Top             =   675
            Width           =   6645
         End
         Begin FPSpread.vaSpread SprTotalBsAs 
            Height          =   2955
            Left            =   -74850
            OleObjectBlob   =   "FrmOnline.frx":32F1
            TabIndex        =   32
            Top             =   645
            Width           =   7110
         End
         Begin FPSpread.vaSpread SprVtasEzeA 
            Height          =   2955
            Left            =   -74550
            OleObjectBlob   =   "FrmOnline.frx":3745
            TabIndex        =   30
            Top             =   570
            Width           =   6645
         End
         Begin FPSpread.vaSpread SprVtasAep 
            Height          =   2955
            Left            =   -74505
            OleObjectBlob   =   "FrmOnline.frx":3B99
            TabIndex        =   47
            Top             =   675
            Width           =   6645
         End
         Begin FPSpread.vaSpread SprVtasEzeB 
            Height          =   2955
            Left            =   -74730
            OleObjectBlob   =   "FrmOnline.frx":3FED
            TabIndex        =   48
            Top             =   615
            Width           =   6645
         End
         Begin FPSpread.vaSpread SprVtasEzeAS 
            Height          =   2955
            Left            =   -74580
            OleObjectBlob   =   "FrmOnline.frx":4441
            TabIndex        =   49
            Top             =   615
            Width           =   6645
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   9
            Left            =   -67470
            Picture         =   "FrmOnline.frx":4895
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   3285
            Width           =   765
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   7
            Left            =   -67590
            Picture         =   "FrmOnline.frx":4E27
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   3135
            Width           =   765
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   4
            Left            =   -67740
            Picture         =   "FrmOnline.frx":53B9
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   3015
            Width           =   765
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   5
            Left            =   -67665
            Picture         =   "FrmOnline.frx":594B
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   3060
            Width           =   765
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   6
            Left            =   -67710
            Picture         =   "FrmOnline.frx":5EDD
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   3090
            Width           =   765
         End
      End
      Begin TabDlg.SSTab TabEspigon 
         Height          =   3960
         Left            =   -74595
         TabIndex        =   1
         Top             =   435
         Width           =   9645
         _ExtentX        =   17013
         _ExtentY        =   6985
         _Version        =   327680
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
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
         TabPicture(0)   =   "FrmOnline.frx":646F
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSTab3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "botExcel(0)"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "EZE - A Llegadas"
         TabPicture(1)   =   "FrmOnline.frx":648B
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "botExcel(1)"
         Tab(1).Control(0).Enabled=   -1  'True
         Tab(1).Control(1)=   "SSTab4"
         Tab(1).Control(1).Enabled=   0   'False
         TabCaption(2)   =   "EZE - A Salidas"
         TabPicture(2)   =   "FrmOnline.frx":64A7
         Tab(2).ControlCount=   2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "botExcel(2)"
         Tab(2).Control(0).Enabled=   -1  'True
         Tab(2).Control(1)=   "SSTab5"
         Tab(2).Control(1).Enabled=   0   'False
         TabCaption(3)   =   "EZE - B"
         TabPicture(3)   =   "FrmOnline.frx":64C3
         Tab(3).ControlCount=   2
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SSTab6"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "botExcel(3)"
         Tab(3).Control(1).Enabled=   -1  'True
         TabCaption(4)   =   "AEP"
         TabPicture(4)   =   "FrmOnline.frx":64DF
         Tab(4).ControlCount=   2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "botExcel(8)"
         Tab(4).Control(0).Enabled=   -1  'True
         Tab(4).Control(1)=   "SSTab2"
         Tab(4).Control(1).Enabled=   0   'False
         TabCaption(5)   =   "INTERIOR"
         TabPicture(5)   =   "FrmOnline.frx":64FB
         Tab(5).ControlCount=   1
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "SSTab8"
         Tab(5).Control(0).Enabled=   0   'False
         Begin TabDlg.SSTab SSTab8 
            Height          =   3285
            Left            =   -74625
            TabIndex        =   103
            Top             =   480
            Width           =   6030
            _ExtentX        =   10636
            _ExtentY        =   5794
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
            TabCaption(0)   =   "Importe"
            TabPicture(0)   =   "FrmOnline.frx":6517
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprInt(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "FrmOnline.frx":6533
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprInt(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Promedios"
            TabPicture(2)   =   "FrmOnline.frx":654F
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprInt(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprInt 
               Height          =   2700
               Index           =   0
               Left            =   300
               OleObjectBlob   =   "FrmOnline.frx":656B
               TabIndex        =   104
               Top             =   120
               Width           =   5400
            End
            Begin FPSpread.vaSpread SprInt 
               Height          =   2700
               Index           =   1
               Left            =   -74715
               OleObjectBlob   =   "FrmOnline.frx":69D6
               TabIndex        =   105
               Top             =   120
               Width           =   5400
            End
            Begin FPSpread.vaSpread SprInt 
               Height          =   2700
               Index           =   2
               Left            =   -74715
               OleObjectBlob   =   "FrmOnline.frx":6E41
               TabIndex        =   106
               Top             =   120
               Width           =   5400
            End
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   8
            Left            =   -67455
            Picture         =   "FrmOnline.frx":72AC
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   3240
            Width           =   765
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   3
            Left            =   -67485
            Picture         =   "FrmOnline.frx":783E
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   3300
            Width           =   765
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   2
            Left            =   -66570
            Picture         =   "FrmOnline.frx":7DD0
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   3300
            Width           =   765
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   1
            Left            =   -67500
            Picture         =   "FrmOnline.frx":8362
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   3285
            Width           =   765
         End
         Begin VB.CommandButton botExcel 
            Caption         =   "Excel"
            Height          =   500
            Index           =   0
            Left            =   8580
            Picture         =   "FrmOnline.frx":88F4
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3300
            Width           =   765
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   3300
            Left            =   300
            TabIndex        =   2
            Top             =   495
            Width           =   8130
            _ExtentX        =   14340
            _ExtentY        =   5821
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
            TabCaption(0)   =   "Importe"
            TabPicture(0)   =   "FrmOnline.frx":8E86
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprTotal(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "FrmOnline.frx":8EA2
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprTotal(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Promedio"
            TabPicture(2)   =   "FrmOnline.frx":8EBE
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprTotal(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprTotal 
               Height          =   2700
               Index           =   0
               Left            =   105
               OleObjectBlob   =   "FrmOnline.frx":8EDA
               TabIndex        =   6
               Top             =   135
               Width           =   7860
            End
            Begin FPSpread.vaSpread SprTotal 
               Height          =   2700
               Index           =   1
               Left            =   -74910
               OleObjectBlob   =   "FrmOnline.frx":93A1
               TabIndex        =   7
               Top             =   105
               Width           =   7905
            End
            Begin FPSpread.vaSpread SprTotal 
               Height          =   2700
               Index           =   2
               Left            =   -74895
               OleObjectBlob   =   "FrmOnline.frx":981F
               TabIndex        =   8
               Top             =   105
               Width           =   7800
            End
         End
         Begin TabDlg.SSTab SSTab4 
            Height          =   3300
            Left            =   -74700
            TabIndex        =   3
            Top             =   500
            Width           =   7000
            _ExtentX        =   12356
            _ExtentY        =   5821
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
            TabCaption(0)   =   "Importe"
            TabPicture(0)   =   "FrmOnline.frx":9CC9
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprEzeA(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "FrmOnline.frx":9CE5
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprEzeA(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Promedio"
            TabPicture(2)   =   "FrmOnline.frx":9D01
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprEzeA(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   2700
               Index           =   2
               Left            =   -74610
               OleObjectBlob   =   "FrmOnline.frx":9D1D
               TabIndex        =   102
               Top             =   135
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   2700
               Index           =   1
               Left            =   -74640
               OleObjectBlob   =   "FrmOnline.frx":A12D
               TabIndex        =   101
               Top             =   135
               Width           =   5505
            End
            Begin FPSpread.vaSpread SprEzeA 
               Height          =   2700
               Index           =   0
               Left            =   375
               OleObjectBlob   =   "FrmOnline.frx":A513
               TabIndex        =   16
               Top             =   135
               Width           =   5505
            End
         End
         Begin TabDlg.SSTab SSTab5 
            Height          =   3300
            Left            =   -74700
            TabIndex        =   4
            Top             =   495
            Width           =   7920
            _ExtentX        =   13970
            _ExtentY        =   5821
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
            TabCaption(0)   =   "Importe"
            TabPicture(0)   =   "FrmOnline.frx":A923
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprEzeAS(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "FrmOnline.frx":A93F
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprEzeAS(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Promedio"
            TabPicture(2)   =   "FrmOnline.frx":A95B
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprEzeAS(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprEzeAS 
               Height          =   2700
               Index           =   2
               Left            =   -74895
               OleObjectBlob   =   "FrmOnline.frx":A977
               TabIndex        =   100
               Top             =   105
               Width           =   6735
            End
            Begin FPSpread.vaSpread SprEzeAS 
               Height          =   2700
               Index           =   0
               Left            =   120
               OleObjectBlob   =   "FrmOnline.frx":AE36
               TabIndex        =   45
               Top             =   120
               Width           =   7650
            End
            Begin FPSpread.vaSpread SprEzeAS 
               Height          =   2700
               Index           =   1
               Left            =   -74865
               OleObjectBlob   =   "FrmOnline.frx":B2F5
               TabIndex        =   99
               Top             =   120
               Width           =   7635
            End
         End
         Begin TabDlg.SSTab SSTab6 
            Height          =   3300
            Left            =   -74700
            TabIndex        =   5
            Top             =   500
            Width           =   7000
            _ExtentX        =   12356
            _ExtentY        =   5821
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
            TabCaption(0)   =   "Importe"
            TabPicture(0)   =   "FrmOnline.frx":B76E
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprEzeB(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "FrmOnline.frx":B78A
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprEzeB(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Promedio"
            TabPicture(2)   =   "FrmOnline.frx":B7A6
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprEzeB(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   2700
               Index           =   2
               Left            =   -74685
               OleObjectBlob   =   "FrmOnline.frx":B7C2
               TabIndex        =   44
               Top             =   150
               Width           =   6405
            End
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   2700
               Index           =   1
               Left            =   -74700
               OleObjectBlob   =   "FrmOnline.frx":BC34
               TabIndex        =   43
               Top             =   135
               Width           =   6405
            End
            Begin FPSpread.vaSpread SprEzeB 
               Height          =   2700
               Index           =   0
               Left            =   285
               OleObjectBlob   =   "FrmOnline.frx":C07E
               TabIndex        =   42
               Top             =   150
               Width           =   6405
            End
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   3300
            Left            =   -74595
            TabIndex        =   35
            Top             =   450
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   5821
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
            TabCaption(0)   =   "Importe"
            TabPicture(0)   =   "FrmOnline.frx":C51C
            Tab(0).ControlCount=   1
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SprAep(0)"
            Tab(0).Control(0).Enabled=   0   'False
            TabCaption(1)   =   "Unidades"
            TabPicture(1)   =   "FrmOnline.frx":C538
            Tab(1).ControlCount=   1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SprAep(1)"
            Tab(1).Control(0).Enabled=   0   'False
            TabCaption(2)   =   "Promedio"
            TabPicture(2)   =   "FrmOnline.frx":C554
            Tab(2).ControlCount=   1
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SprAep(2)"
            Tab(2).Control(0).Enabled=   0   'False
            Begin FPSpread.vaSpread SprAep 
               Height          =   2700
               Index           =   2
               Left            =   -74190
               OleObjectBlob   =   "FrmOnline.frx":C570
               TabIndex        =   41
               Top             =   90
               Width           =   5400
            End
            Begin FPSpread.vaSpread SprAep 
               Height          =   2700
               Index           =   1
               Left            =   -74190
               OleObjectBlob   =   "FrmOnline.frx":C960
               TabIndex        =   40
               Top             =   90
               Width           =   5400
            End
            Begin FPSpread.vaSpread SprAep 
               Height          =   2700
               Index           =   5
               Left            =   -74295
               OleObjectBlob   =   "FrmOnline.frx":CD44
               TabIndex        =   38
               Top             =   105
               Width           =   5400
            End
            Begin FPSpread.vaSpread SprAep 
               Height          =   2700
               Index           =   4
               Left            =   -74295
               OleObjectBlob   =   "FrmOnline.frx":D134
               TabIndex        =   37
               Top             =   105
               Width           =   5400
            End
            Begin FPSpread.vaSpread SprAep 
               Height          =   2700
               Index           =   0
               Left            =   810
               OleObjectBlob   =   "FrmOnline.frx":D518
               TabIndex        =   36
               Top             =   90
               Width           =   5400
            End
         End
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Interior"
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
         Left            =   7590
         TabIndex        =   108
         Top             =   525
         Width           =   1530
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INT A Lleg"
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
         Left            =   4260
         TabIndex        =   94
         Top             =   540
         Width           =   1470
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aeroparque"
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
         Left            =   1110
         TabIndex        =   93
         Top             =   540
         Width           =   1365
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inter B"
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
         Left            =   5925
         TabIndex        =   92
         Top             =   540
         Width           =   1530
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
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
         Left            =   9210
         TabIndex        =   91
         Top             =   525
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INT A Sal"
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
         Left            =   2640
         TabIndex        =   90
         Top             =   540
         Width           =   1440
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9150
      Picture         =   "FrmOnline.frx":D934
      Top             =   1905
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   9135
      Picture         =   "FrmOnline.frx":DD76
      Top             =   1230
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FrmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rstotal  As Recordset
Dim RsData As Recordset
Dim RsPagos As Recordset
Dim RsLocal As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas con porcentajes
Dim UsoPorIntAS As Boolean
Dim UsoPorIntB As Boolean
Dim UsoPorAero As Boolean

Dim UsoPorInte As Boolean

Dim Actualizar As Boolean
Dim FchAnt As String
Private Sub L_DecoLocales()

Dim i As Integer, indDep As Integer
Dim dato As Variant
Dim x As Variant
Dim Sdep As String

Do While RsData!cod_rubr = "AAA"
   
     Select Case RsData!Cod_Sdep
          Case "INTB"
             SprVtasEzeB.MaxRows = SprVtasEzeB.MaxRows + 1
             SprVtasEzeB.SetText 1, SprVtasEzeB.MaxRows, Trim(RsData!cod_local & "-" & RsData!Cod_Sloc)
             SprVtasEzeB.SetText 2, SprVtasEzeB.MaxRows, str(RsData!Importe)
             SprVtasEzeB.SetText 3, SprVtasEzeB.MaxRows, Trim(RsData!cantidad)
             SprVtasEzeB.SetText 4, SprVtasEzeB.MaxRows, Trim(RsData!ticket)
          Case "INTAL"
             SprVtasEzeA.MaxRows = SprVtasEzeA.MaxRows + 1
             SprVtasEzeA.SetText 1, SprVtasEzeA.MaxRows, Trim(RsData!cod_local & "-" & RsData!Cod_Sloc)
             SprVtasEzeA.SetText 2, SprVtasEzeA.MaxRows, str(RsData!Importe)
             SprVtasEzeA.SetText 3, SprVtasEzeA.MaxRows, Trim(RsData!cantidad)
             SprVtasEzeA.SetText 4, SprVtasEzeA.MaxRows, Trim(RsData!ticket)
          Case "INTAS"
             SprVtasEzeAS.MaxRows = SprVtasEzeAS.MaxRows + 1
             SprVtasEzeAS.SetText 1, SprVtasEzeAS.MaxRows, Trim(RsData!cod_local & "-" & RsData!Cod_Sloc)
             SprVtasEzeAS.SetText 2, SprVtasEzeAS.MaxRows, str(RsData!Importe)
             SprVtasEzeAS.SetText 3, SprVtasEzeAS.MaxRows, Trim(RsData!cantidad)
             SprVtasEzeAS.SetText 4, SprVtasEzeAS.MaxRows, Trim(RsData!ticket)
          Case "AEP"
             SprVtasAep.MaxRows = SprVtasAep.MaxRows + 1
             SprVtasAep.SetText 1, SprVtasAep.MaxRows, Trim(RsData!cod_local & "-" & RsData!Cod_Sloc)
             SprVtasAep.SetText 2, SprVtasAep.MaxRows, str(RsData!Importe)
             SprVtasAep.SetText 3, SprVtasAep.MaxRows, Trim(RsData!cantidad)
             SprVtasAep.SetText 4, SprVtasAep.MaxRows, Trim(RsData!ticket)
          Case "BARI", "MEND", "CORD"
             sprVtasInt.MaxRows = sprVtasInt.MaxRows + 1
             sprVtasInt.SetText 1, sprVtasInt.MaxRows, Trim(RsData!cod_local & "-" & RsData!Cod_Sloc)
             sprVtasInt.SetText 2, sprVtasInt.MaxRows, str(RsData!Importe)
             sprVtasInt.SetText 3, sprVtasInt.MaxRows, Trim(RsData!cantidad)
             sprVtasInt.SetText 4, sprVtasInt.MaxRows, Trim(RsData!ticket)
     
     End Select

RsData.MoveNext

Loop

End Sub
Private Sub L_DecoEspigon()
Dim rubro
Dim i As Integer, indDep As Integer
Dim fila As Integer

Do While Not RsData.EOF
   rubro = RsData!cod_rubr
   If Not RsData!cod_rubr = "AAA" Then
        For i = 0 To 2
          sprEzeB(i).MaxRows = sprEzeB(i).MaxRows + 1
          sprEzeB(i).SetText 1, sprEzeB(i).MaxRows, Trim(RsData!cod_rubr)
          sprEzeA(i).MaxRows = sprEzeA(i).MaxRows + 1
          sprEzeA(i).SetText 1, sprEzeA(i).MaxRows, Trim(RsData!cod_rubr)
          sprEzeAS(i).MaxRows = sprEzeAS(i).MaxRows + 1
          sprEzeAS(i).SetText 1, sprEzeAS(i).MaxRows, Trim(RsData!cod_rubr)
          sprAEP(i).MaxRows = sprAEP(i).MaxRows + 1
          sprAEP(i).SetText 1, sprAEP(i).MaxRows, Trim(RsData!cod_rubr)
          sprInt(i).MaxRows = sprInt(i).MaxRows + 1
          sprInt(i).SetText 1, sprInt(i).MaxRows, Trim(RsData!cod_rubr)
          sprTotal(i).MaxRows = sprTotal(i).MaxRows + 1
          sprTotal(i).SetText 1, sprTotal(i).MaxRows, Trim(RsData!cod_rubr)
        Next i
      Else
         L_DecoLocales
   End If
     
    Do While rubro = RsData!cod_rubr And RsData!cod_rubr <> "AAA"
        
        If RsData!cantidad > 0 Then
        Select Case RsData!cod_local
            Case Is = "L01"
               If RsData!Cod_Sloc = 0 Then
                   sprEzeB(0).SetText 2, sprEzeB(0).MaxRows, str(RsData!Importe)
                   sprEzeB(1).SetText 2, sprEzeB(1).MaxRows, str(RsData!cantidad)
                    sprEzeB(2).SetText 2, sprEzeB(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               ElseIf RsData!Cod_Sloc = 1 Then
                   sprEzeB(0).SetText 3, sprEzeB(0).MaxRows, str(RsData!Importe)
                   sprEzeB(1).SetText 3, sprEzeB(1).MaxRows, str(RsData!cantidad)
                   sprEzeB(2).SetText 3, sprEzeB(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               ElseIf RsData!Cod_Sloc = 2 Then
                   sprEzeB(0).SetText 4, sprEzeB(0).MaxRows, str(RsData!Importe)
                   sprEzeB(1).SetText 4, sprEzeB(1).MaxRows, str(RsData!cantidad)
                   sprEzeB(2).SetText 4, sprEzeB(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               End If
            Case Is = "L02"
                   sprEzeB(0).SetText 5, sprEzeB(0).MaxRows, str(RsData!Importe)
                   sprEzeB(1).SetText 5, sprEzeB(1).MaxRows, str(RsData!cantidad)
                   sprEzeB(2).SetText 5, sprEzeB(2).MaxRows, str(RsData!Importe / RsData!cantidad)
            Case Is = "L03"
                   sprEzeB(0).SetText 6, sprEzeB(0).MaxRows, str(RsData!Importe)
                   sprEzeB(1).SetText 6, sprEzeB(1).MaxRows, str(RsData!cantidad)
                   sprEzeB(2).SetText 6, sprEzeB(2).MaxRows, str(RsData!Importe / RsData!cantidad)
            Case Is = "L21"
               If RsData!Cod_Sloc = 0 Then
                   sprEzeA(0).SetText 2, sprEzeA(0).MaxRows, str(RsData!Importe)
                   sprEzeA(1).SetText 2, sprEzeA(1).MaxRows, str(RsData!cantidad)
                   sprEzeA(2).SetText 2, sprEzeA(2).MaxRows, str(RsData!Importe / RsData!cantidad)
                 Else
                   sprEzeA(0).SetText 3, sprEzeA(0).MaxRows, str(RsData!Importe)
                   sprEzeA(1).SetText 3, sprEzeA(1).MaxRows, str(RsData!cantidad)
                   sprEzeA(2).SetText 3, sprEzeA(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               End If
            Case Is = "L06"
                   sprEzeA(0).SetText 4, sprEzeA(0).MaxRows, str(RsData!Importe)
                   sprEzeA(1).SetText 4, sprEzeA(1).MaxRows, str(RsData!cantidad)
                   sprEzeA(2).SetText 4, sprEzeA(2).MaxRows, str(RsData!Importe / RsData!cantidad)
            Case Is = "L22"
               If RsData!Cod_Sloc = 0 Then
                   sprEzeAS(0).SetText 2, sprEzeAS(0).MaxRows, str(RsData!Importe)
                   sprEzeAS(1).SetText 2, sprEzeAS(1).MaxRows, str(RsData!cantidad)
                   sprEzeAS(2).SetText 2, sprEzeAS(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               ElseIf RsData!Cod_Sloc = 1 Then
                   sprEzeAS(0).SetText 3, sprEzeAS(0).MaxRows, str(RsData!Importe)
                   sprEzeAS(1).SetText 3, sprEzeAS(1).MaxRows, str(RsData!cantidad)
                   sprEzeAS(2).SetText 3, sprEzeAS(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               ElseIf RsData!Cod_Sloc = 2 Then
                   sprEzeAS(0).SetText 4, sprEzeAS(0).MaxRows, str(RsData!Importe)
                   sprEzeAS(1).SetText 4, sprEzeAS(1).MaxRows, str(RsData!cantidad)
                   sprEzeAS(2).SetText 4, sprEzeAS(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               ElseIf RsData!Cod_Sloc = 3 Then
                   sprEzeAS(0).SetText 5, sprEzeAS(0).MaxRows, str(RsData!Importe)
                   sprEzeAS(1).SetText 5, sprEzeAS(1).MaxRows, str(RsData!cantidad)
                   sprEzeAS(2).SetText 5, sprEzeAS(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               ElseIf RsData!Cod_Sloc = 4 Then
                   sprEzeAS(0).SetText 6, sprEzeAS(0).MaxRows, str(RsData!Importe)
                   sprEzeAS(1).SetText 6, sprEzeAS(1).MaxRows, str(RsData!cantidad)
                   sprEzeAS(2).SetText 6, sprEzeAS(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               
               End If
            Case Is = "L08"
               If RsData!Cod_Sloc = 0 Then
                   'sprEzeAS(0).SetText 5, sprEzeAS(0).MaxRows, str(RsData!Importe)
                   'sprEzeAS(1).SetText 5, sprEzeAS(1).MaxRows, str(RsData!cantidad)
                   'sprEzeAS(2).SetText 5, sprEzeAS(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               ElseIf RsData!Cod_Sloc = 1 Then
                   'SprEzeAS(0).SetText 5, SprEzeAS(0).MaxRows, str(RsData!Importe)
                   'SprEzeAS(1).SetText 5, SprEzeAS(1).MaxRows, str(RsData!cantidad)
                   'SprEzeAS(2).SetText 5, SprEzeAS(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               ElseIf RsData!Cod_Sloc = 3 Then
                   'SprEzeAS(0).SetText 6, SprEzeAS(0).MaxRows, str(RsData!Importe)
                   'SprEzeAS(1).SetText 6, SprEzeAS(1).MaxRows, str(RsData!cantidad)
                   'SprEzeAS(2).SetText 6, SprEzeAS(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               End If
            Case Is = "L09"
                   sprAEP(0).SetText 2, sprAEP(0).MaxRows, str(RsData!Importe)
                   sprAEP(1).SetText 2, sprAEP(1).MaxRows, str(RsData!cantidad)
                   sprAEP(2).SetText 2, sprAEP(2).MaxRows, str(RsData!Importe / RsData!cantidad)
            Case Is = "L14"
               If RsData!Cod_Sloc = 0 Then
                   sprAEP(0).SetText 3, sprAEP(0).MaxRows, str(RsData!Importe)
                   sprAEP(1).SetText 3, sprAEP(1).MaxRows, str(RsData!cantidad)
                   sprAEP(2).SetText 3, sprAEP(2).MaxRows, str(RsData!Importe / RsData!cantidad)
                 Else
                   sprAEP(0).SetText 4, sprAEP(0).MaxRows, str(RsData!Importe)
                   sprAEP(1).SetText 4, sprAEP(1).MaxRows, str(RsData!cantidad)
                   sprAEP(2).SetText 4, sprAEP(2).MaxRows, str(RsData!Importe / RsData!cantidad)
               End If
            Case Is = "L18"
                   sprInt(0).SetText 2, sprInt(0).MaxRows, str(RsData!Importe)
                   sprInt(1).SetText 2, sprAEP(1).MaxRows, str(RsData!cantidad)
                   sprInt(2).SetText 2, sprAEP(2).MaxRows, str(RsData!Importe / RsData!cantidad)
            
            Case Is = "L"
                  Select Case RsData!Cod_Sdep
                    Case Is = "INTAL"
                        sprTotal(0).SetText 2, sprTotal(0).MaxRows, str(RsData!Importe)
                        sprTotal(1).SetText 2, sprTotal(1).MaxRows, str(RsData!cantidad)
                        sprTotal(2).SetText 2, sprTotal(2).MaxRows, str(RsData!Importe / RsData!cantidad)
                    Case Is = "INTAS"
                        sprTotal(0).SetText 3, sprTotal(0).MaxRows, str(RsData!Importe)
                        sprTotal(1).SetText 3, sprTotal(1).MaxRows, str(RsData!cantidad)
                        sprTotal(2).SetText 3, sprTotal(2).MaxRows, str(RsData!Importe / RsData!cantidad)
                    Case Is = "INTB"
                        sprTotal(0).SetText 4, sprTotal(0).MaxRows, str(RsData!Importe)
                        sprTotal(1).SetText 4, sprTotal(1).MaxRows, str(RsData!cantidad)
                        sprTotal(2).SetText 4, sprTotal(2).MaxRows, str(RsData!Importe / RsData!cantidad)
                    Case Is = "AEP"
                        sprTotal(0).SetText 5, sprTotal(0).MaxRows, str(RsData!Importe)
                        sprTotal(1).SetText 5, sprTotal(1).MaxRows, str(RsData!cantidad)
                        sprTotal(2).SetText 5, sprTotal(2).MaxRows, str(RsData!Importe / RsData!cantidad)
                    Case "BARI"
                        sprTotal(0).SetText 6, sprTotal(0).MaxRows, str(RsData!Importe)
                        sprTotal(1).SetText 6, sprTotal(1).MaxRows, str(RsData!cantidad)
                        sprTotal(2).SetText 6, sprTotal(2).MaxRows, str(RsData!Importe / RsData!cantidad)
                  
                  End Select
                  
        End Select
        End If
        RsData.MoveNext
        
        If RsData.EOF Then
            Exit Do
        End If
    Loop
Loop

End Sub
Private Sub L_DecoPagos()
Dim tipoPago As Integer
Dim tipoLocal As String, tipoMon As Integer, nacion As String
Dim moneda As Variant
Dim ind As Integer
Dim TotPagoSE As Single, TotPagoLE As Single, TotPagoSA As Single, TotPagoLA As Single
Dim PagoE As Single, PagoA As Single
    L_DescripMonedas
        
    ind = 2
    
    Do While Not RsPagos.EOF
    tipoPago = RsPagos!tipo_pago
        TotPagoSA = 0
        TotPagoLA = 0
        TotPagoSE = 0
        TotPagoLE = 0
        Do While tipoPago = RsPagos!tipo_pago And Val(moneda) <> 99
            tipoMon = RsPagos!tipo_moneda
            ind = ind + 1
            sprPago(0).GetText 1, ind, moneda
            PagoA = 0
            PagoE = 0
            Do While Val(moneda) <> 0 And Val(moneda) <> 99 And tipoPago = RsPagos!tipo_pago And tipoMon = RsPagos!tipo_moneda
                nacion = RsPagos!NAC
                If Val(moneda) = tipoMon Then
                    Select Case RsPagos!tipo_loc
                        Case "L"
                            Select Case RsPagos!NAC
                                Case "A"
                                    sprPago(0).SetText 3, ind, str(RsPagos!Importe)
                                    TotPagoLA = TotPagoLA + RsPagos!Importe
                                    PagoA = PagoA + RsPagos!Importe
                                Case "E"
                                    sprPago(0).SetText 6, ind, str(RsPagos!Importe)
                                    TotPagoLE = TotPagoLE + RsPagos!Importe
                                    PagoE = PagoE + RsPagos!Importe
                            End Select
                        Case "S"
                            Select Case RsPagos!NAC
                                Case "A"
                                    sprPago(1).SetText 3, ind, str(RsPagos!Importe)
                                    TotPagoSA = TotPagoSA + RsPagos!Importe
                                    PagoA = PagoA + RsPagos!Importe
                                Case "E"
                                    sprPago(1).SetText 6, ind, str(RsPagos!Importe)
                                    TotPagoSE = TotPagoSE + RsPagos!Importe
                                    PagoE = PagoE + RsPagos!Importe
                            End Select
                    
                    End Select
                    
                    RsPagos.MoveNext
                    If RsPagos.EOF Then
                        Exit Do
                    End If
                Else
                    If RsPagos!tipo_moneda = 14 Then
                        RsPagos.MoveNext
                        tipoMon = RsPagos!tipo_moneda
                        If RsPagos.EOF Then
                            Exit Do
                        End If
                    Else
                        ind = ind + 1
                        sprPago(0).GetText 1, ind, moneda
                    End If
                End If
            Loop

            If Val(moneda) = 99 Or Val(moneda) = 0 Then
                If Not RsPagos.EOF Then
                    RsPagos.MoveNext
                End If
                Exit Do
            Else
                sprPago(2).SetText 3, ind, str(PagoA)
                sprPago(2).SetText 6, ind, str(PagoE)
                'ind = ind + 1
                'sprPago(0).GetText 1, ind, moneda
            End If
            If RsPagos.EOF Then
                Exit Do
            End If
        Loop
        Do While Val(moneda) <> 0 And Val(moneda) <> 99
            ind = ind + 1
            sprPago(0).GetText 1, ind, moneda
        Loop
        sprPago(0).SetText 3, ind, str(TotPagoLA)
        sprPago(0).SetText 6, ind, str(TotPagoLE)
        sprPago(1).SetText 3, ind, str(TotPagoSA)
        sprPago(1).SetText 6, ind, str(TotPagoSE)
            
        sprPago(2).SetText 3, ind, str(TotPagoLA + TotPagoSA)
        sprPago(2).SetText 6, ind, str(TotPagoLE + TotPagoSE)
        
        If Val(moneda) = 99 Then
            Exit Do
        End If
'        ind = ind + 1
    Loop
    
L_Participacion sprPago(2)
L_Participacion sprPago(0)
L_Participacion sprPago(1)

End Sub

Private Sub L_DescripMonedas()
Dim sql As String
Dim rs As Recordset
Dim i

sql = "Select cod_moneda,descrip From ventas.monedas where cod_moneda not in (14,0) order by cod_moneda"
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        For i = 0 To 2
            sprPago(i).MaxRows = sprPago(i).MaxRows + 1
            
            sprPago(i).SetText 1, sprPago(i).MaxRows, str(rs!cod_moneda)
            sprPago(i).SetText 2, sprPago(i).MaxRows, Trim(rs!Descrip)
        Next
        rs.MoveNext
    Loop
    'Spread_TotalesGrillas sprPago, sprPago.MaxCols - 2, 2
    For i = 0 To 2
        sprPago(i).MaxRows = sprPago(i).MaxRows + 1
        Spread_TotalesLinea sprPago(i)
    
        sprPago(i).SetText 1, sprPago(i).MaxRows, "00"
        sprPago(i).SetText 2, sprPago(i).MaxRows, "Total Efectivo"
    Next
    Aplicacion.CerrarDAO rs
End If

sql = "Select cod_tarjeta,descrip From ventas.tarjetas Where cod_tarjeta in (1,2,3,4) ORDER BY cod_tarjeta"
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        For i = 0 To 2
        sprPago(i).MaxRows = sprPago(i).MaxRows + 1
        
        sprPago(i).SetText 1, sprPago(i).MaxRows, str(rs!cod_tarjeta)
        sprPago(i).SetText 2, sprPago(i).MaxRows, Trim(rs!Descrip)
        Next
        rs.MoveNext
    Loop
    For i = 0 To 2
        sprPago(i).MaxRows = sprPago(i).MaxRows + 1
          
        sprPago(i).SetText 1, sprPago(i).MaxRows, "98"
        sprPago(i).SetText 2, sprPago(i).MaxRows, "Otras"
        
        sprPago(i).MaxRows = sprPago(i).MaxRows + 1
        Spread_TotalesLinea sprPago(i)
    
        sprPago(i).SetText 1, sprPago(i).MaxRows, "99"
        sprPago(i).SetText 2, sprPago(i).MaxRows, "Total Tarjeta"
    Next
    Aplicacion.CerrarDAO rs
End If

End Sub
Private Sub L_LimpiarGrillas()
Dim i
    
For i = 0 To 2
    sprPago(i).MaxRows = 2
    
    sprEzeA(i).MaxRows = 0
    sprEzeAS(i).MaxRows = 0
    sprEzeB(i).MaxRows = 0
    sprAEP(i).MaxRows = 0
    sprInt(i).MaxRows = 0
    sprTotal(i).MaxRows = 0
    
Next
For i = 0 To 5
    mskActual(i).Text = 0
    mskEstim(i).Text = 0
    mskAntHora(i).Text = 0
    mskAnt(i).Text = 0
    
    labAntHora(i).caption = "%"
    labAnt(i).caption = "%"
    labEstim(i).caption = "%"
Next

For i = 3 To 4
    mskActual(i).Text = 0
    mskEstim(i).Text = 0
    mskAntHora(i).Text = 0
    mskAnt(i).Text = 0
    labAntHora(i).caption = "%"
    labAnt(i).caption = "%"
    labEstim(i).caption = "%"
Next i
    
SprVtasEzeA.MaxRows = 0
SprVtasEzeAS.MaxRows = 0
SprVtasEzeB.MaxRows = 0
SprVtasAep.MaxRows = 0
sprVtasInt.MaxRows = 0
sprTotalBsAs.MaxRows = 0
    
DoEvents
End Sub
Private Sub L_LlenarResumen()
Dim valor As Variant
Dim x As Integer

sprTotalBsAs.MaxRows = 4
sprTotalBsAs.SetText 1, 1, "EZE - A Lleg"
sprTotalBsAs.SetText 1, 2, "EZE - A Sal"
sprTotalBsAs.SetText 1, 3, "EZE - B"
sprTotalBsAs.SetText 1, 4, "AEP"
sprTotalBsAs.SetText 1, 5, "INT"

For x = 2 To 6
  SprVtasEzeA.GetText x, SprVtasEzeA.MaxRows, valor
  sprTotalBsAs.SetText x, 1, str(valor)
  SprVtasEzeAS.GetText x, SprVtasEzeAS.MaxRows, valor
  sprTotalBsAs.SetText x, 2, str(valor)
  SprVtasEzeB.GetText x, SprVtasEzeB.MaxRows, valor
  sprTotalBsAs.SetText x, 3, str(valor)
  SprVtasAep.GetText x, SprVtasAep.MaxRows, valor
  sprTotalBsAs.SetText x, 4, str(valor)
  If sprVtasInt.MaxRows > 0 Then
    sprVtasInt.GetText x, sprVtasInt.MaxRows, valor
    sprTotalBsAs.SetText x, 5, str(valor)
  End If

Next x

Spread_TotalesGrillas sprTotalBsAs, sprTotalBsAs.MaxCols - 2, 2
   
End Sub

Private Sub L_Participacion(spr As control)
Dim TotA00 As Variant, TotE00 As Variant, TotA99 As Variant, TotE99 As Variant
Dim Tot00 As Variant, Tot99 As Variant
Dim i, codigo As Variant, valor As Variant
Dim tipo As String

For i = 3 To spr.MaxRows
    spr.GetText 1, i, codigo
    If codigo = "00" Then
        spr.GetText 3, i, TotA00
        spr.GetText 6, i, TotE00
        spr.GetText 9, i, Tot00
    ElseIf codigo = "99" Then
        spr.GetText 3, i, TotA99
        spr.GetText 6, i, TotE99
        spr.GetText 9, i, Tot99
    End If
Next

tipo = "E"
For i = 3 To spr.MaxRows

    spr.GetText 1, i, codigo
    If tipo = "E" Then
        spr.GetText 3, i, valor
        If Val(TotA00) > 0 Then
            spr.SetText 4, i, str(Val(valor) / Val(TotA00) * 100)
            spr.SetText 5, i, str(Val(valor) / Val(TotA00 + TotA99) * 100)
        End If
        spr.GetText 6, i, valor
        If Val(TotE00) > 0 Then
            spr.SetText 7, i, str(Val(valor) / Val(TotE00) * 100)
            spr.SetText 8, i, str(Val(valor) / Val(TotE00 + TotE99) * 100)
        End If
        spr.GetText 9, i, valor
        If Val(Tot00) > 0 Then
            spr.SetText 10, i, str(Val(valor) / Val(Tot00) * 100)
            spr.SetText 11, i, str(Val(valor) / Val(Tot00 + Tot99) * 100)
        End If
    
    Else
        spr.GetText 3, i, valor
        If Val(TotA99) > 0 Then
            spr.SetText 4, i, str(Val(valor) / Val(TotA99) * 100)
            spr.SetText 5, i, str(Val(valor) / Val(TotA99 + TotA00) * 100)
        End If
        spr.GetText 6, i, valor
        If Val(TotE99) > 0 Then
            spr.SetText 7, i, str(Val(valor) / Val(TotE99) * 100)
            spr.SetText 8, i, str(Val(valor) / Val(TotE99 + TotE00) * 100)
        End If
        spr.GetText 9, i, valor
        If Val(Tot99) > 0 Then
            spr.SetText 10, i, str(Val(valor) / Val(Tot99) * 100)
            spr.SetText 11, i, str(Val(valor) / Val(Tot99 + Tot00) * 100)
        End If
    
    End If
    If codigo = "00" Then
        tipo = "T"
    End If
    
Next
End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim SqlZ As String
Dim rs As Recordset

On Error GoTo ErrOnline:

FrmOnline.caption = Aplicacion.SeteoProceso(FrmOnline.caption)

Aplicacion.ComienzoTrans

sql = "           SELECT cod_sdep,cod_local,cod_sloc,sum(cantidad) cantidad,sum(importe) importe,sum(ticket) ticket,cod_rubr "
sql = sql & " FROM (  "
sql = sql & "     (SELECT A.cod_depn, cod_ssdep cod_sdep,P.cod_rubr,T.cod_local,T.cod_sloc,SUM(importe) importe,SUM(cantidad) cantidad, 0 ticket "
sql = sql & "      FROM  TICKDIA.VIEW_VENTAS_RL@LINK_NT_VTA T,BAIRES.PRODUCTO p, VENTAS.apertura_sdep A "
sql = sql & "      WHERE p.Cod_prod = T.Cod_prod "
sql = sql & "      and a.cod_depn = t.cod_depn and a.cod_sdep = t.cod_sdep and a.cod_local = t.cod_local "
sql = sql & "      GROUP BY a.cod_depn,cod_ssdep,T.cod_local,T.cod_sloc,P.cod_rubr) "
sql = sql & "      UNION "
sql = sql & "     (SELECT a.cod_depn,cod_ssdep cod_sdep,P.cod_rubr,'L' cod_local,0 cod_sloc ,SUM(importe) importe,SUM(cantidad) cantidad, 0 ticket "
sql = sql & "      FROM  TICKDIA.VIEW_VENTAS_RL@LINK_NT_VTA T,BAIRES.PRODUCTO p, VENTAS.apertura_sdep A "
sql = sql & "      WHERE p.Cod_prod = T.Cod_prod "
sql = sql & "      and a.cod_depn = t.cod_depn and a.cod_sdep = t.cod_sdep and a.cod_local = t.cod_local "
sql = sql & "      GROUP BY a.cod_depn,cod_ssdep,P.cod_rubr) "
sql = sql & "      UNION "
sql = sql & "     (SELECT a.cod_depn,cod_ssdep cod_sdep,'AAA' cod_rubr,T.cod_local,T.cod_sloc ,SUM(importe) importe,SUM(cantidad) cantidad, 0 ticket "
sql = sql & "      FROM  TICKDIA.VIEW_VENTAS_RL@LINK_NT_VTA T,BAIRES.PRODUCTO p, VENTAS.apertura_sdep A "
sql = sql & "      WHERE p.Cod_prod = T.Cod_prod "
sql = sql & "      and a.cod_depn = t.cod_depn and a.cod_sdep = t.cod_sdep and a.cod_local = t.cod_local "
sql = sql & "      GROUP BY a.cod_depn,cod_ssdep,T.cod_local,T.cod_sloc) "
sql = sql & "      UNION "
sql = sql & "     (SELECT a.cod_depn,cod_ssdep cod_sdep,'AAA' cod_rubr,t.cod_local,cod_sloc ,0 importe,0 cantidad, count(*) ticket "
sql = sql & "      FROM  TICKDIA.ticket_h_dia@LINK_NT_VTA T, VENTAS.apertura_sdep A"
sql = sql & "      WHERE Importe > 0 And tipo_oper > 0 "
sql = sql & "      and a.cod_depn = t.cod_depn and a.cod_sdep = t.cod_sdep and a.cod_local = t.cod_local "
sql = sql & "      GROUP BY a.cod_depn,cod_ssdep,t.cod_local,cod_sloc) "
sql = sql & "      ) T  "
sql = sql & "GROUP BY t.cod_sdep,t.cod_local,cod_sloc,cod_rubr "
sql = sql & "ORDER BY cod_rubr,t.cod_sdep,t.cod_local, cod_sloc "


sqlX = " select tipo_pago,tipo_loc,tipo_moneda, nac , sum(importe) Importe "
sqlX = sqlX & " From ( "
sqlX = sqlX & " Select tipo_pago, tipo_loc,tipo_moneda,decode(cod_nac,1,'A','E') nac,sum(p.importe) importe "
sqlX = sqlX & " from baires.local l, tickdia.pagos_dia@LINK_NT_VTA p, tickdia.ticket_h_dia@LINK_NT_VTA t "
sqlX = sqlX & " Where l.cod_loc = t.cod_local AND t.fch_ticket = p.fch_ticket "
sqlX = sqlX & " and t.nro_ticket= p.nro_ticket and t.cod_caja = p.cod_caja "
sqlX = sqlX & " and t.cod_local = p.cod_local and t.cod_sloc=p.cod_sloc"
sqlX = sqlX & " and tipo_pago = 1 and t.cod_local <> 'L03' "
sqlX = sqlX & " group by tipo_pago,tipo_loc,decode(cod_nac,1,'A','E'),tipo_moneda "
sqlX = sqlX & " Union "
sqlX = sqlX & " Select tipo_pago,decode(MOD(cod_vuelo,2),0,'S','L') tipo_local ,tipo_moneda,decode(cod_nac,1,'A','E') nac,sum(p.importe) "
sqlX = sqlX & " From tickdia.pagos_dia@LINK_NT_VTA p, tickdia.ticket_h_dia@LINK_NT_VTA t"
sqlX = sqlX & " Where t.fch_ticket = P.fch_ticket And t.nro_ticket = P.nro_ticket "
sqlX = sqlX & " and t.cod_caja = p.cod_caja and t.cod_local = p.cod_local"
sqlX = sqlX & " and t.cod_sloc=p.cod_sloc and tipo_pago = 1 and t.cod_local = 'L03' "
sqlX = sqlX & " group by  tipo_pago,decode(MOD(cod_vuelo,2),0,'S','L'),decode(cod_nac,1,'A','E'),tipo_moneda "
sqlX = sqlX & " Union "
sqlX = sqlX & " select tipo_pago,tipo_loc "
sqlX = sqlX & ",decode(cod_tarjeta,1,1,16,1,23,1,2,2,18,2,25,2,3,3,19,3,26,3,4,4,17,4,24,4,98) cod_tarjeta "
sqlX = sqlX & ",decode(cod_nac,1,'A','E') nac,sum(p.importe) importe"
sqlX = sqlX & " From baires.local l, tickdia.pagos_dia@LINK_NT_VTA p,tickdia.pago_tarjeta_dia@LINK_NT_VTA j, tickdia.ticket_h_dia@LINK_NT_VTA t "
sqlX = sqlX & " Where l.cod_loc = t.cod_local And t.fch_ticket = P.fch_ticket "
sqlX = sqlX & " and t.nro_ticket= p.nro_ticket and t.cod_caja = p.cod_caja "
sqlX = sqlX & " and t.cod_local = p.cod_local and t.cod_sloc=p.cod_sloc "
sqlX = sqlX & " AND j.fch_ticket = p.fch_ticket and j.nro_ticket= p.nro_ticket "
sqlX = sqlX & " and j.cod_caja = p.cod_caja and j.cod_local = p.cod_local"
sqlX = sqlX & " and j.cod_sloc=p.cod_sloc and j.nro_secuencia = p.nro_secuencia "
sqlX = sqlX & " and tipo_pago = 2 and t.cod_local <> 'L03' "
sqlX = sqlX & " group by tipo_pago,tipo_loc,decode(cod_nac,1,'A','E'),decode(cod_tarjeta,1,1,16,1,23,1,2,2,18,2,25,2,3,3,19,3,26,3,4,4,17,4,24,4,98) "
sqlX = sqlX & " Union"
sqlX = sqlX & " select tipo_pago,decode(MOD(cod_vuelo,2),0,'S','L') tipo_local  "
sqlX = sqlX & ",decode(cod_tarjeta,1,1,16,1,23,1,2,2,18,2,25,2,3,3,19,3,26,3,4,4,17,4,24,4,98) COD_TARJETA "
sqlX = sqlX & ",decode(cod_nac,1,'A','E') nac,sum(p.importe) "
sqlX = sqlX & " From tickdia.pagos_dia@LINK_NT_VTA p,tickdia.pago_tarjeta_dia@LINK_NT_VTA j ,tickdia.ticket_h_dia@LINK_NT_VTA t"
sqlX = sqlX & " Where t.fch_ticket = P.fch_ticket And t.nro_ticket = P.nro_ticket "
sqlX = sqlX & " and t.cod_caja = p.cod_caja and t.cod_local = p.cod_local "
sqlX = sqlX & " and t.cod_sloc=p.cod_sloc "
sqlX = sqlX & " AND j.fch_ticket = p.fch_ticket and j.nro_ticket= p.nro_ticket "
sqlX = sqlX & " and j.cod_caja = p.cod_caja and j.cod_local = p.cod_local "
sqlX = sqlX & " and j.cod_sloc=p.cod_sloc and j.nro_secuencia = p.nro_secuencia "
sqlX = sqlX & " and t.cod_local = 'L03' and tipo_pago = 2 "
sqlX = sqlX & " group by tipo_pago,decode(MOD(cod_vuelo,2),0,'S','L'),decode(cod_nac,1,'A','E'),decode(cod_tarjeta,1,1,16,1,23,1,2,2,18,2,25,2,3,3,19,3,26,3,4,4,17,4,24,4,98) "
sqlX = sqlX & " )group by  tipo_pago,tipo_loc,tipo_moneda, nac "
sqlX = sqlX & "  order by  tipo_pago,tipo_moneda,tipo_loc, nac "

If Aplicacion.ObtenerRsDAO(sql, RsData) Then

    If Aplicacion.CantReg(RsData) > 0 Then
        L_DecoEspigon
    End If
    
    L_PonerTotales
    L_LlenarResumen
    
    If Aplicacion.ObtenerRsDAO(sqlX, RsPagos) Then
        L_DecoPagos
        Aplicacion.CerrarDAO RsPagos
    End If
    
    Aplicacion.CerrarDAO RsData
    
End If
'Aplicacion.TerminarConExitoTrans
'Exit Sub
ErrOnline:
    Aplicacion.TerminarConExitoTrans
    FrmOnline.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub

Private Sub L_SeteoEjecutar(valor As Boolean, dato As String)

If valor Then
    Select Case dato
        Case EZEAL
            UsoPorIntA = True
        Case EZEAS
            UsoPorIntAS = True
        Case EZEB
            UsoPorIntB = True
        Case AERO
            UsoPorAero = True
        Case INTE
            UsoPorInte = True
    End Select
    botEjecutar(1).Enabled = False
Else
    Select Case dato
        Case EZEAL
            UsoPorIntA = False
        Case EZEAS
            UsoPorIntAS = True
        Case EZEB
            UsoPorIntB = False
        Case AERO
            UsoPorAero = False
        Case INTE
            UsoPorInte = True
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorAero) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub

Private Sub L_TratarComparativos()
Dim valor
Dim sql As String
Dim rs As Recordset
Dim fch As String
Dim hora As Long

hora = Format$(txtTmpC.Text, "hhmmss")

sql = " SELECT min(fch_ticket) fch "
sql = sql & " FROM  TICKDIA.ticket_h_dia@LINK_NT_VTA  "
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    FchAnt = Format$(rs!fch - Aplicacion.anio, FTOFECHA)
Else
    FchAnt = Format$(CDate(Now) - Aplicacion.anio, FTOFECHA)
End If


SprVtasAep.GetText 2, SprVtasAep.MaxRows, valor
mskActual(0).Text = valor
SprVtasEzeA.GetText 2, SprVtasEzeA.MaxRows, valor
mskActual(2).Text = valor
SprVtasEzeAS.GetText 2, SprVtasEzeAS.MaxRows, valor
mskActual(1).Text = valor
SprVtasEzeB.GetText 2, SprVtasEzeB.MaxRows, valor
mskActual(3).Text = valor
sprVtasInt.GetText 2, sprVtasInt.MaxRows, valor
mskActual(5).Text = valor


If hora > 50000 Then
    sql = " select 'HORA' tipo ,decode(a.cod_depn, 'INT','INT',cod_ssdep) sdep, sum(importe) imp "
    sql = sql & " From ventas.ticket_h T,ventas.apertura_sdep A "
    sql = sql & " where a.cod_sdep not in ('CORD','MEND') And "
    sql = sql & " fch_ticket = " & func_ToDate(FchAnt) 'func_ToDate(Format(Now - Aplicacion.anio, FTOFECHA))
    sql = sql & " and fch_ticket = fch_proceso "
    sql = sql & " and hora_fin <= " & Format$(txtTmpC.Text, "hhmmss")
    sql = sql & " and a.cod_depn = t.cod_depn and a.cod_sdep = t.cod_sdep and a.cod_local = t.cod_local "
    sql = sql & " group by decode(a.cod_depn, 'INT','INT',cod_ssdep) "
    sql = sql & " Union select 'ANT' tipo ,decode(a.cod_depn, 'INT','INT',cod_ssdep) sdep, sum(importe) imp"
    sql = sql & " From ventas.ticket_h T,ventas.apertura_sdep A"
    sql = sql & " where a.cod_sdep not in ('CORD','MEND') And "
    sql = sql & " fch_ticket = " & func_ToDate(FchAnt) 'func_ToDate(Format(Now - Aplicacion.anio, FTOFECHA))
    sql = sql & " and a.cod_depn = t.cod_depn and a.cod_sdep = t.cod_sdep and a.cod_local = t.cod_local "
    sql = sql & " group by decode(a.cod_depn, 'INT','INT',cod_ssdep) "
    sql = sql & " Union Select 'EST' tipo,decode(depn,'INT','INT',sdep) sdep, sum(trunc(porc,2)) imp"
    sql = sql & " From estadis.porcentaje_d"
    sql = sql & " Where anio = " & Year(Now) & " And Mes = " & Month(Now)
    sql = sql & " And Tipo_porc = 'I' And Nivel = 0 and sdep not in ('MEND','CORD') "
    sql = sql & " and dia = " & func_ToDate(Format(Now, FTOFECHA))
    sql = sql & " Group by decode(depn,'INT','INT',sdep) "
Else
    sql = " SELECT tipo,sdep,sum(imp) imp FROM ("
    sql = sql & " select 'HORA' tipo ,decode(a.cod_depn, 'INT','INT',cod_ssdep) sdep, sum(importe) imp "
    sql = sql & " From ventas.ticket_h T,ventas.apertura_sdep A "
    sql = sql & " where a.cod_sdep not in ('CORD','MEND')  "
    sql = sql & " and fch_ticket = " & func_ToDate(FchAnt) 'func_ToDate(Format(Now - Aplicacion.anio, FTOFECHA))
    sql = sql & " and fch_ticket = fch_proceso "
    sql = sql & " and a.cod_depn = t.cod_depn and a.cod_sdep = t.cod_sdep and a.cod_local = t.cod_local "
    sql = sql & " group by decode(a.cod_depn, 'INT','INT',cod_ssdep) "
    sql = sql & " Union select 'ANT' tipo ,decode(a.cod_depn, 'INT','INT',cod_ssdep) sdep, sum(importe) imp"
    sql = sql & " From ventas.ticket_h T,ventas.apertura_sdep A"
    sql = sql & " where a.cod_sdep not in ('CORD','MEND') "
    sql = sql & " and fch_ticket = " & func_ToDate(FchAnt) 'func_ToDate(Format(Now - Aplicacion.anio, FTOFECHA))
    sql = sql & " and a.cod_depn = t.cod_depn and a.cod_sdep = t.cod_sdep and a.cod_local = t.cod_local "
    sql = sql & " group by  decode(a.cod_depn, 'INT','INT',cod_ssdep)  "
    sql = sql & " Union Select 'EST' tipo,decode(depn,'INT','INT',sdep) sdep, trunc(porc,2) imp"
    sql = sql & " From estadis.porcentaje_d"
    sql = sql & " Where anio = " & Year(Now) & " And Mes = " & Month(Now)
    sql = sql & " And Tipo_porc = 'I' And Nivel = 0 and sdep not in ('MEND','CORD') "
    sql = sql & " and dia = " & func_ToDate(Format(Now, FTOFECHA))
    sql = sql & " Group by decode(depn,'INT','INT',sdep)"
    sql = sql & " UNION select 'HORA' tipo ,decode(a.cod_depn, 'INT','INT',cod_ssdep) sdep, sum(importe) imp "
    sql = sql & " From ventas.ticket_h T,ventas.apertura_sdep A "
    sql = sql & " where  a.cod_sdep not in ('CORD','MEND')  "
    sql = sql & " and fch_ticket = " & func_ToDate(FchAnt) 'func_ToDate(Format(Now - Aplicacion.anio, FTOFECHA))
    sql = sql & " and fch_ticket + 1 = fch_proceso "
    sql = sql & " and hora_fin <= " & Format$(txtTmpC.Text, "hhmmss")
    sql = sql & " and a.cod_depn = t.cod_depn and a.cod_sdep = t.cod_sdep and a.cod_local = t.cod_local "
    sql = sql & " group by decode(a.cod_depn, 'INT','INT',cod_ssdep)  )"
    sql = sql & " GROUP BY tipo,sdep"
End If
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        Select Case rs!Sdep
            Case "AEP"
                Select Case rs!tipo
                    Case "ANT"
                        mskAnt(0).Text = rs!imp
                    Case "EST"
                        mskEstim(0).Text = rs!imp
                    Case "HORA"
                        mskAntHora(0).Text = rs!imp
                End Select
            Case "INTAS"
                 Select Case rs!tipo
                    Case "ANT"
                        mskAnt(1).Text = rs!imp
                    Case "EST"
                        mskEstim(1).Text = rs!imp
                    Case "HORA"
                        mskAntHora(1).Text = rs!imp
                 End Select
            Case "INTAL"
                Select Case rs!tipo
                    Case "ANT"
                        mskAnt(2).Text = rs!imp
                    Case "EST"
                        mskEstim(2).Text = rs!imp
                    Case "HORA"
                        mskAntHora(2).Text = rs!imp
                 End Select
            Case "INTB"
                Select Case rs!tipo
                    Case "ANT"
                        mskAnt(3).Text = rs!imp
                    Case "EST"
                        mskEstim(3).Text = rs!imp
                    Case "HORA"
                        mskAntHora(3).Text = rs!imp
                End Select
            Case "INT"
                Select Case rs!tipo
                    Case "ANT"
                        mskAnt(5).Text = rs!imp
                    Case "EST"
                        mskEstim(5).Text = rs!imp
                    Case "HORA"
                        mskAntHora(5).Text = rs!imp
                
                End Select
        End Select
        rs.MoveNext
    Loop
    Aplicacion.CerrarDAO rs
End If

L_Diferencias

End Sub

Private Sub L_Diferencias()
Dim i
Dim Importe As Single
'FchAnt = Format$(CDate(Now) - Aplicacion.anio, FTOFECHA)

Label6(3).caption = FchAnt & " a la misma hora"
Label6(2).caption = FchAnt & " completo"

For i = 0 To 3

If Val(mskAntHora(i).Text) > 0 Then
Importe = Format((Val(mskActual(i).Text) - Val(mskAntHora(i).Text)) / Val(mskAntHora(i).Text) * 100, "#.00")
End If
labAntHora(i).caption = Importe & "%"
    If Importe < 0 Then
        labAntHora(i).ForeColor = vbRed
    Else
        labAntHora(i).ForeColor = vbBlue
    End If
 If Val(mskAnt(i).Text) > 0 Then
 Importe = Format((Val(mskActual(i).Text) - Val(mskAnt(i).Text)) / Val(mskAnt(i).Text) * 100, "#.00")
 End If
labAnt(i).caption = Importe & "%"
    If Importe < 0 Then
        labAnt(i).ForeColor = vbRed
    Else
        labAnt(i).ForeColor = vbBlue
    End If
    
If Val(mskEstim(i).Text) > 0 Then
    Importe = Format((Val(mskActual(i).Text) - Val(mskEstim(i).Text)) / Val(mskEstim(i).Text) * 100, "#.00")
End If
labEstim(i).caption = Importe & "%"
    If Importe < 0 Then
        labEstim(i).ForeColor = vbRed
    Else
        labEstim(i).ForeColor = vbBlue
    End If

mskActual(4).Text = Val(mskActual(4).Text) + Val(mskActual(i).Text)
mskAntHora(4).Text = Val(mskAntHora(4).Text) + Val(mskAntHora(i).Text)
mskAnt(4).Text = Val(mskAnt(4).Text) + Val(mskAnt(i).Text)
mskEstim(4).Text = Val(mskEstim(4).Text) + Val(mskEstim(i).Text)
Next

'Parche para  agregar Bariloche
mskActual(4).Text = Val(mskActual(4).Text) + Val(mskActual(5).Text)
mskAntHora(4).Text = Val(mskAntHora(4).Text) + Val(mskAntHora(5).Text)
mskAnt(4).Text = Val(mskAnt(4).Text) + Val(mskAnt(5).Text)
mskEstim(4).Text = Val(mskEstim(4).Text) + Val(mskEstim(5).Text)

'Fin parche



Importe = Format((Val(mskActual(4).Text) - Val(mskAntHora(4).Text)) / Val(mskAntHora(4).Text) * 100, "#.00")
labAntHora(4).caption = Importe & "%"
    If Importe < 0 Then
        labAntHora(4).ForeColor = vbRed
    Else
        labAntHora(4).ForeColor = vbBlue
    End If

Importe = Format((Val(mskActual(4).Text) - Val(mskAnt(4).Text)) / Val(mskAnt(4).Text) * 100, "#.00")
labAnt(4).caption = Importe & "%"
    If Importe < 0 Then
        labAnt(4).ForeColor = vbRed
    Else
        labAnt(4).ForeColor = vbBlue
    End If

If Val(mskEstim(4).Text) > 0 Then
Importe = Format((Val(mskActual(4).Text) - Val(mskEstim(4).Text)) / Val(mskEstim(4).Text) * 100, "#.00")
End If
labEstim(4).caption = Importe & "%"
    If Importe < 0 Then
        labEstim(4).ForeColor = vbRed
    Else
        labEstim(4).ForeColor = vbBlue
    End If

End Sub

Private Sub L_TratarExcel(titulo As String, subTit As String)

Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim filaant As Integer
Dim i As Integer
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

FrmOnline.caption = Aplicacion.SeteoProceso(FrmOnline.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
'    AppExcel.application.Visible = True
    
    ReDim titCol(sprTotal(0).MaxCols)
    Col = 1
    fila = 5
    
    Exl_PonerValor AppExcel, 3, 1, titulo & " a las " & Format(Now, "hh:mm:ss")
    rango = Exl_rangos(1, 3, 1, sprTotal(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, sprTotal(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Importe"
    rango = Exl_rangos(fila, fila, 1, sprTotal(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 1
    
    For i = 1 To sprTotal(0).MaxCols
        sprTotal(0).GetText i, 0, tit
        titCol(i) = tit
    Next
                      
    fila = fila + 1
    filaant = fila
    Exl_BajarGrillaExel sprTotal(0), AppExcel, fila, Col, titCol
    fila = fila + sprTotal(0).MaxRows
    rango = Exl_rangos(filaant, fila, sprTotal(0).MaxCols, sprTotal(0).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
     rango = Exl_rangos(filaant, fila, sprTotal(0).MaxCols, sprTotal(0).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
     rango = Exl_rangos(filaant, fila, 2, sprTotal(0).MaxCols)
     Exl_Format AppExcel, rango
    
    '''''''''''''
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Unidades"
    rango = Exl_rangos(fila, fila, 1, sprTotal(1).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
                                                   
        
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel sprTotal(1), AppExcel, fila, Col, titCol
    fila = fila + sprTotal(1).MaxRows
    rango = Exl_rangos(filaant, fila, sprTotal(1).MaxCols, sprTotal(1).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
     
     rango = Exl_rangos(filaant, fila, sprTotal(1).MaxCols, sprTotal(1).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
     'rango = Exl_rangos(filaant, fila, 2, SprTotal(1).MaxCols)
     'Exl_Format AppExcel, rango

     '''''''''''''''''
     
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Promedio de Importe / Unidades"
    rango = Exl_rangos(fila, fila, 1, sprTotal(2).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
                                                   
        
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel sprTotal(2), AppExcel, fila, Col, titCol
    fila = fila + sprTotal(2).MaxRows
    rango = Exl_rangos(filaant, fila, sprTotal(2).MaxCols, sprTotal(2).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
     
     
     rango = Exl_rangos(filaant, fila, sprTotal(2).MaxCols, sprTotal(2).MaxCols)
     rango = Exl_rangos(filaant, fila, 2, sprTotal(2).MaxCols)
     Exl_Format AppExcel, rango
     
     rango = Exl_rangos(filaant, fila, sprTotal(2).MaxCols, sprTotal(2).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
     rango = Exl_rangos(filaant, fila, 2, sprTotal(2).MaxCols)
     Exl_Format AppExcel, rango
     
     '''''''''''''''''
      
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    
    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel.Application.ActiveSheet = Nothing
    Set AppExcel = Nothing
    
End If

ErrorExl:

    FrmOnline.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub



Private Sub L_TratarExcel_Vtas(titulo As String, subTit As String)

Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim filaant As Integer
Dim i As Integer
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

FrmOnline.caption = Aplicacion.SeteoProceso(FrmOnline.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
'    AppExcel.application.Visible = True
    
    ReDim titCol(SprVtasEzeA.MaxCols)
    Col = 1
    fila = 5
    
    Exl_PonerValor AppExcel, 3, 1, titulo & " a las " & Format(Now, "hh:mm:ss")
    rango = Exl_rangos(1, 3, 1, SprVtasEzeA.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, SprVtasEzeA.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
      
    fila = fila + 1
    
    For i = 1 To SprVtasEzeA.MaxCols
        SprVtasEzeA.GetText i, 0, tit
        titCol(i) = tit
    Next
                      
    fila = fila + 1
    filaant = fila
    Exl_BajarGrillaExel SprVtasEzeA, AppExcel, fila, Col, titCol
    fila = fila + SprVtasEzeA.MaxRows
    rango = Exl_rangos(fila, fila, 1, SprVtasEzeA.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris

    rango = Exl_rangos(filaant, fila, 2, 2)
    Exl_Format AppExcel, rango
    rango = Exl_rangos(filaant, fila, 5, 6)
    Exl_Format AppExcel, rango
    
    '''''''''''''
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Espigón A Salidas"
    rango = Exl_rangos(fila, fila, 1, SprVtasEzeAS.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
        
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel SprVtasEzeAS, AppExcel, fila, Col, titCol
    fila = fila + SprVtasEzeAS.MaxRows
    rango = Exl_rangos(fila, fila, 1, SprVtasEzeAS.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    rango = Exl_rangos(filaant, fila, 2, 2)
    Exl_Format AppExcel, rango
    rango = Exl_rangos(filaant, fila, 5, 6)
    Exl_Format AppExcel, rango
     
     '''''''''''''''''
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Espigón B"
    rango = Exl_rangos(fila, fila, 1, SprVtasEzeB.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
        
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel SprVtasEzeB, AppExcel, fila, Col, titCol
    fila = fila + SprVtasEzeB.MaxRows
    rango = Exl_rangos(fila, fila, 1, SprVtasEzeB.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
'     rango = Exl_rangos(filaant, fila, SprVtasEzeB.MaxCols, SprVtasEzeB.MaxCols)
'     rango = Exl_rangos(filaant, fila, 2, SprVtasEzeB.MaxCols)
'     Exl_Format AppExcel, rango
    rango = Exl_rangos(filaant, fila, 2, 2)
    Exl_Format AppExcel, rango
    rango = Exl_rangos(filaant, fila, 5, 6)
    Exl_Format AppExcel, rango
     '''''''''''''''''
     
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "AEROPARQUE"
    rango = Exl_rangos(fila, fila, 1, SprVtasAep.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
                                                           
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel SprVtasAep, AppExcel, fila, Col, titCol
    fila = fila + SprVtasAep.MaxRows
    rango = Exl_rangos(fila, fila, 1, SprVtasAep.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
     
     'rango = Exl_rangos(filaant, fila, SprVtasAep.MaxCols, SprVtasAep.MaxCols)
     'rango = Exl_rangos(filaant, fila, 2, SprVtasAep.MaxCols)
     'Exl_Format AppExcel, rango
     
    rango = Exl_rangos(filaant, fila, 2, 2)
    Exl_Format AppExcel, rango
    rango = Exl_rangos(filaant, fila, 5, 6)
    Exl_Format AppExcel, rango

     '''''''''''''''''
      
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    'AppExcel.application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    
    AppExcel.SaveAs NOMBRE & ".xls"
'    AppExcel.Workbooks.Open nombre & ".xls"
    Set AppExcel.Application.ActiveSheet = Nothing
    Set AppExcel = Nothing
    
End If

ErrorExl:

    FrmOnline.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub

Private Sub L_TratarExcel_Aep(titulo As String, subTit As String)

Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim filaant As Integer
Dim i As Integer
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

FrmOnline.caption = Aplicacion.SeteoProceso(FrmOnline.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    ReDim titCol(sprAEP(0).MaxCols)
    Col = 1
    fila = 5
    
    Exl_PonerValor AppExcel, 3, 1, titulo & " a las " & Format(Now, "hh:mm:ss")
    rango = Exl_rangos(1, 3, 1, sprAEP(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, sprAEP(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Importe"
    rango = Exl_rangos(fila, fila, 1, sprAEP(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 1
    
    For i = 1 To sprAEP(0).MaxCols
        sprAEP(0).GetText i, 0, tit
        titCol(i) = tit
    Next
        
    fila = fila + 1
    filaant = fila
    Exl_BajarGrillaExel sprAEP(0), AppExcel, fila, Col, titCol
    fila = fila + sprAEP(0).MaxRows
    rango = Exl_rangos(filaant, fila, sprAEP(0).MaxCols, sprAEP(0).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
     rango = Exl_rangos(filaant, fila, sprAEP(0).MaxCols, sprAEP(0).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
     rango = Exl_rangos(filaant, fila, 2, sprAEP(0).MaxCols)
     Exl_Format AppExcel, rango
    
    '''''''''''''
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Unidades"
    rango = Exl_rangos(fila, fila, 1, sprAEP(1).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris

    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel sprAEP(1), AppExcel, fila, Col, titCol
    fila = fila + sprAEP(1).MaxRows
    rango = Exl_rangos(filaant, fila, sprAEP(1).MaxCols, sprAEP(1).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
     
     rango = Exl_rangos(filaant, fila, sprAEP(1).MaxCols, sprAEP(1).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
'     rango = Exl_rangos(filaant, fila, 2, SprAep(1).MaxCols)
'     Exl_Format AppExcel, rango
     
     '''''''''''''''''
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Promedio de Importe / Unidades"
    rango = Exl_rangos(fila, fila, 1, sprAEP(2).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
                                                   
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel sprAEP(2), AppExcel, fila, Col, titCol
    fila = fila + sprAEP(2).MaxRows
    rango = Exl_rangos(filaant, fila, sprAEP(2).MaxCols, sprAEP(2).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
     
     rango = Exl_rangos(filaant, fila, sprAEP(2).MaxCols, sprAEP(2).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
     rango = Exl_rangos(filaant, fila, 2, sprAEP(2).MaxCols)
     Exl_Format AppExcel, rango
     
     '''''''''''''''''
      
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    
    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel.Application.ActiveSheet = Nothing
    Set AppExcel = Nothing
    
End If

ErrorExl:

    FrmOnline.caption = Aplicacion.SeteoFin
    Exit Sub

    
End Sub
Private Sub L_TratarExcel_EzeA(titulo As String, subTit As String)

Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim filaant As Integer
Dim i As Integer
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

FrmOnline.caption = Aplicacion.SeteoProceso(FrmOnline.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
'    AppExcel.application.Visible = True
    
    ReDim titCol(sprEzeA(0).MaxCols)
    Col = 1
    fila = 5
    
    Exl_PonerValor AppExcel, 3, 1, titulo & " a las " & Format(Now, "hh:mm:ss")
    rango = Exl_rangos(1, 3, 1, sprEzeA(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, sprEzeA(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Importe"
    rango = Exl_rangos(fila, fila, 1, sprEzeA(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 1
    
    For i = 1 To sprEzeA(0).MaxCols
        sprEzeA(0).GetText i, 0, tit
        titCol(i) = tit
    Next
        
    fila = fila + 1
    filaant = fila
    Exl_BajarGrillaExel sprEzeA(0), AppExcel, fila, Col, titCol
    fila = fila + sprEzeA(0).MaxRows
    rango = Exl_rangos(filaant, fila, sprEzeA(0).MaxCols, sprEzeA(0).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
     rango = Exl_rangos(filaant, fila, sprEzeA(0).MaxCols, sprEzeA(0).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
     rango = Exl_rangos(filaant, fila, 2, sprEzeA(0).MaxCols)
     Exl_Format AppExcel, rango
    
    '''''''''''''
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Unidades"
    rango = Exl_rangos(fila, fila, 1, sprEzeA(1).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
        
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel sprEzeA(1), AppExcel, fila, Col, titCol
    fila = fila + sprEzeA(1).MaxRows
    rango = Exl_rangos(filaant, fila, sprEzeA(1).MaxCols, sprEzeA(1).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
     
    'rango = Exl_rangos(filaant, fila, 2, SprEzeA(1).MaxCols)
    'Exl_Format AppExcel, rango
     
     '''''''''''''''''
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Promedio de Importe / Unidades"
    rango = Exl_rangos(fila, fila, 1, sprEzeA(2).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
                                                   
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel sprEzeA(2), AppExcel, fila, Col, titCol
    fila = fila + sprEzeA(2).MaxRows
    rango = Exl_rangos(filaant, fila, sprEzeA(2).MaxCols, sprEzeA(2).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
     
     rango = Exl_rangos(filaant, fila, sprEzeA(2).MaxCols, sprEzeA(2).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
     rango = Exl_rangos(filaant, fila, 2, sprEzeA(2).MaxCols)
     Exl_Format AppExcel, rango
     
     '''''''''''''''''
      
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    
    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel.Application.ActiveSheet = Nothing
    Set AppExcel = Nothing
    
End If

ErrorExl:

    FrmOnline.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub L_TratarExcel_EzeAS(titulo As String, subTit As String)

Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim filaant As Integer
Dim i As Integer
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

FrmOnline.caption = Aplicacion.SeteoProceso(FrmOnline.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
'    AppExcel.application.Visible = True
    
    ReDim titCol(sprEzeAS(0).MaxCols)
    Col = 1
    fila = 5
    
    Exl_PonerValor AppExcel, 3, 1, titulo & " a las " & Format(Now, "hh:mm:ss")
    rango = Exl_rangos(1, 3, 1, sprEzeA(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, sprEzeA(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Importe"
    rango = Exl_rangos(fila, fila, 1, sprEzeAS(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 1
    
    For i = 1 To sprEzeAS(0).MaxCols
        sprEzeA(0).GetText i, 0, tit
        titCol(i) = tit
    Next
        
    fila = fila + 1
    filaant = fila
    Exl_BajarGrillaExel sprEzeAS(0), AppExcel, fila, Col, titCol
    fila = fila + sprEzeAS(0).MaxRows
    rango = Exl_rangos(filaant, fila, sprEzeAS(0).MaxCols, sprEzeAS(0).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
     rango = Exl_rangos(filaant, fila, sprEzeAS(0).MaxCols, sprEzeAS(0).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
     rango = Exl_rangos(filaant, fila, 2, sprEzeAS(0).MaxCols)
     Exl_Format AppExcel, rango
    
    '''''''''''''
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Unidades"
    rango = Exl_rangos(fila, fila, 1, sprEzeAS(1).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
        
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel sprEzeAS(1), AppExcel, fila, Col, titCol
    fila = fila + sprEzeAS(1).MaxRows
    rango = Exl_rangos(filaant, fila, sprEzeAS(1).MaxCols, sprEzeAS(1).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
     
    'rango = Exl_rangos(filaant, fila, 2, SprEzeA(1).MaxCols)
    'Exl_Format AppExcel, rango
     
     '''''''''''''''''
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Promedio de Importe / Unidades"
    rango = Exl_rangos(fila, fila, 1, sprEzeAS(2).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
                                                   
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel sprEzeAS(2), AppExcel, fila, Col, titCol
    fila = fila + sprEzeAS(2).MaxRows
    rango = Exl_rangos(filaant, fila, sprEzeAS(2).MaxCols, sprEzeAS(2).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
     
     rango = Exl_rangos(filaant, fila, sprEzeAS(2).MaxCols, sprEzeAS(2).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
     rango = Exl_rangos(filaant, fila, 2, sprEzeAS(2).MaxCols)
     Exl_Format AppExcel, rango
     
     '''''''''''''''''
      
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    
    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel.Application.ActiveSheet = Nothing
    Set AppExcel = Nothing
    
End If

ErrorExl:

    FrmOnline.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub

Private Sub L_TratarExcel_EzeB(titulo As String, subTit As String)

Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim filaant As Integer
Dim i As Integer
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

FrmOnline.caption = Aplicacion.SeteoProceso(FrmOnline.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
'    AppExcel.application.Visible = True
    
    ReDim titCol(sprEzeB(0).MaxCols)
    Col = 1
    fila = 5
    
    Exl_PonerValor AppExcel, 3, 1, titulo & " a las " & Format(Now, "hh:mm:ss")
    rango = Exl_rangos(1, 3, 1, sprEzeB(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, sprEzeB(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Importe"
    rango = Exl_rangos(fila, fila, 1, sprEzeB(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 1
    
    For i = 1 To sprEzeB(0).MaxCols
        sprEzeB(0).GetText i, 0, tit
        titCol(i) = tit
    Next
              
    fila = fila + 1
    filaant = fila
    Exl_BajarGrillaExel sprEzeB(0), AppExcel, fila, Col, titCol
    fila = fila + sprEzeB(0).MaxRows
    rango = Exl_rangos(filaant, fila, sprEzeB(0).MaxCols, sprEzeB(0).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
     rango = Exl_rangos(filaant, fila, sprEzeB(0).MaxCols, sprEzeB(0).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
     rango = Exl_rangos(filaant, fila, 2, sprEzeB(0).MaxCols)
     Exl_Format AppExcel, rango
    
    '''''''''''''
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Unidades"
    rango = Exl_rangos(fila, fila, 1, sprEzeB(1).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
                                                   
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel sprEzeB(1), AppExcel, fila, Col, titCol
    fila = fila + sprEzeB(1).MaxRows
    rango = Exl_rangos(filaant, fila, sprEzeB(1).MaxCols, sprEzeB(1).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
     
     'rango = Exl_rangos(filaant, fila, 2, SprEzeB(1).MaxCols)
     'Exl_Format AppExcel, rango
     
     '''''''''''''''''
    fila = fila + 3
    
    Exl_PonerValor AppExcel, fila, Col, "Promedio de Importe / Unidades"
    rango = Exl_rangos(fila, fila, 1, sprEzeB(2).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
                                                   
    fila = fila + 2
    filaant = fila
    Exl_BajarGrillaExel sprEzeB(2), AppExcel, fila, Col, titCol
    fila = fila + sprEzeB(2).MaxRows
    rango = Exl_rangos(filaant, fila, sprEzeB(2).MaxCols, sprEzeB(2).MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
     
     rango = Exl_rangos(filaant, fila, sprEzeB(2).MaxCols, sprEzeB(2).MaxCols)
     Exl_ColorInt AppExcel, rango, Exl_Gris
     rango = Exl_rangos(filaant, fila, 2, sprEzeB(2).MaxCols)
     Exl_Format AppExcel, rango
     
     '''''''''''''''''
      
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    
    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel.Application.ActiveSheet = Nothing
    Set AppExcel = Nothing
    
End If

ErrorExl:

    FrmOnline.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub


Private Sub botEjecutar_Click(Index As Integer)

Select Case Index
    Case 0
        L_LimpiarGrillas

        txtTmpC.Text = Format$(Now, "hh:mm:ss")
        L_Refrescar
        L_TratarComparativos
        'txtTmpC.Text = Format$(Now, "hh:mm:ss")
    Case 2
        Unload Me
End Select

End Sub







Private Sub botExcel_Click(Index As Integer)

Select Case Index
  Case Is = 0
     L_TratarExcel "Ventas Online por Rubro", "Total Buenos Aires"
  Case Is = 1
     L_TratarExcel_EzeA "Ventas Online por Rubro", "Espigón A Llegadas"
  Case Is = 2
     L_TratarExcel_EzeAS "Ventas Online por Rubro", "Espigón A Salidas"
  Case Is = 3
     L_TratarExcel_EzeB "Ventas Online por Rubro", "Espigón B"
  Case Is = 8
     L_TratarExcel_Aep "Ventas Online por Rubro", "Aeroparque"
  Case Is = 4, 5, 6, 7, 9
     L_TratarExcel_Vtas "Ventas Online por Local", "Espigón A Llegadas"
End Select

End Sub

Private Sub botSemaforo_Click()

If Actualizar Then
    botSemaforo.Picture = Image1.Picture
Else
    botSemaforo.Picture = Image2.Picture
End If
Actualizar = Not Actualizar
sldTmp.Enabled = Actualizar
End Sub

Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6500
Me.Width = 11000

UsoPorIntA = False
UsoPorIntB = False
UsoPorAero = False
labFecha.caption = Format$(Now, "dddd, dd mmmm - yyyy")

L_LimpiarGrillas
frmPrincipal.lstForms.AddItem "frmOnline"

Actualizar = True
botSemaforo.Picture = Image2.Picture
sldTmp.Enabled = True

DoEvents

Call botEjecutar_Click(0)

End Sub





Private Sub Form_Unload(Cancel As Integer)
    FuncLocal_SacarForm "frmLGC"
End Sub

Private Sub L_PonerTotales()
Dim n As Integer

For n = 0 To 1
  Spread_TotalesGrillas sprTotal(n), sprTotal(n).MaxCols - 2, 2
  Spread_TotalesGrillas sprEzeA(n), sprEzeA(n).MaxCols - 2, 2
  Spread_TotalesGrillas sprEzeAS(n), sprEzeAS(n).MaxCols - 2, 2
  Spread_TotalesGrillas sprEzeB(n), sprEzeB(n).MaxCols - 2, 2
  Spread_TotalesGrillas sprAEP(n), sprAEP(n).MaxCols - 2, 2
  Spread_TotalesGrillas sprInt(n), sprInt(n).MaxCols - 2, 2
Next n

'Pone los promedios totales
Func_PromediosCol sprEzeA(0), sprEzeA(1), sprEzeA(2), 5, 1, sprEzeA(0).MaxRows
Func_PromediosCol sprEzeAS(0), sprEzeAS(1), sprEzeAS(2), 7, 1, sprEzeAS(0).MaxRows
Func_PromediosCol sprEzeB(0), sprEzeB(1), sprEzeB(2), 6, 1, sprEzeB(0).MaxRows
Func_PromediosCol sprAEP(0), sprAEP(1), sprAEP(2), 5, 1, sprAEP(0).MaxRows
Func_PromediosCol sprInt(0), sprInt(1), sprInt(2), 5, 1, sprInt(0).MaxRows
Func_PromediosCol sprTotal(0), sprTotal(1), sprTotal(2), 7, 1, sprTotal(0).MaxRows

Spread_TotalesGrillas SprVtasEzeA, 3, 2
Spread_TotalesGrillas SprVtasEzeAS, 3, 2
Spread_TotalesGrillas SprVtasEzeB, 3, 2
Spread_TotalesGrillas SprVtasAep, 3, 2
Spread_TotalesGrillas sprVtasInt, 3, 2

Spread_TotalesLinea SprVtasEzeA
Spread_TotalesLinea SprVtasEzeAS
Spread_TotalesLinea SprVtasEzeB
Spread_TotalesLinea SprVtasAep
Spread_TotalesLinea sprVtasInt

sprTotal(2).MaxRows = sprTotal(2).MaxRows + 1
Spread_TotalesLinea sprTotal(2)

sprEzeA(2).MaxRows = sprEzeA(2).MaxRows + 1
Spread_TotalesLinea sprEzeA(2)

sprEzeAS(2).MaxRows = sprEzeAS(2).MaxRows + 1
Spread_TotalesLinea sprEzeAS(2)

sprEzeB(2).MaxRows = sprEzeB(2).MaxRows + 1
Spread_TotalesLinea sprEzeB(2)

sprAEP(2).MaxRows = sprAEP(2).MaxRows + 1
Spread_TotalesLinea sprAEP(2)

sprAEP(2).MaxRows = sprInt(2).MaxRows + 1
Spread_TotalesLinea sprInt(2)


End Sub


Private Sub sldTmp_Scroll()
    labMin.caption = sldTmp.Value
End Sub

Private Sub Timer1_Timer()
txtTmpA.Text = Format$(Now, "hh:mm:ss")

If Actualizar Then
    If ((Format$(txtTmpA.Text, "hhmmss")) - (Format$(txtTmpC.Text, "hhmmss"))) > (sldTmp.Value * 100) Then
        Call botEjecutar_Click(0)
    End If
End If
End Sub


