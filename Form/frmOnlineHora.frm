VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmOnlineHora 
   Caption         =   "Venta Horaria"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   9195
   Begin ComctlLib.Slider sldTmp 
      Height          =   210
      Left            =   6510
      TabIndex        =   25
      Top             =   570
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   370
      _Version        =   327680
      MouseIcon       =   "frmOnlineHora.frx":0000
      OLEDropMode     =   1
      LargeChange     =   1
      Min             =   1
      SelStart        =   5
      Value           =   5
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2730
      Top             =   5205
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
         Picture         =   "frmOnlineHora.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   135
         Picture         =   "frmOnlineHora.frx":011E
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmOnlineHora.frx":0220
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   990
         Width           =   570
      End
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   3915
      Left            =   150
      TabIndex        =   2
      Top             =   1365
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   6906
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   4
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
      TabPicture(0)   =   "frmOnlineHora.frx":0A42
      Tab(0).ControlCount=   6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "botExcel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "sprTotCIA"
      Tab(0).Control(5).Enabled=   0   'False
      TabCaption(1)   =   "EZEIZA"
      TabPicture(1)   =   "frmOnlineHora.frx":0A5E
      Tab(1).ControlCount=   6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "botExcelEzeA"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "sprDep(0)"
      Tab(1).Control(5).Enabled=   0   'False
      TabCaption(2)   =   "AEROPARQUE"
      TabPicture(2)   =   "frmOnlineHora.frx":0A7A
      Tab(2).ControlCount=   6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label14(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label15(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label16(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "botExcelEzeb"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "sprDep(1)"
      Tab(2).Control(5).Enabled=   0   'False
      TabCaption(3)   =   "BARILOCHE"
      TabPicture(3)   =   "frmOnlineHora.frx":0A96
      Tab(3).ControlCount=   5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label16(1)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label15(1)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label14(1)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label13(1)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "sprDep(2)"
      Tab(3).Control(4).Enabled=   0   'False
      Begin FPSpread.vaSpread sprDep 
         Height          =   2880
         Index           =   2
         Left            =   -74910
         OleObjectBlob   =   "frmOnlineHora.frx":0AB2
         TabIndex        =   36
         Top             =   825
         Width           =   8775
      End
      Begin FPSpread.vaSpread sprTotCIA 
         Height          =   2880
         Left            =   150
         OleObjectBlob   =   "frmOnlineHora.frx":1174
         TabIndex        =   13
         Top             =   825
         Width           =   8775
      End
      Begin FPSpread.vaSpread sprDep 
         Height          =   2880
         Index           =   0
         Left            =   -74910
         OleObjectBlob   =   "frmOnlineHora.frx":1836
         TabIndex        =   26
         Top             =   810
         Width           =   8775
      End
      Begin FPSpread.vaSpread sprDep 
         Height          =   2880
         Index           =   1
         Left            =   -74910
         OleObjectBlob   =   "frmOnlineHora.frx":1EF8
         TabIndex        =   31
         Top             =   825
         Width           =   8775
      End
      Begin VB.CommandButton botExcelEzeA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -74835
         Picture         =   "frmOnlineHora.frx":25BA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   330
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton botExcelEzeb 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -74820
         Picture         =   "frmOnlineHora.frx":2B4C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   330
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton botExcel 
         Caption         =   "Excel"
         Height          =   510
         Left            =   195
         Picture         =   "frmOnlineHora.frx":30DE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   315
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual - Estim."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   -68265
         TabIndex        =   40
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   -70200
         TabIndex        =   39
         Top             =   525
         Width           =   1950
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estimado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   -71250
         TabIndex        =   38
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Elegida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   -73215
         TabIndex        =   37
         Top             =   525
         Width           =   1980
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Elegida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   -73230
         TabIndex        =   35
         Top             =   525
         Width           =   1980
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estimado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   -71265
         TabIndex        =   34
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   -70215
         TabIndex        =   33
         Top             =   525
         Width           =   1950
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual - Estim."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   -68280
         TabIndex        =   32
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Elegida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -73230
         TabIndex        =   30
         Top             =   510
         Width           =   1980
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estimado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -71265
         TabIndex        =   29
         Top             =   510
         Width           =   1065
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -70215
         TabIndex        =   28
         Top             =   510
         Width           =   1950
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual - Estim."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -68280
         TabIndex        =   27
         Top             =   510
         Width           =   1800
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual - Estim."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6780
         TabIndex        =   24
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4845
         TabIndex        =   23
         Top             =   525
         Width           =   1950
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estimado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3795
         TabIndex        =   15
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Elegida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1830
         TabIndex        =   14
         Top             =   525
         Width           =   1980
      End
   End
   Begin VB.Frame frdatos 
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   45
      TabIndex        =   0
      Top             =   -30
      Width           =   9000
      Begin VB.Frame Frame1 
         Caption         =   "Fecha de  Consultas"
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
         Height          =   1230
         Left            =   120
         TabIndex        =   3
         Top             =   60
         Width           =   8940
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
            Left            =   5250
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   330
            Width           =   1080
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
            Left            =   5250
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   690
            Width           =   1080
         End
         Begin VB.CommandButton botHelpFD 
            Height          =   345
            Left            =   3105
            Picture         =   "frmOnlineHora.frx":3670
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   510
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskFDesde 
            Height          =   285
            Left            =   1545
            TabIndex        =   9
            Top             =   525
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "##-##-####"
            PromptChar      =   " "
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
            Index           =   1
            Left            =   3930
            TabIndex        =   22
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label6 
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
            Left            =   3915
            TabIndex        =   21
            Top             =   735
            Width           =   1290
         End
         Begin VB.Label Label4 
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
            Left            =   6450
            TabIndex        =   20
            Top             =   330
            Width           =   1440
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
            Left            =   6540
            TabIndex        =   19
            Top             =   750
            Width           =   330
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
            Left            =   6990
            TabIndex        =   18
            Top             =   795
            Width           =   735
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
            Left            =   300
            TabIndex        =   10
            Top             =   525
            Width           =   1185
         End
      End
   End
End
Attribute VB_Name = "frmOnlineHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset

Dim estEze As Single
Dim estAep As Single
Dim estTot As Single
Dim estBar As Single
Dim Actualizar As Boolean

Private Sub l_AplicarParticipacionEstimado(ByRef spr As control, Total As Single)
Dim i
Dim valor As Variant

For i = 1 To spr.MaxRows
    spr.GetText 4, i, valor
    
    spr.SetText 5, i, str(valor * Total / 100)
    
Next

For i = 1 To spr.MaxRows
    spr.GetText 6, i, valor
    If valor <> 0 Then
        spr.SetText 7, i, str(valor / Total * 100)
    'Exit For
    End If
        
Next

End Sub

Private Function L_OtenerEstimado()
Dim sql As String
Dim rs As Recordset

'sql = " select depn,"
sql = " select decode(depn,'INT','BAR',DEPN) DEPN,"
sql = sql & " Sum (Round(Porc, 2)) est"
sql = sql & " From estadis.porcentaje_d"
sql = sql & " Where anio = " & Format(Now, "YYYY") & " And Mes = " & Format(Now, "mm")
sql = sql & " and tipo_porc = 'I' and nivel = 0 "
sql = sql & " and ( depn <> 'INT' or sdep = 'BARI') "
sql = sql & " and dia = " & Func.func_ToDate(Format(Now, FTOFECHA))
sql = sql & " group by depn"

If Aplicacion.ObtenerRsDAO(sql, rs) Then
    
    Do While Not rs.EOF
        Select Case rs!depn
           Case "EZE"
                estEze = rs!est
           Case "AEP"
                estAep = rs!est
           Case "BAR"
                estBar = rs!est
        End Select
        rs.MoveNext
    Loop
    estTot = estEze + estAep
    
    rs.Close
End If


End Function

Private Sub L_TratarExcel(titulo As String, subTit As String, Esp As String, CantCol As Integer)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer, filaant As Integer
Dim i As Integer
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmVtaHora.caption = Aplicacion.SeteoProceso(frmVtaHora.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.application.Visible = True
    
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
        
    
    
    Select Case Esp
           
        Case "TOT"
            For i = 1 To CantCol
                sprTotCIA.GetText i, 0, tit
                titCol(i) = tit
            Next
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información Total  "
               Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
               
               fila = fila + 2
               
               Exl_BajarGrillaExel sprTotCIA, AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + 25
               rango = Exl_rangos(fila, fila, Col, sprTotCIA.MaxCols + 1)
               Exl_ColorInt AppExcel, rango, Exl_Gris
               Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
               Exl_LetraColor AppExcel, rango, Exl_Blanco
               
               Exl_AnchoCol AppExcel, 1, 1, 10
    End Select
    With AppExcel.Application.ActiveSheet.PageSetup
'        .PrintTitleRows = "$1:$" & Trim(str(filaTit))
        .CenterHorizontally = True
        .TopMargin = Exl_TopMargen
'        .BottomMargin = Exl_BotMargen
'        .CenterFooter = "Página &P de &N"
'        .PrintGridlines = False
    End With
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
        
    AppExcel.SaveAs NOMBRE & ".xls"
'    AppExcel.Workbooks.Open nombre & ".xls"
    Set AppExcel = Nothing
    
End If

ErrorExl:

    frmVtaHora.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub

            
Public Sub L_BajarGrillaExel(spr As control, apl As Object, FilaInit As Integer, ColInit As Integer, titCol() As String)
Dim gr_fila As Integer
Dim gr_col As Integer, gr_aux_col As Integer
Dim dato As Variant, hora As Integer
Dim rango As String, RangoArmado As String

hora = 0
gr_fila = 1

      Exl_PonerValor apl, FilaInit, 1, Format(mskFDesde.FormattedText, "dddd")
      Exl_PonerValor apl, FilaInit, 2, "Rango"
      Exl_PonerValor apl, FilaInit, 3, "Dotación Teórica"
      Exl_PonerValor apl, FilaInit, 4, "Dotación Real"
      Exl_PonerValor apl, FilaInit, 5, "Importes"
      Exl_PonerValor apl, FilaInit, 6, "Ticket"
      Exl_PonerValor apl, FilaInit, 7, "Vuelos/Bodega"
      Exl_PonerValor apl, FilaInit, 8, "Promedio Ticket"
      Exl_PonerValor apl, FilaInit, 9, "Productividad x Pers"

'For gr_fila = 0 To spr.MaxRows
 For hora = 0 To 23
    If hora = 0 Then
    RangoArmado = Format$(hora, "0") & " - " & Format$(hora + 1, "00")
    Else
    RangoArmado = Format$(hora, "00") & " - " & Format$(hora + 1, "00")
    End If
    
    spr.GetText 2, gr_fila, dato
        
    If dato = RangoArmado Then
    Exl_PonerValor apl, FilaInit + hora + 1, 2, dato
    
    spr.GetText 1, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 1, dato
        
    spr.GetText 3, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 5, dato
    
    spr.GetText 4, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 6, dato
        
    spr.GetText 5, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 8, dato
        
        If gr_fila <= spr.MaxRows Then
            gr_fila = gr_fila + 1
        End If
    Else
        Exl_PonerValor apl, FilaInit + hora + 1, 1, Format(mskFDesde.FormattedText, "dd-mm-yy")
        Exl_PonerValor apl, FilaInit + hora + 1, 2, RangoArmado
    End If
Next
    Exl_PonerValor apl, FilaInit + hora + 1, 1, "Total"
    spr.GetText 3, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 5, dato
    
    spr.GetText 4, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 6, dato

rango = Exl_rangos(FilaInit, FilaInit, ColInit, spr.MaxCols + 1)

Exl_Justificacion apl, rango, Exl_CentroCol, Exl_CentroVert, False

Exl_ColorInt apl, rango, Exl_Gris
Exl_Letra apl, rango, NEGRITA, 10, "Ms Serif"
Exl_LetraColor apl, rango, Exl_Blanco
rango = Exl_rangos(FilaInit, 25 + FilaInit, ColInit, spr.MaxCols + 1)

Exl_Lineas apl, rango, Exl_Linsimple

rango = Exl_rangos(FilaInit + 6, FilaInit + 6, ColInit, spr.MaxCols + 1)
apl.Application.Range(rango).Borders(4).Weight = 4
rango = Exl_rangos(FilaInit + 14, FilaInit + 14, ColInit, spr.MaxCols + 1)
apl.Application.Range(rango).Borders(4).Weight = 4

rango = Exl_rangos(FilaInit + 1, FilaInit + 25, 5, 5)
apl.Application.Range(rango).NumberFormat = "$ #,##0_ ;[Red]-#,##0 "
rango = Exl_rangos(FilaInit + 1, FilaInit + 25, 8, 9)
apl.Application.Range(rango).NumberFormat = "$ #,##0_ ;[Red]-#,##0 "

End Sub

            


Private Sub L_LimpiarGrillas()
Dim i
       

sprDep(0).MaxRows = 0
sprDep(1).MaxRows = 0
sprDep(2).MaxRows = 0

sprTotCIA.MaxRows = 0

End Sub

Private Sub L_PonerenGrilla(spr As control, item As String, valor As Single)
    spr.MaxRows = spr.MaxRows + 1
    spr.SetText 1, spr.MaxRows, Trim(item)
    spr.SetText 2, spr.MaxRows, str(valor)
End Sub

Private Sub L_Refrescar()
Dim sql As String

'On Error GoTo ErrInd:
sql = " SELECT cod_depn, rango, sum(imp) imp,sum(impOL) impOL,  fch"
sql = sql & " from("
sql = sql & " SELECT DECODE(cod_depn,'INT',COD_SDEP,COD_DEPN) cod_depn, rango, sum(imp) imp,0 impOL,  fch_proceso fch"
    If Month(Date) = Month(mskFDesde.FormattedText) And Year(Date) = Year(mskFDesde.FormattedText) Then
        sql = sql & " FROM estadis.View_Ticket_hora "
    Else
        sql = sql & " FROM estadis.View_Ticket_hora_hist "
    End If
sql = sql & " WHERE fch_ticket = " & Func.func_ToDate(mskFDesde.FormattedText)
sql = sql & " GROUP BY fch_proceso,DECODE(cod_depn,'INT',COD_SDEP,COD_DEPN),rango"
sql = sql & " Union"
sql = sql & " SELECT 'TOT' cod_depn, rango  , sum(imp) imp ,0 ,  fch_proceso"
    If Month(Date) = Month(mskFDesde.FormattedText) And Year(Date) = Year(mskFDesde.FormattedText) Then
        sql = sql & " FROM estadis.View_Ticket_hora "
    Else
        sql = sql & " FROM estadis.View_Ticket_hora_hist "
    End If
sql = sql & " WHERE fch_ticket = " & Func.func_ToDate(mskFDesde.FormattedText)
sql = sql & " GROUP BY fch_proceso,rango"
sql = sql & " Union"
sql = sql & " SELECT DECODE(cod_depn,'INT',COD_SDEP,COD_DEPN) cod_depn, rango,0, sum(imp) imp, " & Func.func_ToDate(mskFDesde.FormattedText) & "  fch"
sql = sql & " From estadis.View_Ticket_hora_dia"
sql = sql & " GROUP BY fch_proceso,DECODE(cod_depn,'INT',COD_SDEP,COD_DEPN),rango"
sql = sql & " Union"
sql = sql & " SELECT 'TOT' cod_depn, rango,0, sum(imp) imp," & Func.func_ToDate(mskFDesde.FormattedText) & "  fch"
sql = sql & " FROM estadis.View_Ticket_hora_dia  GROUP BY fch_proceso,rango)"
sql = sql & " GROUP BY cod_depn,fch,rango"
sql = sql & " ORDER BY cod_depn,fch,rango"

frmOnlineHora.caption = Aplicacion.SeteoProceso(frmOnlineHora.caption)

    
If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        L_DecoEspigon
                
    End If
    Aplicacion.CerrarDAO RsData
    
    'L_TratarTotales

End If

ErrInd:
    frmOnlineHora.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub


Private Sub L_Resaltar()
Dim fila As Integer
Dim i

    For i = 0 To 2
        For fila = 1 To sprDep(i).MaxRows
        Spread.spread_ResaltarCelda sprDep(i), 8, fila
        Spread.spread_ResaltarCelda sprDep(i), 9, fila
        Next
    Next
        For fila = 1 To sprTotCIA.MaxRows
        Spread.spread_ResaltarCelda sprTotCIA, 8, fila
        Spread.spread_ResaltarCelda sprTotCIA, 9, fila
        Next

End Sub


Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant
Dim ImpOL As Single, ImpOLEze As Single, ImpOLAep As Single, ImpOLBARI As Single
Dim Importe As Single, ImporteEze As Single, ImporteAep As Single, ImporteBARI As Single

ImpOL = 0
Importe = 0
ImpOLEze = 0
ImporteEze = 0
ImpOLAep = 0
ImporteAep = 0
ImpOLBARI = 0
ImporteBARI = 0


Do While Not RsData.EOF
    Select Case RsData!cod_depn
      Case "TOT"
        sprTotCIA.MaxRows = sprTotCIA.MaxRows + 1
        sprTotCIA.SetText 1, sprTotCIA.MaxRows, Format$(RsData!fch, "dd-mm-yy")
        sprTotCIA.SetText 2, sprTotCIA.MaxRows, Trim(RsData!rango)
        Importe = Importe + RsData!imp
        sprTotCIA.SetText 3, sprTotCIA.MaxRows, str(Importe)
        If RsData!ImpOL <> 0 Then
           ImpOL = ImpOL + RsData!ImpOL
           sprTotCIA.SetText 6, sprTotCIA.MaxRows, str(ImpOL)
        End If
        'Spread_TotalesGrillaAcum sprTotCIA, 3, 4, sprTotCIA.MaxRows
      Case "EZE"
        sprDep(0).MaxRows = sprDep(0).MaxRows + 1
        sprDep(0).SetText 1, sprDep(0).MaxRows, Format$(RsData!fch, "dd-mm-yy")
        sprDep(0).SetText 2, sprDep(0).MaxRows, Trim(RsData!rango)
        ImporteEze = ImporteEze + RsData!imp
        sprDep(0).SetText 3, sprDep(0).MaxRows, str(ImporteEze)
        If RsData!ImpOL <> 0 Then
           ImpOLEze = ImpOLEze + RsData!ImpOL
           sprDep(0).SetText 6, sprDep(0).MaxRows, str(ImpOLEze)
        End If
      Case "AEP"
        sprDep(1).MaxRows = sprDep(1).MaxRows + 1
        sprDep(1).SetText 1, sprDep(1).MaxRows, Format$(RsData!fch, "dd-mm-yy")
        sprDep(1).SetText 2, sprDep(1).MaxRows, Trim(RsData!rango)
        ImporteAep = ImporteAep + RsData!imp
        sprDep(1).SetText 3, sprDep(1).MaxRows, str(ImporteAep)
        If RsData!ImpOL <> 0 Then
           ImpOLAep = ImpOLAep + RsData!ImpOL
           sprDep(1).SetText 6, sprDep(1).MaxRows, str(ImpOLAep)
        End If
      Case "BARI"
        sprDep(2).MaxRows = sprDep(2).MaxRows + 1
        sprDep(2).SetText 1, sprDep(2).MaxRows, Format$(RsData!fch, "dd-mm-yy")
        sprDep(2).SetText 2, sprDep(2).MaxRows, Trim(RsData!rango)
        ImporteBARI = ImporteBARI + RsData!imp
        sprDep(2).SetText 3, sprDep(2).MaxRows, str(ImporteBARI)
        If RsData!ImpOL <> 0 Then
           ImpOLBARI = ImpOLBARI + RsData!ImpOL
           sprDep(2).SetText 6, sprDep(2).MaxRows, str(ImpOLBARI)
        End If
                                                        
    End Select
    RsData.MoveNext
Loop
L_Participacion sprTotCIA
L_Participacion sprDep(0)
L_Participacion sprDep(1)
L_Participacion sprDep(2)

l_AplicarParticipacionEstimado sprTotCIA, estTot
l_AplicarParticipacionEstimado sprDep(0), estEze
l_AplicarParticipacionEstimado sprDep(1), estAep
l_AplicarParticipacionEstimado sprDep(2), estBar

L_Resaltar

'For i = 0 To 3
'    Spread_TotalesGrillas sprEzeA(i), sprEzeA(i).MaxCols - 5, 2
'Next
'For i = 0 To 2
'    Spread_TotalesGrillas sprEzeB(i), sprEzeB(i).MaxCols - 5, 2
'Next
'For i = 0 To 1
 '   Spread_TotalesGrillas sprAep(i), sprAep(i).MaxCols - 5, 2
'Next


End Sub

Private Sub L_Participacion(ByRef spr As control)
Dim i As Integer
Dim valor As Variant
Dim tot As Variant, fecha As Variant
Dim Result As Double

On Error GoTo ErrPorc:

spr.GetText 3, spr.MaxRows, tot

For i = 1 To spr.MaxRows
    spr.GetText 3, i, valor
    
    spr.SetText 4, i, str(valor / tot * 100)
Next


ErrPorc:
    Exit Sub
End Sub



Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        L_LimpiarGrillas
        txtTmpC.Text = Format$(Now, "hh:mm:ss")
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

    'L_TratarExcel " TOTAL BUENOS AIRES ", "Informe de Ventas Horarias del día (" & mskFDesde.FormattedText & " )", "TOT", sprTotCIA.MaxCols

End Sub

Private Sub botExcelEzeA_Click()
'    L_TratarExcel "ESPIGON INTERNACIONAL 'A'", "Informe de Ventas Horarias del día (" & mskFDesde.FormattedText & " )", EZEA, sprTot(0).MaxCols

End Sub

Private Sub botExcelEzeb_Click()
'    L_TratarExcel "ESPIGON INTERNACIONAL 'B'", "Informe de Ventas Horarias del día (" & mskFDesde.FormattedText & " )", EZEB, sprTot(1).MaxCols
    
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


Private Sub Form_Activate()
FuncLocal_SeteoTABS tabEspigon
End Sub

Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6000
Me.Width = 9300

'If Day(Date) = 1 Then
'    mskFDesde.Text = Format$(Func.func_Dia1SegunMes_Anio(Month(Date - 1), Year(Date - 1)), FTOFECHA)
'Else
'    mskFDesde.Text = "01-" & Month(Date) & "-" & Format$(Year(Date), "####")
'End If
mskFDesde.Text = Format$(Date - 7, FTOFECHA)

Actualizar = True

L_LimpiarGrillas

L_OtenerEstimado
txtTmpC.Text = Format$(Now, "hh:mm:ss")

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

End Sub





Private Sub sldTmp_Click()
    labMin.caption = sldTmp.Value
End Sub


Private Sub tabEspigon_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    Select Case tabEspigon.Tab
        Case 0
            'sprTotal.SetFocus
        Case 1
            'sprEzeA.SetFocus
        Case 2
            'sprEzeB.SetFocus
        Case 3
            'sprAep.SetFocus
    End Select
    
    
ErrT:
    Exit Sub



End Sub

Private Sub Timer1_Timer()
txtTmpA.Text = Format$(Now, "hh:mm:ss")

If Actualizar Then
    If ((Format$(txtTmpA.Text, "hhmmss")) - (Format$(txtTmpC.Text, "hhmmss"))) > (sldTmp.Value * 100) Then
        Call botEjecutar_Click(0)
    End If
End If

End Sub


