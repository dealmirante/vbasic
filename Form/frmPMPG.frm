VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmPMPG 
   Caption         =   "Ventas por Proveedor-Marca-Producto"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   9195
   Begin VB.Frame frdatos 
      Height          =   1575
      Left            =   105
      TabIndex        =   22
      Top             =   -75
      Width           =   8025
      Begin ComctlLib.Toolbar Toolbar3 
         Height          =   390
         Left            =   1080
         TabIndex        =   46
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         ButtonWidth     =   635
         ButtonHeight    =   582
         ImageList       =   "ImageList2"
         _Version        =   327680
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   3
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "a"
               Object.ToolTipText     =   "Ayuda Marcas"
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "B"
               Object.ToolTipText     =   "Limpiar Prov."
               Object.Tag             =   ""
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin ComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   7035
         TabIndex        =   50
         Top             =   615
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         ButtonWidth     =   635
         ButtonHeight    =   582
         ImageList       =   "ImageList2"
         _Version        =   327680
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   3
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Enabled         =   0   'False
               Key             =   "a"
               Object.ToolTipText     =   "Ayuda Producto"
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "B"
               Object.ToolTipText     =   "Limpiar Prov."
               Object.Tag             =   ""
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   1080
         TabIndex        =   47
         Top             =   630
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         ButtonWidth     =   635
         ButtonHeight    =   582
         ImageList       =   "ImageList2"
         _Version        =   327680
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   3
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "a"
               Object.ToolTipText     =   "Ayuda Provvedores"
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "B"
               Object.ToolTipText     =   "Limpiar Prov."
               Object.Tag             =   ""
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   6315
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   225
         Width           =   1530
      End
      Begin VB.TextBox txtDescProd 
         Height          =   315
         Left            =   5100
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1095
         Width           =   2790
      End
      Begin VB.TextBox txtCodProd 
         BackColor       =   &H80000018&
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
         Left            =   6105
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   660
         Width           =   885
      End
      Begin VB.TextBox txtCodMr 
         Height          =   300
         Left            =   1185
         TabIndex        =   30
         Top             =   1110
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.TextBox txtCodPv 
         BackColor       =   &H80000018&
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
         Left            =   1170
         TabIndex        =   29
         Top             =   690
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   5025
         Picture         =   "frmPMPG.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2325
         Picture         =   "frmPMPG.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   210
         Width           =   375
      End
      Begin VB.TextBox txtDescrip 
         Height          =   330
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   675
         Width           =   2880
      End
      Begin VB.TextBox txtDescripMr 
         Height          =   315
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1110
         Width           =   2880
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1230
         TabIndex        =   27
         Top             =   225
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFHasta 
         Height          =   285
         Left            =   3945
         TabIndex        =   28
         Top             =   225
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.Label LblEspigon 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rubro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   5520
         TabIndex        =   67
         Top             =   225
         Width           =   780
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
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
         Left            =   5100
         TabIndex        =   48
         Top             =   660
         Width           =   960
      End
      Begin ComctlLib.ImageList ImageList2 
         Left            =   7410
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327680
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   11
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmPMPG.frx":02E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmPMPG.frx":03F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmPMPG.frx":0508
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmPMPG.frx":0D3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmPMPG.frx":12B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmPMPG.frx":184A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmPMPG.frx":1B64
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmPMPG.frx":1E7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmPMPG.frx":2198
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmPMPG.frx":24B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmPMPG.frx":2CE4
               Key             =   ""
            EndProperty
         EndProperty
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
         Left            =   2835
         TabIndex        =   34
         Top             =   225
         Width           =   1080
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
         Left            =   90
         TabIndex        =   33
         Top             =   225
         Width           =   1125
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
         Height          =   315
         Index           =   3
         Left            =   105
         TabIndex        =   32
         Top             =   645
         Width           =   960
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Marca"
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
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   1110
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1560
      Left            =   8145
      TabIndex        =   20
      Top             =   -60
      Width           =   810
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
         Left            =   120
         Picture         =   "frmPMPG.frx":2FFE
         Style           =   1  'Graphical
         TabIndex        =   53
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
         Left            =   120
         Picture         =   "frmPMPG.frx":3100
         Style           =   1  'Graphical
         TabIndex        =   52
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
         Height          =   450
         Index           =   2
         Left            =   120
         Picture         =   "frmPMPG.frx":3202
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1005
         Width           =   570
      End
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   3855
      Left            =   60
      TabIndex        =   0
      Top             =   1560
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   6800
      _Version        =   327680
      Tabs            =   5
      Tab             =   1
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
      TabPicture(0)   =   "frmPMPG.frx":3A24
      Tab(0).ControlCount=   5
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "botPorcT(2)"
      Tab(0).Control(0).Enabled=   -1  'True
      Tab(0).Control(1)=   "botPorcT(1)"
      Tab(0).Control(1).Enabled=   -1  'True
      Tab(0).Control(2)=   "botPorcT(0)"
      Tab(0).Control(2).Enabled=   -1  'True
      Tab(0).Control(3)=   "botExcelTot"
      Tab(0).Control(3).Enabled=   -1  'True
      Tab(0).Control(4)=   "tabTot"
      Tab(0).Control(4).Enabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmPMPG.frx":3A40
      Tab(1).ControlCount=   6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "tabEzeA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "botExcelEzeA"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "botPorc(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "botPorc(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "botPorc(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "botGrTortaEA"
      Tab(1).Control(5).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmPMPG.frx":3A5C
      Tab(2).ControlCount=   6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabEzeB"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "botExcelEzeB"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "botPorcB(2)"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "botPorcB(1)"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "botPorcB(0)"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "botGrTortaEB"
      Tab(2).Control(5).Enabled=   -1  'True
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmPMPG.frx":3A78
      Tab(3).ControlCount=   6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "botGrTortaAEP"
      Tab(3).Control(0).Enabled=   -1  'True
      Tab(3).Control(1)=   "botPorcAep(0)"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "botPorcAep(1)"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "botPorcAep(2)"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "botExcelAep"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "tabAep"
      Tab(3).Control(5).Enabled=   0   'False
      TabCaption(4)   =   "INTERIOR"
      TabPicture(4)   =   "frmPMPG.frx":3A94
      Tab(4).ControlCount=   6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "botExcelInt"
      Tab(4).Control(0).Enabled=   -1  'True
      Tab(4).Control(1)=   "botPorcInt(2)"
      Tab(4).Control(1).Enabled=   -1  'True
      Tab(4).Control(2)=   "botPorcInt(1)"
      Tab(4).Control(2).Enabled=   -1  'True
      Tab(4).Control(3)=   "botPorcInt(0)"
      Tab(4).Control(3).Enabled=   -1  'True
      Tab(4).Control(4)=   "botGrTortaINTE"
      Tab(4).Control(4).Enabled=   -1  'True
      Tab(4).Control(5)=   "SSTab1"
      Tab(4).Control(5).Enabled=   0   'False
      Begin VB.CommandButton botPorcT 
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
         Left            =   -67425
         Picture         =   "frmPMPG.frx":3AB0
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   1890
         Width           =   795
      End
      Begin VB.CommandButton botPorcT 
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
         Left            =   -67425
         Picture         =   "frmPMPG.frx":4176
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   1380
         Width           =   795
      End
      Begin VB.CommandButton botPorcT 
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
         Left            =   -67425
         Picture         =   "frmPMPG.frx":47CC
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   870
         Width           =   800
      End
      Begin VB.CommandButton botExcelTot 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67425
         Picture         =   "frmPMPG.frx":513E
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   2430
         Width           =   795
      End
      Begin TabDlg.SSTab tabTot 
         Height          =   3405
         Left            =   -74805
         TabIndex        =   82
         Top             =   375
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   6006
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   5
         TabHeight       =   520
         ForeColor       =   16711680
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "frmPMPG.frx":56D0
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1(0)"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmPMPG.frx":56EC
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame2(0)"
         Tab(1).Control(1).Enabled=   0   'False
         Begin VB.Frame Frame1 
            Height          =   2985
            Index           =   0
            Left            =   135
            TabIndex        =   83
            Top             =   45
            Width           =   6825
            Begin FPSpread.vaSpread sprTotImp 
               Height          =   2700
               Left            =   105
               OleObjectBlob   =   "frmPMPG.frx":5708
               TabIndex        =   84
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame Frame1 
            Height          =   2985
            Index           =   1
            Left            =   165
            TabIndex        =   93
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprTotPorc 
               Height          =   2700
               Index           =   0
               Left            =   105
               OleObjectBlob   =   "frmPMPG.frx":5BE5
               TabIndex        =   94
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame Frame2 
            Height          =   2970
            Index           =   0
            Left            =   -74895
            TabIndex        =   85
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprTotUnid 
               Height          =   2700
               Left            =   105
               OleObjectBlob   =   "frmPMPG.frx":5F9A
               TabIndex        =   86
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame Frame2 
            Height          =   2985
            Index           =   1
            Left            =   -74910
            TabIndex        =   91
            Top             =   30
            Width           =   6810
            Begin FPSpread.vaSpread sprTotPorc 
               Height          =   2700
               Index           =   1
               Left            =   105
               OleObjectBlob   =   "frmPMPG.frx":6477
               TabIndex        =   92
               Top             =   195
               Width           =   6675
            End
         End
      End
      Begin VB.CommandButton botExcelInt 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67500
         Picture         =   "frmPMPG.frx":682C
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   2490
         Width           =   795
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
         Left            =   -67500
         Picture         =   "frmPMPG.frx":6DBE
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1500
         Width           =   795
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
         Index           =   1
         Left            =   -67500
         Picture         =   "frmPMPG.frx":7484
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1005
         Width           =   795
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
         Left            =   -67500
         Picture         =   "frmPMPG.frx":7ADA
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   500
         Width           =   800
      End
      Begin VB.CommandButton botGrTortaINTE 
         Height          =   495
         Left            =   -67500
         Picture         =   "frmPMPG.frx":844C
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   1995
         Width           =   795
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3350
         Left            =   -74850
         TabIndex        =   68
         Top             =   400
         Width           =   7000
         _ExtentX        =   12356
         _ExtentY        =   5900
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
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
         TabPicture(0)   =   "frmPMPG.frx":888E
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frIntPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frInt0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmPMPG.frx":88AA
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frInt1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frIntPorc1"
         Tab(1).Control(1).Enabled=   0   'False
         Begin VB.Frame frInt0 
            Height          =   3000
            Left            =   100
            TabIndex        =   80
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprInt 
               Height          =   2700
               Index           =   0
               Left            =   100
               OleObjectBlob   =   "frmPMPG.frx":88C6
               TabIndex        =   81
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frIntPorc0 
            Height          =   3000
            Left            =   105
            TabIndex        =   76
            Top             =   45
            Width           =   6800
            Begin FPSpread.vaSpread sprIntPorc 
               Height          =   2700
               Index           =   0
               Left            =   90
               OleObjectBlob   =   "frmPMPG.frx":8DA8
               TabIndex        =   79
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frInt1 
            Height          =   3000
            Left            =   -74900
            TabIndex        =   74
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprInt 
               Height          =   2700
               Index           =   1
               Left            =   100
               OleObjectBlob   =   "frmPMPG.frx":9212
               TabIndex        =   75
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frIntPorc1 
            Height          =   3000
            Left            =   -74900
            TabIndex        =   77
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprIntPorc 
               Height          =   2700
               Index           =   1
               Left            =   100
               OleObjectBlob   =   "frmPMPG.frx":96BC
               TabIndex        =   78
               Top             =   200
               Width           =   6600
            End
         End
      End
      Begin VB.CommandButton botGrTortaAEP 
         Height          =   495
         Left            =   -67500
         Picture         =   "frmPMPG.frx":9B10
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   2025
         Width           =   795
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
         Picture         =   "frmPMPG.frx":9F52
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   500
         Width           =   800
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
         Picture         =   "frmPMPG.frx":A8C4
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   1005
         Width           =   795
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
         Picture         =   "frmPMPG.frx":AF1A
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1515
         Width           =   795
      End
      Begin VB.CommandButton botGrTortaEB 
         Height          =   495
         Left            =   -67500
         Picture         =   "frmPMPG.frx":B5E0
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   2025
         Width           =   795
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
         Left            =   -67500
         Picture         =   "frmPMPG.frx":BA22
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   500
         Width           =   800
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
         Left            =   -67500
         Picture         =   "frmPMPG.frx":C394
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1005
         Width           =   795
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
         Left            =   -67500
         Picture         =   "frmPMPG.frx":C9EA
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1515
         Width           =   795
      End
      Begin VB.CommandButton botGrTortaEA 
         Height          =   495
         Left            =   7500
         Picture         =   "frmPMPG.frx":D0B0
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2025
         Width           =   780
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
         Picture         =   "frmPMPG.frx":D4F2
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   500
         Width           =   800
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
         Height          =   495
         Index           =   1
         Left            =   7500
         Picture         =   "frmPMPG.frx":DE64
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1005
         Width           =   795
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
         Height          =   495
         Index           =   2
         Left            =   7500
         Picture         =   "frmPMPG.frx":E4BA
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1515
         Width           =   795
      End
      Begin VB.CommandButton botExcelAep 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67500
         Picture         =   "frmPMPG.frx":EB80
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2535
         Width           =   795
      End
      Begin VB.CommandButton botExcelEzeB 
         Caption         =   "Excel"
         Height          =   525
         Left            =   -67500
         Picture         =   "frmPMPG.frx":F112
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2535
         Width           =   795
      End
      Begin VB.CommandButton botExcelEzeA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   7500
         Picture         =   "frmPMPG.frx":F6A4
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2535
         Width           =   780
      End
      Begin TabDlg.SSTab tabEzeB 
         Height          =   3350
         Left            =   -74850
         TabIndex        =   1
         Top             =   400
         Width           =   7000
         _ExtentX        =   12356
         _ExtentY        =   5900
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
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
         TabPicture(0)   =   "frmPMPG.frx":FC36
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frEzeBPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frEzeB0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmPMPG.frx":FC52
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeBPorc1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frEzeB1"
         Tab(1).Control(1).Enabled=   0   'False
         Begin VB.Frame frEzeB1 
            Height          =   3000
            Left            =   -74900
            TabIndex        =   5
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   1
               Left            =   100
               OleObjectBlob   =   "frmPMPG.frx":FC6E
               TabIndex        =   44
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeBPorc1 
            Height          =   3000
            Left            =   -74900
            TabIndex        =   8
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   1
               Left            =   100
               OleObjectBlob   =   "frmPMPG.frx":10150
               TabIndex        =   15
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeB0 
            Height          =   3000
            Left            =   100
            TabIndex        =   3
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeB 
               Height          =   2700
               Index           =   0
               Left            =   100
               OleObjectBlob   =   "frmPMPG.frx":1050A
               TabIndex        =   11
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeBPorc0 
            Height          =   3000
            Left            =   100
            TabIndex        =   7
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeBPorc 
               Height          =   2700
               Index           =   0
               Left            =   100
               OleObjectBlob   =   "frmPMPG.frx":109EC
               TabIndex        =   13
               Top             =   200
               Width           =   6600
            End
         End
      End
      Begin TabDlg.SSTab tabEzeA 
         Height          =   3350
         Left            =   150
         TabIndex        =   35
         Top             =   400
         Width           =   7000
         _ExtentX        =   12356
         _ExtentY        =   5900
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
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
         TabPicture(0)   =   "frmPMPG.frx":10DA6
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frEzeAPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frEzeA0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmPMPG.frx":10DC2
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frEzeA1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frEzeAPorc1"
         Tab(1).Control(1).Enabled=   0   'False
         Begin VB.Frame frEzeA0 
            Height          =   2985
            Left            =   100
            TabIndex        =   36
            Top             =   30
            Width           =   6765
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   0
               Left            =   105
               OleObjectBlob   =   "frmPMPG.frx":10DDE
               TabIndex        =   37
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeAPorc0 
            Height          =   2985
            Left            =   90
            TabIndex        =   40
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeAPorc 
               Height          =   2700
               Index           =   0
               Left            =   105
               OleObjectBlob   =   "frmPMPG.frx":112BB
               TabIndex        =   41
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeA1 
            Height          =   3000
            Left            =   -74900
            TabIndex        =   38
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeA 
               Height          =   2700
               Index           =   1
               Left            =   105
               OleObjectBlob   =   "frmPMPG.frx":11670
               TabIndex        =   39
               Top             =   195
               Width           =   6600
            End
         End
         Begin VB.Frame frEzeAPorc1 
            Height          =   3000
            Left            =   -74900
            TabIndex        =   42
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprEzeAPorc 
               Height          =   2700
               Index           =   1
               Left            =   105
               OleObjectBlob   =   "frmPMPG.frx":11B4D
               TabIndex        =   43
               Top             =   195
               Width           =   6600
            End
         End
      End
      Begin TabDlg.SSTab tabAep 
         Height          =   3350
         Left            =   -74850
         TabIndex        =   2
         Top             =   400
         Width           =   7000
         _ExtentX        =   12356
         _ExtentY        =   5900
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
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
         TabPicture(0)   =   "frmPMPG.frx":11F02
         Tab(0).ControlCount=   2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frAepPorc0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frAep0"
         Tab(0).Control(1).Enabled=   0   'False
         TabCaption(1)   =   "Unidades"
         TabPicture(1)   =   "frmPMPG.frx":11F1E
         Tab(1).ControlCount=   2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frAep1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "frAepPorc1"
         Tab(1).Control(1).Enabled=   0   'False
         Begin VB.Frame frAep1 
            Height          =   3000
            Left            =   -74900
            TabIndex        =   6
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   1
               Left            =   100
               OleObjectBlob   =   "frmPMPG.frx":11F3A
               TabIndex        =   45
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frAepPorc1 
            Height          =   3000
            Left            =   -74900
            TabIndex        =   10
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   1
               Left            =   100
               OleObjectBlob   =   "frmPMPG.frx":123E4
               TabIndex        =   16
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frAep0 
            Height          =   3000
            Left            =   100
            TabIndex        =   4
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprAep 
               Height          =   2700
               Index           =   0
               Left            =   100
               OleObjectBlob   =   "frmPMPG.frx":12784
               TabIndex        =   12
               Top             =   200
               Width           =   6600
            End
         End
         Begin VB.Frame frAepPorc0 
            Height          =   3000
            Left            =   100
            TabIndex        =   9
            Top             =   30
            Width           =   6800
            Begin FPSpread.vaSpread sprAepPorc 
               Height          =   2700
               Index           =   0
               Left            =   100
               OleObjectBlob   =   "frmPMPG.frx":12C66
               TabIndex        =   14
               Top             =   200
               Width           =   6600
            End
         End
      End
   End
End
Attribute VB_Name = "frmPMPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsData  As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas
                       'con porcentajes
Dim UsoPorIntB As Boolean
Dim UsoPorAero As Boolean
Dim UsoPorInte As Boolean
Dim UsoPorTotal As Boolean

Dim HelpPROD As Boolean

Private Sub L_AcumularTotales()
Dim valorGA As Variant, valorGB As Variant, valorGC As Variant

'Para acumular importes
sprEzeA(0).GetText 2, sprEzeA(0).MaxRows, valorGA
sprEzeA(0).GetText 3, sprEzeA(0).MaxRows, valorGB
sprEzeA(0).GetText 4, sprEzeA(0).MaxRows, valorGC

sprTotImp.MaxRows = sprTotImp.MaxRows + 1
sprTotImp.SetText 1, sprTotImp.MaxRows, "EZEA"
sprTotImp.SetText 2, sprTotImp.MaxRows, valorGA
sprTotImp.SetText 3, sprTotImp.MaxRows, valorGB
sprTotImp.SetText 4, sprTotImp.MaxRows, valorGC

sprEzeB(0).GetText 2, sprEzeB(0).MaxRows, valorGA
sprEzeB(0).GetText 3, sprEzeB(0).MaxRows, valorGB
sprEzeB(0).GetText 4, sprEzeB(0).MaxRows, valorGC

sprTotImp.MaxRows = sprTotImp.MaxRows + 1
sprTotImp.SetText 1, sprTotImp.MaxRows, "EZEB"
sprTotImp.SetText 2, sprTotImp.MaxRows, valorGA
sprTotImp.SetText 3, sprTotImp.MaxRows, valorGB
sprTotImp.SetText 4, sprTotImp.MaxRows, valorGC

sprAEP(0).GetText 2, sprAEP(0).MaxRows, valorGA
sprAEP(0).GetText 3, sprAEP(0).MaxRows, valorGB
sprAEP(0).GetText 4, sprAEP(0).MaxRows, valorGC

sprTotImp.MaxRows = sprTotImp.MaxRows + 1
sprTotImp.SetText 1, sprTotImp.MaxRows, "AEP"
sprTotImp.SetText 2, sprTotImp.MaxRows, valorGA
sprTotImp.SetText 3, sprTotImp.MaxRows, valorGB
sprTotImp.SetText 4, sprTotImp.MaxRows, valorGC

sprInt(0).GetText 2, sprInt(0).MaxRows, valorGA
sprInt(0).GetText 3, sprInt(0).MaxRows, valorGB
sprInt(0).GetText 4, sprInt(0).MaxRows, valorGC

sprTotImp.MaxRows = sprTotImp.MaxRows + 1
sprTotImp.SetText 1, sprTotImp.MaxRows, "INT"
sprTotImp.SetText 2, sprTotImp.MaxRows, valorGA
sprTotImp.SetText 3, sprTotImp.MaxRows, valorGB
sprTotImp.SetText 4, sprTotImp.MaxRows, valorGC


'Para acumular unidades
sprEzeA(1).GetText 2, sprEzeA(1).MaxRows, valorGA
sprEzeA(1).GetText 3, sprEzeA(1).MaxRows, valorGB
sprEzeA(1).GetText 4, sprEzeA(1).MaxRows, valorGC

sprTotUnid.MaxRows = sprTotUnid.MaxRows + 1
sprTotUnid.SetText 1, sprTotUnid.MaxRows, "EZEA"
sprTotUnid.SetText 2, sprTotUnid.MaxRows, str(valorGA * 1)
sprTotUnid.SetText 3, sprTotUnid.MaxRows, str(valorGB * 1)
sprTotUnid.SetText 4, sprTotUnid.MaxRows, str(valorGC * 1)

sprEzeB(1).GetText 2, sprEzeB(1).MaxRows, valorGA
sprEzeB(1).GetText 3, sprEzeB(1).MaxRows, valorGB
sprEzeB(1).GetText 4, sprEzeB(1).MaxRows, valorGC

sprTotUnid.MaxRows = sprTotUnid.MaxRows + 1
sprTotUnid.SetText 1, sprTotUnid.MaxRows, "EZEB"
sprTotUnid.SetText 2, sprTotUnid.MaxRows, str(valorGA * 1)
sprTotUnid.SetText 3, sprTotUnid.MaxRows, str(valorGB * 1)
sprTotUnid.SetText 4, sprTotUnid.MaxRows, str(valorGC * 1)

sprAEP(1).GetText 2, sprAEP(1).MaxRows, valorGA
sprAEP(1).GetText 3, sprAEP(1).MaxRows, valorGB
sprAEP(1).GetText 4, sprAEP(1).MaxRows, valorGC

sprTotUnid.MaxRows = sprTotUnid.MaxRows + 1
sprTotUnid.SetText 1, sprTotUnid.MaxRows, "AEP"
sprTotUnid.SetText 2, sprTotUnid.MaxRows, str(valorGA * 1)
sprTotUnid.SetText 3, sprTotUnid.MaxRows, str(valorGB * 1)
sprTotUnid.SetText 4, sprTotUnid.MaxRows, str(valorGC * 1)

sprInt(1).GetText 2, sprInt(1).MaxRows, valorGA
sprInt(1).GetText 3, sprInt(1).MaxRows, valorGB
sprInt(1).GetText 4, sprInt(1).MaxRows, valorGC

sprTotUnid.MaxRows = sprTotUnid.MaxRows + 1
sprTotUnid.SetText 1, sprTotUnid.MaxRows, "INT"
sprTotUnid.SetText 2, sprTotUnid.MaxRows, str(valorGA * 1)
sprTotUnid.SetText 3, sprTotUnid.MaxRows, str(valorGB * 1)
sprTotUnid.SetText 4, sprTotUnid.MaxRows, str(valorGC * 1)


Spread_TotalesGrillas sprTotImp, sprTotImp.MaxCols - 1, 2
Spread_TotalesGrillas sprTotUnid, sprTotUnid.MaxCols - 1, 2

End Sub


Private Function L_Armarcondicion() As String
Dim Cond As String, condLoc As String
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant, i

fechaDesde = mskFDesde.FormattedText
fechaHasta = mskFHasta.FormattedText

Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

If txtCodPv.Text <> "" Then
    Cond = Cond & " and cod_prov = '" & txtCodPv.Text & "'"
End If

If txtCodMr.Text <> "" Then
    Cond = Cond & " and cod_marca = '" & txtCodMr.Text & "'"
End If

If txtCodProd.Text <> "" Then
    Cond = Cond & " and cod_PROD = " & txtCodProd.Text
End If

If cboRubro.Text <> "" Then
    Cond = Cond & " and cod_rubr = '" & cboRubro.Text & "'"
End If

L_Armarcondicion = Cond

End Function

Private Sub L_LimpiarGrillas()
Dim i

sprTotImp.MaxRows = 0
sprTotUnid.MaxRows = 0
For i = 0 To 1
    sprEzeA(i).MaxRows = 0
    sprEzeB(i).MaxRows = 0
    sprAEP(i).MaxRows = 0
    sprInt(i).MaxRows = 0
Next


End Sub

Private Sub L_Porcentages(ByRef sprPorc As control, sprDato As control, Orientacion As String)
Dim i As Integer, j As Integer
Dim valor As Variant
Dim tot As Variant, Locale As Variant
Dim Result As Double

On Error GoTo ErrPorc:

sprPorc.MaxRows = 0

Select Case Orientacion
    Case "FILAS"
        For i = 1 To sprDato.MaxRows
            sprDato.GetText sprDato.MaxCols, i, tot
            sprDato.GetText 1, i, Locale
            Result = 0
            sprPorc.MaxRows = i
            
            sprPorc.SetText 1, i, Locale
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
            sprDato.GetText 1, i, Locale
            Result = 0
            sprPorc.MaxRows = i
        
            sprPorc.SetText 1, i, Locale
            
            For j = 2 To sprDato.MaxCols
                Result = 0
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
    frEzeBPorc0.Refresh
    frEzeBPorc1.Refresh
    frAepPorc0.Refresh
    frAepPorc1.Refresh
    frEzeA0.Refresh
    frEzeA1.Refresh
    frEzeB0.Refresh
    frEzeB1.Refresh
    frAep0.Refresh
    frAep1.Refresh
    FrInt0.Refresh
    FrInt1.Refresh
    
    Frame1(0).Refresh
    Frame2(0).Refresh
    
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
        Case INTE
            UsoPorInte = True
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorAero Or UsoPorInte) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub

Private Sub L_SeteosGral()
        
    frdatos.Enabled = False
    botEjecutar(0).Enabled = False
    tabEspigon.Enabled = True

End Sub

Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim Col As Integer
Dim LocSloc As String

Do While Not RsData.EOF
        
        Select Case RsData!cod_depn
            Case DSLoc(1).Dep
                   LocSloc = RsData!cod_local '& "-" & RsData!cod_sloc
                   For i = 0 To 1
                        sprAEP(i).MaxRows = sprAEP(i).MaxRows + 1
                        sprAEP(i).SetText 1, sprAEP(i).MaxRows, LocSloc
                   Next
                   Do While LocSloc = RsData!cod_local '& "-" & RsData!cod_sloc
                        Select Case RsData!Grupo_Venta
                           Case "A"
                               Col = 2
                           Case "B"
                               Col = 3
                           Case "C"
                               Col = 4
                        End Select
                       sprAEP(0).SetText Col, sprAEP(0).MaxRows, str(RsData!imp)
                       sprAEP(1).SetText Col, sprAEP(1).MaxRows, str(RsData!cant)
                
                       RsData.MoveNext
                    
                       If RsData.EOF Then
                           Exit Do
                       End If
                   Loop

            Case DSLoc(2).Dep
            
                Select Case RsData!Cod_Sdep
                    Case DSLoc(2).Sdep
                            LocSloc = RsData!cod_local '& "-" & RsData!cod_sloc
                            For i = 0 To 1
                                sprEzeA(i).MaxRows = sprEzeA(i).MaxRows + 1
                                sprEzeA(i).SetText 1, sprEzeA(i).MaxRows, LocSloc
                            Next
                            Do While LocSloc = RsData!cod_local '& "-" & RsData!cod_sloc
                                Select Case RsData!Grupo_Venta
                                    Case "A"
                                        Col = 2
                                    Case "B"
                                        Col = 3
                                    Case "C"
                                        Col = 4
                                End Select
        
                                sprEzeA(0).SetText Col, sprEzeA(0).MaxRows, str(RsData!imp)
                                sprEzeA(1).SetText Col, sprEzeA(1).MaxRows, str(RsData!cant)
                                
                                RsData.MoveNext
                                
                                If RsData.EOF Then
                                    Exit Do
                                End If
                            Loop

                    Case DSLoc(3).Sdep
                            LocSloc = RsData!cod_local '& "-" & RsData!cod_sloc
                            For i = 0 To 1
                                sprEzeB(i).MaxRows = sprEzeB(i).MaxRows + 1
                                sprEzeB(i).SetText 1, sprEzeB(i).MaxRows, LocSloc
                            Next
                            
                            Do While LocSloc = RsData!cod_local '& "-" & RsData!cod_sloc
                                 Select Case RsData!Grupo_Venta
                                     Case "A"
                                        Col = 2
                                    Case "B"
                                        Col = 3
                                    Case "C"
                                        Col = 4
                                 End Select
                
                                 sprEzeB(0).SetText Col, sprEzeB(0).MaxRows, str(RsData!imp)
                                 sprEzeB(1).SetText Col, sprEzeB(1).MaxRows, str(RsData!cant)
                            
                                RsData.MoveNext
                                If RsData.EOF Then
                                    Exit Do
                                End If
                             Loop
                       End Select
        
          Case DSLoc(4).Dep  'Dependencia Interior
               LocSloc = RsData!cod_local '& "-" & RsData!cod_sloc
               For i = 0 To 1
                   sprInt(i).MaxRows = sprInt(i).MaxRows + 1
                   sprInt(i).SetText 1, sprInt(i).MaxRows, LocSloc & " " & RsData!Cod_Sdep
                Next
                            
                Do While LocSloc = RsData!cod_local '& "-" & RsData!cod_sloc
                      Select Case RsData!Grupo_Venta
                                 Case "A"
                                    Col = 2
                                 Case "B"
                                    Col = 3
                                 Case "C"
                                    Col = 4
                        End Select
                
                        sprInt(0).SetText Col, sprInt(0).MaxRows, str(RsData!imp)
                        sprInt(1).SetText Col, sprInt(1).MaxRows, str(RsData!cant)
                            
                         RsData.MoveNext
                         
                         If RsData.EOF Then
                             Exit Do
                         End If
                Loop
        
         End Select

        If RsData.EOF Then
            Exit Do
        End If

Loop

For i = 0 To 1
    Spread_TotalesGrillas sprEzeA(i), sprEzeA(i).MaxCols - 1, 2
    Spread_TotalesGrillas sprEzeB(i), sprEzeB(i).MaxCols - 1, 2
    Spread_TotalesGrillas sprAEP(i), sprAEP(i).MaxCols - 1, 2
    Spread_TotalesGrillas sprInt(i), sprInt(i).MaxCols - 1, 2
Next

L_AcumularTotales
L_RefrescarFrames


End Sub

Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        L_Refrescar
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


Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset

On Error GoTo ErrPMPG:

frmPMPG.caption = Aplicacion.SeteoProceso(frmPMPG.caption)

sql = " SELECT "
sql = sql & " grupo_venta,"
sql = sql & " cod_depn, "
sql = sql & " cod_sdep, "
sql = sql & " cod_local, "
'sql = sql & " cod_sloc, "
sql = sql & " sum(importe) imp, "
sql = sql & " sum(cantidad) cant "
sql = sql & "FROM " & funcLocal_Vista("venta_plg", Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion
sql = sql & " group by grupo_venta,cod_depn,cod_sdep,cod_local "
sql = sql & " order by cod_depn,cod_sdep,cod_local"


If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        L_DecoEspigon
        L_SeteosGral
        
    End If
    Aplicacion.CerrarDAO RsData
    
End If

ErrPMPG:
    frmPMPG.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub


Private Sub botExcelAep_Click()
Select Case TabAep.Tab
    Case 0
        L_TratarExcel sprAEP(0), "AEROPARQUE ", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Importes"
    Case 1
        L_TratarExcel sprAEP(1), "AEROPARQUE", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Unidades"
    Case 2
        L_TratarExcel sprAEP(2), "AEROPARQUE", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Pasajeros"
    Case 3
        L_TratarExcel sprAEP(3), "AEROPARQUE", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Promedios por Tickets"
    Case 4
        L_TratarExcel sprAEP(4), "AEROPARQUE", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Promedios por Pasajeros"
End Select

End Sub

Private Sub botExcelEzeA_Click()
Select Case TabEzeA.Tab
    Case 0
        L_TratarExcel sprEzeA(0), "ESPIGON INTERNACIONAL A", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Importes"
    Case 1
        L_TratarExcel sprEzeA(1), "ESPIGON INTERNACIONAL A", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Unidades"
    Case 2
        L_TratarExcel sprEzeA(2), "ESPIGON INTERNACIONAL A", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Pasajeros"
    Case 3
        L_TratarExcel sprEzeA(3), "ESPIGON INTERNACIONAL A", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Promedios por Tickets"
    Case 4
        L_TratarExcel sprEzeA(4), "ESPIGON INTERNACIONAL A", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")", "Promedios por Pasajeros"
End Select

End Sub

Private Sub botExcelEzeb_Click()
Select Case TabEzeB.Tab
    Case 0
        L_TratarExcel sprEzeB(0), "ESPIGON INTERNACIONAL B", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Importes"
    Case 1
        L_TratarExcel sprEzeB(1), "ESPIGON INTERNACIONAL B", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Unidades"
    Case 2
        L_TratarExcel sprEzeB(2), "ESPIGON INTERNACIONAL B", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Pasajeros"
    Case 3
        L_TratarExcel sprEzeB(3), "ESPIGON INTERNACIONAL B", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Promedios por Tickets"
    Case 4
        L_TratarExcel sprEzeB(4), "ESPIGON INTERNACIONAL B", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Promedios por Pasajeros"
End Select

End Sub

Private Sub L_TratarExcel(spr As control, titulo As String, subTit As String, Info As String)
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

frmPMPG.caption = Aplicacion.SeteoProceso(frmPMPG.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
'    AppExcel.application.Visible = True
    
    ReDim titCol(spr.MaxCols)
    Col = 1
    fila = 4
    
    For i = 1 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    'Titulo
    Exl_PonerValor AppExcel, 2, 1, titulo
    rango = Exl_rangos(2, 2, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    'SubTitulo
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    fila = fila + 2
    rango = Exl_rangos(fila, fila + 5, 1, 1)
    
    Exl_PonerValor AppExcel, fila, Col, "Información sobre :  " & Info
    Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, "Proveedor : " & txtDescrip.Text
    
    fila = fila + 1
    Exl_PonerValor AppExcel, fila, Col, "Marca     : " & txtDescripMr.Text
    
    fila = fila + 1
    Exl_PonerValor AppExcel, fila, Col, "Producto  : " & txtDescProd.Text & " ( " & txtCodProd.Text & " )"

    fila = fila + 2
    
    Exl_BajarGrillaExel spr, AppExcel, fila, Col, titCol
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

    frmPMPG.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub




Private Sub botExcelInt_Click()
Select Case TabAep.Tab
    Case 0
        L_TratarExcel sprInt(0), "INTERIOR ", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Importes"
    Case 1
        L_TratarExcel sprInt(1), "INTERIOR", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Unidades"
    Case 2
        L_TratarExcel sprInt(2), "INTERIOR", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Pasajeros"
    Case 3
        L_TratarExcel sprInt(3), "INTERIOR", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Promedios por Tickets"
    Case 4
        L_TratarExcel sprInt(4), "INTERIOR", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Promedios por Pasajeros"
End Select

End Sub

Private Sub botExcelTot_Click()

Select Case tabTot.Tab
    Case 0
        L_TratarExcel sprTotImp, "TOTAL COMPAÑIA", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Importes"
    Case 1
        L_TratarExcel sprTotUnid, "TOTAL COMPAÑIA", "Informe Acumulado por Prov-Marc-Prod (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ") ", "Unidades"
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

Private Sub botGrTortaINTE_Click()
Dim col_datos As Collection

    Set col_datos = New Collection
    
    Select Case TabAep.Tab
        Case 0
            L_LLenarColeccion col_datos, sprInt(0)
            FrmGraficos.CargarGrafico INTE, "Importes", col_datos
        Case 1
            L_LLenarColeccion col_datos, sprInt(1)
            FrmGraficos.CargarGrafico INTE, "Tickets", col_datos
        Case 2
            L_LLenarColeccion col_datos, sprInt(2)
            FrmGraficos.CargarGrafico INTE, "Pasajeros", col_datos

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
        botPorc(0).Enabled = False
        L_SeteoEjecutar False, EZEA
        frEzeA0.Refresh
        frEzeA1.Refresh
    Case 1
        frEzeA0.Visible = False
        frEzeA1.Visible = False
        L_Porcentages sprEzeAPorc(0), sprEzeA(0), "FILAS"
        L_Porcentages sprEzeAPorc(1), sprEzeA(1), "FILAS"

        botPorc(0).Enabled = True
        frEzeAPorc0.Refresh
        frEzeAPorc1.Refresh
        L_SeteoEjecutar True, EZEA
    Case 2
        frEzeA0.Visible = False
        frEzeA1.Visible = False
        L_Porcentages sprEzeAPorc(0), sprEzeA(0), "COL"
        L_Porcentages sprEzeAPorc(1), sprEzeA(1), "COL"

        botPorc(0).Enabled = True
        frEzeAPorc0.Refresh
        frEzeAPorc1.Refresh

        L_SeteoEjecutar True, EZEA
        
End Select

End Sub

Private Sub botPorcaep_Click(Index As Integer)

Select Case Index
    Case 0
        frAep0.Visible = True
        frAep1.Visible = True
        botPorcAep(0).Enabled = False
        L_SeteoEjecutar False, AERO
        frAep0.Refresh
        frAep1.Refresh
    Case 1
        frAep0.Visible = False
        frAep1.Visible = False
        L_Porcentages sprAepPorc(0), sprAEP(0), "FILAS"
        L_Porcentages sprAepPorc(1), sprAEP(1), "FILAS"
        'L_Porcentages sprAepPorc(2), sprAep(2), "FILAS"
        botPorcAep(0).Enabled = True
        frAepPorc0.Refresh
        frAepPorc1.Refresh
        L_SeteoEjecutar True, AERO
    Case 2
        frAep0.Visible = False
        frAep1.Visible = False
        L_Porcentages sprAepPorc(0), sprAEP(0), "COL"
        L_Porcentages sprAepPorc(1), sprAEP(1), "COL"
        'L_Porcentages sprAepPorc(2), sprAep(2), "COL"
        
        botPorcAep(0).Enabled = True
        frAepPorc0.Refresh
        frAepPorc1.Refresh
        L_SeteoEjecutar True, AERO
        
End Select


End Sub

Private Sub botPorcB_Click(Index As Integer)

Select Case Index
    Case 0
        frEzeB0.Visible = True
        frEzeB1.Visible = True
        botPorcB(0).Enabled = False
        L_SeteoEjecutar False, EZEB
        frEzeB0.Refresh
        frEzeB1.Refresh
    Case 1
        frEzeB0.Visible = False
        frEzeB1.Visible = False
        L_Porcentages sprEzeBPorc(0), sprEzeB(0), "FILAS"
        L_Porcentages sprEzeBPorc(1), sprEzeB(1), "FILAS"
        'L_Porcentages sprEzeBPorc(2), sprEzeB(2), "FILAS"
        botPorcB(0).Enabled = True
        frEzeBPorc0.Refresh
        frEzeBPorc1.Refresh
        L_SeteoEjecutar True, EZEB
    Case 2
        frEzeB0.Visible = False
        frEzeB1.Visible = False
        L_Porcentages sprEzeBPorc(0), sprEzeB(0), "COL"
        L_Porcentages sprEzeBPorc(1), sprEzeB(1), "COL"
        'L_Porcentages sprEzeBPorc(2), sprEzeB(2), "COL"
        
        botPorcB(0).Enabled = True
        frEzeBPorc0.Refresh
        frEzeBPorc1.Refresh
        L_SeteoEjecutar True, EZEB
        
End Select

End Sub







Private Sub botPorcInt_Click(Index As Integer)
Select Case Index
    Case 0
        FrInt0.Visible = True
        FrInt1.Visible = True
        botPorcint(0).Enabled = False
        L_SeteoEjecutar False, INTE
        FrInt0.Refresh
        FrInt1.Refresh
    Case 1
        FrInt0.Visible = False
        FrInt1.Visible = False
        L_Porcentages SprIntPorc(0), sprInt(0), "FILAS"
        L_Porcentages SprIntPorc(1), sprInt(1), "FILAS"
        L_Porcentages SprIntPorc(2), sprInt(2), "FILAS"
        botPorcint(0).Enabled = True
        FrIntPorc0.Refresh
        FrIntPorc1.Refresh
        L_SeteoEjecutar True, INTE
    Case 2
        FrInt0.Visible = False
        FrInt1.Visible = False
        L_Porcentages SprIntPorc(0), sprInt(0), "COL"
        L_Porcentages SprIntPorc(1), sprInt(1), "COL"
        L_Porcentages SprIntPorc(2), sprInt(2), "COL"
        
        botPorcint(0).Enabled = True
        FrIntPorc0.Refresh
        FrIntPorc1.Refresh
        L_SeteoEjecutar True, INTE
        
End Select

End Sub

Private Sub botPorcT_Click(Index As Integer)

Select Case Index
    Case 0
        Frame1(0).Visible = True
        Frame2(0).Visible = True
        botPorcT(0).Enabled = False
        L_SeteoEjecutar False, "TOT"
        Frame1(0).Refresh
        Frame2(0).Refresh
    Case 1
        Frame1(0).Visible = False
        Frame2(0).Visible = False
        L_Porcentages sprTotPorc(0), sprTotImp, "FILAS"
        L_Porcentages sprTotPorc(1), sprTotUnid, "FILAS"
        botPorcT(0).Enabled = True
        Frame1(1).Refresh
        Frame2(1).Refresh
        L_SeteoEjecutar True, "TOT"
    Case 2
        Frame1(0).Visible = False
        Frame2(0).Visible = False
        L_Porcentages sprTotPorc(0), sprTotImp, "COL"
        L_Porcentages sprTotPorc(1), sprTotUnid, "COL"
        botPorcT(0).Enabled = True
        Frame1(1).Refresh
        Frame2(1).Refresh
        L_SeteoEjecutar True, "TOT"
        
End Select

End Sub

Private Sub cboRubro_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    cboRubro.ListIndex = -1
End If
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
FuncLocal_SeteoTABS tabEspigon
End Sub

Private Sub Form_Load()
Dim i
Dim sql As String

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

UsoPorIntA = False
UsoPorIntB = False
UsoPorTotal = False
UsoPorAero = False

L_RefrescarFrames

HelpPROD = False
L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmPMPG"

sql = " SELECT cod_rubr,descrip FROM baires.rubro WHERE cod_rubr <> 'REG' "
sql = sql & " ORDER BY cod_rubr"

FuncCbos_LlenarCbo cboRubro, sql

End Sub




Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmPMPG"
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



Private Sub tabAep_Click(PreviousTab As Integer)
frAep0.Refresh
frAep1.Refresh

Select Case TabAep.Tab
    Case 0, 1, 2
        botPorcAep(0).Visible = True
        botPorcAep(1).Visible = True
        botPorcAep(2).Visible = True
        botGrTortaAEP.Visible = True
    Case 3, 4
        botPorcAep(0).Visible = False
        botPorcAep(1).Visible = False
        botPorcAep(2).Visible = False
        botGrTortaAEP.Visible = False
End Select

End Sub

Private Sub tabEzeA_Click(PreviousTab As Integer)
frEzeA0.Refresh
frEzeA1.Refresh

Select Case TabEzeA.Tab
    Case 0, 1, 2
        botPorc(0).Visible = True
        botPorc(1).Visible = True
        botPorc(2).Visible = True
        botGrTortaEA.Visible = True

    Case 3, 4
        botPorc(0).Visible = False
        botPorc(1).Visible = False
        botPorc(2).Visible = False
        botGrTortaEA.Visible = False

End Select


End Sub

Private Sub tabEzeB_Click(PreviousTab As Integer)
frEzeB0.Refresh
frEzeB1.Refresh


Select Case TabEzeB.Tab
    Case 0, 1, 2
        botPorcB(0).Visible = True
        botPorcB(1).Visible = True
        botPorcB(2).Visible = True
        botGrTortaEB.Visible = True
    Case 3, 4
        botPorcB(0).Visible = False
        botPorcB(1).Visible = False
        botPorcB(2).Visible = False
        botGrTortaEB.Visible = False
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

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim codigo$, desc$
Dim sqlProv$
    
Select Case Button.Key
    Case "a"
        If frmHelpPv.MuestraHlp(codigo$, desc$, "PROVEEDOR", "Select cod_prov, descrip from baires.proveedor ") = vbOK Then
            txtCodPv.Text = codigo$
            txtDescrip.Text = desc$
            SendKeys "{TAB}"
    
        End If

    Case "B"
            txtCodPv.Text = ""
            txtDescrip.Text = ""
            txtCodMr.Text = ""
            txtDescripMr.Text = ""
            txtCodProd.Text = ""
            txtDescProd.Text = ""
            txtCodProd.Locked = True
            Toolbar2.Buttons(1).Enabled = False

End Select


End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
Dim codigo$, desc$
Dim sqlProv$
Dim sql As String
    
Select Case Button.Key
    Case "a"
        sql = "Select cod_PROD, descrip from baires.producto WHERE cod_prov = '" & txtCodPv.Text & "' and cod_marca = '" & txtCodMr.Text & "' "
        If frmHelpPv.MuestraHlp(codigo$, desc$, "PRODUCTO", sql) = vbOK Then
            txtCodProd.Text = codigo$
            txtDescProd.Text = desc$
            
            HelpPROD = True
            SendKeys "{TAB}"
    
        End If

    Case "B"
            txtCodProd.Text = ""
            txtDescProd.Text = ""

End Select


End Sub


Private Sub Toolbar3_ButtonClick(ByVal Button As ComctlLib.Button)
Dim codigo$, desc$
Dim sqlProv$
    
Select Case Button.Key
    Case "a"
        If frmHelpPv.MuestraHlp(codigo$, desc$, "MARCA", "Select cod_marca, descrip from baires.marca WHERE cod_prov = '" & txtCodPv.Text & "' ") = vbOK Then
            txtCodMr.Text = codigo$
            txtDescripMr.Text = desc$
            txtCodProd.Locked = False
            Toolbar2.Buttons(1).Enabled = True
            SendKeys "{TAB}"
    
        End If
    Case "B"
            txtCodMr.Text = ""
            txtDescripMr.Text = ""
            txtCodProd.Text = ""
            txtDescProd.Text = ""
            txtCodProd.Locked = True
            Toolbar2.Buttons(1).Enabled = False
            

End Select


End Sub


Private Sub txtCodProd_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub txtCodProd_LostFocus()
Dim sql$
Dim desc$

desc$ = ""

If txtCodProd.Text <> "" Then
      If Not HelpPROD Then
        sql$ = "SELECT  descrip FROM baires.producto WHERE cod_prod = " & txtCodProd.Text
        
        Call Func_ObtenerDesc(sql$, desc$)
            
        txtDescProd = desc$

      End If
Else
    txtDescProd.Text = ""
End If
  
HelpPROD = False


End Sub


