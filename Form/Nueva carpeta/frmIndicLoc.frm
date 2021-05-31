VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIndicLoc 
   Caption         =   "INDICADORES por Local"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   9195
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   8700
      Top             =   1830
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   3810
      Left            =   135
      TabIndex        =   2
      Top             =   1560
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   6720
      _Version        =   327680
      Tabs            =   5
      Tab             =   1
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
      TabPicture(0)   =   "frmIndicLoc.frx":0000
      Tab(0).ControlCount=   0
      Tab(0).ControlEnabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmIndicLoc.frx":001C
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "tabLocA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "botExcelA"
      Tab(1).Control(1).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmIndicLoc.frx":0038
      Tab(2).ControlCount=   2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabLocB"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "botExcelB"
      Tab(2).Control(1).Enabled=   -1  'True
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmIndicLoc.frx":0054
      Tab(3).ControlCount=   2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tabLocE"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "botExcelE"
      Tab(3).Control(1).Enabled=   -1  'True
      TabCaption(4)   =   "INTERIOR"
      TabPicture(4)   =   "frmIndicLoc.frx":0070
      Tab(4).ControlCount=   2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TabLocI"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "botExcelI"
      Tab(4).Control(1).Enabled=   -1  'True
      Begin VB.CommandButton botExcelI 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67575
         Picture         =   "frmIndicLoc.frx":008C
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   720
         Width           =   780
      End
      Begin TabDlg.SSTab TabLocI 
         Height          =   3195
         Left            =   -74325
         TabIndex        =   47
         Top             =   405
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   5636
         _Version        =   327680
         Tabs            =   6
         TabsPerRow      =   6
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
         TabCaption(0)   =   "L04 CORD"
         TabPicture(0)   =   "frmIndicLoc.frx":061E
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprInt(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "L10 MEND"
         TabPicture(1)   =   "frmIndicLoc.frx":063A
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprInt(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "L11 IGUA"
         TabPicture(2)   =   "frmIndicLoc.frx":0656
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprInt(2)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "L15 MDPL"
         TabPicture(3)   =   "frmIndicLoc.frx":0672
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprInt(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "L16 MEND"
         TabPicture(4)   =   "frmIndicLoc.frx":068E
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprInt(4)"
         Tab(4).Control(0).Enabled=   0   'False
         TabCaption(5)   =   "L18 BARI"
         TabPicture(5)   =   "frmIndicLoc.frx":06AA
         Tab(5).ControlCount=   1
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "sprInt(5)"
         Tab(5).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprInt 
            Height          =   2655
            Index           =   5
            Left            =   -74865
            OleObjectBlob   =   "frmIndicLoc.frx":06C6
            TabIndex        =   54
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprInt 
            Height          =   2655
            Index           =   4
            Left            =   -74865
            OleObjectBlob   =   "frmIndicLoc.frx":0BE7
            TabIndex        =   53
            Top             =   435
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprInt 
            Height          =   2655
            Index           =   3
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":1108
            TabIndex        =   52
            Top             =   420
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprInt 
            Height          =   2655
            Index           =   2
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":1629
            TabIndex        =   51
            Top             =   450
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprInt 
            Height          =   2655
            Index           =   1
            Left            =   -74865
            OleObjectBlob   =   "frmIndicLoc.frx":1B4A
            TabIndex        =   50
            Top             =   420
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprInt 
            Height          =   2655
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "frmIndicLoc.frx":206B
            TabIndex        =   49
            Top             =   435
            Width           =   6375
         End
      End
      Begin VB.CommandButton botExcelE 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67545
         Picture         =   "frmIndicLoc.frx":258C
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   675
         Width           =   780
      End
      Begin VB.CommandButton botExcelB 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67560
         Picture         =   "frmIndicLoc.frx":2B1E
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   675
         Width           =   780
      End
      Begin VB.CommandButton botExcelA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   7425
         Picture         =   "frmIndicLoc.frx":30B0
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   630
         Width           =   780
      End
      Begin TabDlg.SSTab tabLocA 
         Height          =   3195
         Left            =   675
         TabIndex        =   23
         Top             =   420
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   5636
         _Version        =   327680
         Tabs            =   4
         TabsPerRow      =   5
         TabHeight       =   441
         BackColor       =   12632256
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
         TabCaption(0)   =   "L05"
         TabPicture(0)   =   "frmIndicLoc.frx":3642
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprEzeA(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "L06"
         TabPicture(1)   =   "frmIndicLoc.frx":365E
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprEzeA(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "L07"
         TabPicture(2)   =   "frmIndicLoc.frx":367A
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprEzeA(2)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "L08"
         TabPicture(3)   =   "frmIndicLoc.frx":3696
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprEzeA(3)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "frmIndicLoc.frx":36B2
            TabIndex        =   24
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   1
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":3BD3
            TabIndex        =   25
            Top             =   390
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   2
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":40F4
            TabIndex        =   26
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   3
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":4615
            TabIndex        =   27
            Top             =   405
            Width           =   6375
         End
      End
      Begin TabDlg.SSTab tabLocB 
         Height          =   3195
         Left            =   -74325
         TabIndex        =   28
         Top             =   435
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   5636
         _Version        =   327680
         TabsPerRow      =   5
         TabHeight       =   441
         BackColor       =   12632256
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
         TabCaption(0)   =   "L01"
         TabPicture(0)   =   "frmIndicLoc.frx":4B36
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprEzeB(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "L02"
         TabPicture(1)   =   "frmIndicLoc.frx":4B52
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprEzeB(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "L03"
         TabPicture(2)   =   "frmIndicLoc.frx":4B6E
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprEzeB(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2655
            Index           =   2
            Left            =   -74850
            OleObjectBlob   =   "frmIndicLoc.frx":4B8A
            TabIndex        =   34
            Top             =   420
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2655
            Index           =   1
            Left            =   -74850
            OleObjectBlob   =   "frmIndicLoc.frx":50AB
            TabIndex        =   33
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   4
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":55CC
            TabIndex        =   29
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   5
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":5AED
            TabIndex        =   30
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   6
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":600E
            TabIndex        =   31
            Top             =   390
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2655
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "frmIndicLoc.frx":652F
            TabIndex        =   32
            Top             =   405
            Width           =   6375
         End
      End
      Begin TabDlg.SSTab tabLocE 
         Height          =   3195
         Left            =   -74295
         TabIndex        =   35
         Top             =   435
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   5636
         _Version        =   327680
         Tabs            =   2
         TabsPerRow      =   5
         TabHeight       =   441
         BackColor       =   12632256
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
         TabCaption(0)   =   "L09"
         TabPicture(0)   =   "frmIndicLoc.frx":6A50
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprAep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "L14"
         TabPicture(1)   =   "frmIndicLoc.frx":6A6C
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprAep(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprAep 
            Height          =   2655
            Index           =   1
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":6A88
            TabIndex        =   42
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprAep 
            Height          =   2655
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "frmIndicLoc.frx":6FA9
            TabIndex        =   41
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   9
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":74CA
            TabIndex        =   40
            Top             =   390
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   8
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":79EB
            TabIndex        =   39
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   7
            Left            =   -74880
            OleObjectBlob   =   "frmIndicLoc.frx":7F0C
            TabIndex        =   38
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2655
            Index           =   4
            Left            =   -74850
            OleObjectBlob   =   "frmIndicLoc.frx":842D
            TabIndex        =   37
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2655
            Index           =   3
            Left            =   -74850
            OleObjectBlob   =   "frmIndicLoc.frx":894E
            TabIndex        =   36
            Top             =   420
            Width           =   6375
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
         Left            =   135
         Picture         =   "frmIndicLoc.frx":8E6F
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
         Picture         =   "frmIndicLoc.frx":8F71
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
         Picture         =   "frmIndicLoc.frx":9073
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
            Left            =   4980
            TabIndex        =   16
            Top             =   210
            Width           =   3015
            Begin VB.CommandButton botHelpFDAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmIndicLoc.frx":9895
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmIndicLoc.frx":9A07
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   600
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesdeAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   19
               Top             =   255
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               _Version        =   327680
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-##-####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFHastaAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   20
               Top             =   600
               Width           =   1125
               _ExtentX        =   1984
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
               TabIndex        =   22
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
               TabIndex        =   21
               Top             =   600
               Width           =   1185
            End
         End
         Begin VB.OptionButton optFechas 
            Caption         =   "Acumulado Anual "
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   15
            Top             =   375
            Width           =   1800
         End
         Begin VB.OptionButton optFechas 
            Caption         =   "Acumulado Rango Mes Corriente"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   90
            TabIndex        =   14
            Top             =   735
            Value           =   -1  'True
            Width           =   1845
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
            Left            =   1935
            TabIndex        =   4
            Top             =   225
            Width           =   3015
            Begin VB.CommandButton botHelpFH 
               Height          =   345
               Left            =   2535
               Picture         =   "frmIndicLoc.frx":9B79
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2550
               Picture         =   "frmIndicLoc.frx":9CEB
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
               _Version        =   327680
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
      Height          =   450
      Left            =   135
      TabIndex        =   46
      Top             =   5445
      Width           =   8730
   End
End
Attribute VB_Name = "frmIndicLoc"
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

Private Sub L_TratarExcel(SprTot As control, Esp As String, Titulo As String)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim fila As Integer, filaant As Integer
Dim i
Dim tit As Variant
Dim nombre As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

frmIndicLoc.caption = Aplicacion.SeteoProceso(frmIndicLoc.caption)

If nombre <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.application.Visible = True
    
    ReDim titCol(SprTot.MaxCols)
    col = 1
    fila = 3
    
    For i = 1 To SprTot.MaxCols
        SprTot.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 1, 1, Titulo
    rango = Exl_rangos(1, 1, 1, SprTot.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, col, "Indicadores por Local (" & mskFDesde.FormattedText & "-" & mskFHasta.FormattedText & ")"
    rango = Exl_rangos(fila, fila, 1, SprTot.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    Select Case Esp
    
    Case EZEA
    For i = 0 To 3
        fila = fila + 2
        Exl_PonerValor AppExcel, fila, col, tabLocA.TabCaption(i)
        rango = Exl_rangos(fila, fila, 1, SprTot.MaxCols)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        fila = fila + 1
        Exl_BajarGrillaExel sprEzeA(i), AppExcel, fila, 1, titCol
        rango = Exl_rangos(fila + 1, fila + SprTot.MaxRows, col + 1, SprTot.MaxCols)
        Exl_Format AppExcel, rango
        fila = fila + SprTot.MaxRows + 1
        rango = Exl_rangos(fila, fila, col, SprTot.MaxCols)
        Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Next
    Case EZEB
    For i = 0 To 2
        fila = fila + 2
        Exl_PonerValor AppExcel, fila, col, tabLocB.TabCaption(i)
        rango = Exl_rangos(fila, fila, 1, SprTot.MaxCols)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        fila = fila + 1
        Exl_BajarGrillaExel sprEzeB(i), AppExcel, fila, 1, titCol
        rango = Exl_rangos(fila + 1, fila + SprTot.MaxRows, col + 1, SprTot.MaxCols)
        Exl_Format AppExcel, rango
        fila = fila + SprTot.MaxRows + 1
        rango = Exl_rangos(fila, fila, col, SprTot.MaxCols)
        Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Next
    Case AERO
    For i = 0 To 1
        fila = fila + 2
        Exl_PonerValor AppExcel, fila, col, tabLocE.TabCaption(i)
        rango = Exl_rangos(fila, fila, 1, SprTot.MaxCols)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        fila = fila + 1
        Exl_BajarGrillaExel sprAep(i), AppExcel, fila, 1, titCol
        rango = Exl_rangos(fila + 1, fila + SprTot.MaxRows, col + 1, SprTot.MaxCols)
        Exl_Format AppExcel, rango
        fila = fila + SprTot.MaxRows + 1
        rango = Exl_rangos(fila, fila, col, SprTot.MaxCols)
        Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Next
    
    ''''''''''''
  Case INTE
    For i = 0 To 5
        fila = fila + 1
        Exl_PonerValor AppExcel, fila, col, TabLocI.TabCaption(i)
        rango = Exl_rangos(fila, fila, 1, SprTot.MaxCols)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        fila = fila + 1
        Exl_BajarGrillaExel sprInt(i), AppExcel, fila, 1, titCol
        rango = Exl_rangos(fila + 1, fila + SprTot.MaxRows, col + 1, SprTot.MaxCols)
        Exl_Format AppExcel, rango
        fila = fila + SprTot.MaxRows + 1
        rango = Exl_rangos(fila, fila, col, SprTot.MaxCols)
        Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Next
    
    End Select
    
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If

    AppExcel.SaveAs nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmIndicLoc.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub



                       'con porcentajes
Private Function L_Armarcondicion(anio As Integer) As String
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant

If optFechas(0).Value Then
    If anio = Year(Date) Then
        fechaDesde = "01-01-" & Trim(str(anio))
        fechaHasta = Format$(Date - 1, "dd-mm") & "-" & Trim(str(anio))
    Else
        fechaDesde = "01-01-" & Trim(str(anio))
        fechaHasta = Format$(Date - 1, "dd-mm") & "-" & Trim(str(anio))
    End If
Else
    If anio = Year(mskFDesde.FormattedText) Then
        fechaDesde = mskFDesde.FormattedText
        fechaHasta = mskFHasta.FormattedText
    Else
        fechaDesde = mskFDesdeAnt.FormattedText
        fechaHasta = mskFHastaAnt.FormattedText
    End If
End If

Cond = " WHERE fch_vta between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

L_Armarcondicion = Cond

End Function

Private Sub L_LimpiarGrillas()
Dim i As Integer
labInfo.caption = ""

sprAep(0).MaxRows = 0
sprAep(1).MaxRows = 0

sprEzeA(0).MaxRows = 0
sprEzeA(1).MaxRows = 0
sprEzeA(2).MaxRows = 0
sprEzeA(3).MaxRows = 0

sprEzeB(0).MaxRows = 0
sprEzeB(1).MaxRows = 0
sprEzeB(2).MaxRows = 0

For i = 0 To 5
sprInt(i).MaxRows = 0
Next i

End Sub

Private Sub L_PonerenGrilla(spr As control, item As String, valor As Single)
    spr.MaxRows = spr.MaxRows + 1
    spr.SetText 1, spr.MaxRows, Trim(item)
    spr.SetText 2, spr.MaxRows, str(valor)
End Sub

Private Sub L_Refrescar()
Dim SQL As String
Dim sqlX As String
Dim rs As Recordset

On Error GoTo ErrInd:

frmIndicLoc.caption = Aplicacion.SeteoProceso(frmIndicLoc.caption)


SQL = " SELECT "
SQL = SQL & " cod_depn, "
SQL = SQL & " cod_sdep, "
SQL = SQL & " cod_local, "
SQL = SQL & " sum(cant_tickets) cant_t, "
SQL = SQL & " sum(importe) imp "
'sql = sql & " sum(cant_pax) cant_p "
SQL = SQL & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)))
SQL = SQL & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)))
SQL = SQL & "group by cod_depn,cod_sdep,cod_local"
SQL = SQL & " order by cod_depn,cod_sdep,cod_local"

sqlX = " SELECT "
sqlX = sqlX & " cod_depn, "
sqlX = sqlX & " cod_sdep, "
sqlX = sqlX & " cod_local, "
sqlX = sqlX & " sum(cant_tickets) cant_t, "
sqlX = sqlX & " sum(importe) imp "
'sqlX = sqlX & " sum(cant_pax) cant_p "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)) - 1)
sqlX = sqlX & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)) - 1)
sqlX = sqlX & "group by cod_depn,cod_sdep,cod_local "
sqlX = sqlX & " order by cod_depn,cod_sdep,cod_local"


If Aplicacion.ObtenerRsDAO(SQL, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        L_DecoEspigon
        If Aplicacion.ObtenerRsDAO(sqlX, RsDataAnt) Then
            L_DecoEspigonAnt
        End If
        L_TratarVolados
        L_TratarVoladosHIST
        'L_LlenarTotales
    
        L_Resaltar
        
        Aplicacion.CerrarDAO RsDataAnt
    End If
    
    Aplicacion.CerrarDAO RsData

End If

ErrInd:
    frmIndicLoc.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub

Private Sub L_Resaltar()
Dim fila As Integer
Dim i

For fila = 1 To 6
    For i = 0 To 3
        Spread.spread_ResaltarCelda sprEzeA(i), 4, fila
        Spread.spread_ResaltarCelda sprEzeA(i), 5, fila
    Next
    For i = 0 To 2
        Spread.spread_ResaltarCelda sprEzeB(i), 4, fila
        Spread.spread_ResaltarCelda sprEzeB(i), 5, fila
    Next
    For i = 0 To 1
        Spread.spread_ResaltarCelda sprAep(i), 4, fila
        Spread.spread_ResaltarCelda sprAep(i), 5, fila
    Next
Next
End Sub


Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant

Do While Not RsData.EOF
    Select Case RsData!cod_depn
        Case DSLoc(1).Dep
            Select Case RsData!cod_local
                Case Left(DSLoc(1).locales(1), 3)
                    L_PonerenGrilla sprAep(0), "Facturación", str(RsData!imp)
                    L_PonerenGrilla sprAep(0), "Tickets", str(RsData!cant_t)
                    L_PonerenGrilla sprAep(0), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                Case Left(DSLoc(1).locales(2), 3)
                    L_PonerenGrilla sprAep(1), "Facturación", str(RsData!imp)
                    L_PonerenGrilla sprAep(1), "Tickets", str(RsData!cant_t)
                    L_PonerenGrilla sprAep(1), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
            End Select
        Case DSLoc(2).Dep
            Select Case RsData!cod_sdep
                Case DSLoc(2).Sdep
                    Select Case RsData!cod_local
                        Case Left(DSLoc(2).locales(1), 3)
                            L_PonerenGrilla sprEzeA(0), "Facturación", str(RsData!imp)
                            L_PonerenGrilla sprEzeA(0), "Tickets", str(RsData!cant_t)
                            L_PonerenGrilla sprEzeA(0), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                        Case Left(DSLoc(2).locales(3), 3)  '2
                            L_PonerenGrilla sprEzeA(1), "Facturación", str(RsData!imp)
                            L_PonerenGrilla sprEzeA(1), "Tickets", str(RsData!cant_t)
                            L_PonerenGrilla sprEzeA(1), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                        Case Left(DSLoc(2).locales(4), 3)  '3
                            L_PonerenGrilla sprEzeA(2), "Facturación", str(RsData!imp)
                            L_PonerenGrilla sprEzeA(2), "Tickets", str(RsData!cant_t)
                            L_PonerenGrilla sprEzeA(2), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                        Case Left(DSLoc(2).locales(6), 3)  '5
                            L_PonerenGrilla sprEzeA(3), "Facturación", str(RsData!imp)
                            L_PonerenGrilla sprEzeA(3), "Tickets", str(RsData!cant_t)
                            L_PonerenGrilla sprEzeA(3), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                    End Select
                
                Case DSLoc(3).Sdep
                    Select Case RsData!cod_local
                        Case Left(DSLoc(3).locales(1), 3)
                            L_PonerenGrilla sprEzeB(0), "Facturación", str(RsData!imp)
                            L_PonerenGrilla sprEzeB(0), "Tickets", str(RsData!cant_t)
                            L_PonerenGrilla sprEzeB(0), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                        Case Left(DSLoc(3).locales(3), 3)
                            L_PonerenGrilla sprEzeB(1), "Facturación", str(RsData!imp)
                            L_PonerenGrilla sprEzeB(1), "Tickets", str(RsData!cant_t)
                            L_PonerenGrilla sprEzeB(1), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                        Case Left(DSLoc(3).locales(4), 3)
                            L_PonerenGrilla sprEzeB(2), "Facturación", str(RsData!imp)
                            L_PonerenGrilla sprEzeB(2), "Tickets", str(RsData!cant_t)
                            L_PonerenGrilla sprEzeB(2), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                    End Select
            End Select
        
             Case DSLoc(4).Dep
            Select Case RsData!cod_local
                Case Left(DSLoc(4).locales(1), 3)
                    L_PonerenGrilla sprInt(0), "Facturación", str(RsData!imp)
                    L_PonerenGrilla sprInt(0), "Tickets", str(RsData!cant_t)
                    L_PonerenGrilla sprInt(0), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                Case Left(DSLoc(5).locales(1), 3)
                    L_PonerenGrilla sprInt(2), "Facturación", str(RsData!imp)
                    L_PonerenGrilla sprInt(2), "Tickets", str(RsData!cant_t)
                    L_PonerenGrilla sprInt(2), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                Case Left(DSLoc(6).locales(1), 3)
                    L_PonerenGrilla sprInt(1), "Facturación", str(RsData!imp)
                    L_PonerenGrilla sprInt(1), "Tickets", str(RsData!cant_t)
                    L_PonerenGrilla sprInt(1), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                Case Left(DSLoc(6).locales(2), 3)
                    L_PonerenGrilla sprInt(4), "Facturación", str(RsData!imp)
                    L_PonerenGrilla sprInt(4), "Tickets", str(RsData!cant_t)
                    L_PonerenGrilla sprInt(4), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                Case Left(DSLoc(7).locales(1), 3)
                    L_PonerenGrilla sprInt(3), "Facturación", str(RsData!imp)
                    L_PonerenGrilla sprInt(3), "Tickets", str(RsData!cant_t)
                    L_PonerenGrilla sprInt(3), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
                Case Left(DSLoc(8).locales(1), 3)
                    L_PonerenGrilla sprInt(5), "Facturación", str(RsData!imp)
                    L_PonerenGrilla sprInt(5), "Tickets", str(RsData!cant_t)
                    L_PonerenGrilla sprInt(5), "Pr. x Tickets", str(RsData!imp / RsData!cant_t)
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
            Select Case RsDataAnt!cod_local
                Case Left(DSLoc(1).locales(1), 3)
                     sprAep(0).SetText 3, 1, str(RsDataAnt!imp)
                     sprAep(0).SetText 3, 2, str(RsDataAnt!cant_t)
                     sprAep(0).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                Case Left(DSLoc(1).locales(2), 3)
                     sprAep(1).SetText 3, 1, str(RsDataAnt!imp)
                     sprAep(1).SetText 3, 2, str(RsDataAnt!cant_t)
                     sprAep(1).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
            End Select
        Case DSLoc(2).Dep
            Select Case RsDataAnt!cod_sdep
                Case DSLoc(2).Sdep
                    Select Case RsDataAnt!cod_local
                        Case Left(DSLoc(2).locales(1), 3)
                            sprEzeA(0).SetText 3, 1, str(RsDataAnt!imp)
                            sprEzeA(0).SetText 3, 2, str(RsDataAnt!cant_t)
                            sprEzeA(0).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                        Case Left(DSLoc(2).locales(3), 3)
                             sprEzeA(1).SetText 3, 1, str(RsDataAnt!imp)
                             sprEzeA(1).SetText 3, 2, str(RsDataAnt!cant_t)
                             sprEzeA(1).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                        Case Left(DSLoc(2).locales(4), 3)
                             sprEzeA(2).SetText 3, 1, str(RsDataAnt!imp)
                             sprEzeA(2).SetText 3, 2, str(RsDataAnt!cant_t)
                             sprEzeA(2).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                        Case Left(DSLoc(2).locales(6), 3)
                             sprEzeA(3).SetText 3, 1, str(RsDataAnt!imp)
                             sprEzeA(3).SetText 3, 2, str(RsDataAnt!cant_t)
                             sprEzeA(3).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                    End Select
                
                Case DSLoc(3).Sdep
                    Select Case RsDataAnt!cod_local
                        Case Left(DSLoc(3).locales(1), 3)
                             sprEzeB(0).SetText 3, 1, str(RsDataAnt!imp)
                             sprEzeB(0).SetText 3, 2, str(RsDataAnt!cant_t)
                             sprEzeB(0).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                        Case Left(DSLoc(3).locales(3), 3)
                             sprEzeB(1).SetText 3, 1, str(RsDataAnt!imp)
                             sprEzeB(1).SetText 3, 2, str(RsDataAnt!cant_t)
                             sprEzeB(1).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                        Case Left(DSLoc(3).locales(4), 3)
                             sprEzeB(2).SetText 3, 1, str(RsDataAnt!imp)
                             sprEzeB(2).SetText 3, 2, str(RsDataAnt!cant_t)
                             sprEzeB(2).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                    End Select
              End Select
      Case DSLoc(4).Dep
            Select Case RsDataAnt!cod_local
                Case Left(DSLoc(4).locales(1), 3)
                    sprInt(0).SetText 3, 1, str(RsDataAnt!imp)
                    sprInt(0).SetText 3, 2, str(RsDataAnt!cant_t)
                    sprInt(0).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                Case Left(DSLoc(5).locales(1), 3)
                    sprInt(2).SetText 3, 1, str(RsDataAnt!imp)
                    sprInt(2).SetText 3, 2, str(RsDataAnt!cant_t)
                    sprInt(2).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                Case Left(DSLoc(6).locales(1), 3)
                    sprInt(1).SetText 3, 1, str(RsDataAnt!imp)
                    sprInt(1).SetText 3, 2, str(RsDataAnt!cant_t)
                    sprInt(1).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                Case Left(DSLoc(6).locales(2), 3)
                    sprInt(4).SetText 3, 1, str(RsDataAnt!imp)
                    sprInt(4).SetText 3, 2, str(RsDataAnt!cant_t)
                    sprInt(4).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                Case Left(DSLoc(7).locales(1), 3)
                    sprInt(3).SetText 3, 1, str(RsDataAnt!imp)
                    sprInt(3).SetText 3, 2, str(RsDataAnt!cant_t)
                    sprInt(3).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                Case Left(DSLoc(8).locales(1), 3)
                    sprInt(5).SetText 3, 1, str(RsDataAnt!imp)
                    sprInt(5).SetText 3, 2, str(RsDataAnt!cant_t)
                    sprInt(5).SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
            End Select
   End Select

RsDataAnt.MoveNext
Loop


End Sub


Private Sub L_TratarVolados()
Dim SQL As String
Dim FD As String, FH As String
Dim rs_pax As Recordset

If optFechas(0).Value Then
    FD = "01-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText)
    FH = mskFDesde.FormattedText
Else
    FD = mskFDesde.FormattedText
    FH = mskFHasta.FormattedText
End If

SQL = "select min(fch_actualizado) Fch from estadis.control_carga"

If Aplicacion.ObtenerRsDAO(SQL, rs_pax) Then
    
    If CDate(FH) <= CDate(rs_pax!fch) Then
  
        SQL = " SELECT "
        SQL = SQL & "cod_depn,"
        SQL = SQL & "cod_sdep,"
        SQL = SQL & "local,"
        SQL = SQL & " SUM(pax_volados.cantidad) CANT_V "
        SQL = SQL & " From estadis.pax_volados "
        SQL = SQL & " Where "
        SQL = SQL & " fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
        SQL = SQL & " GROUP BY COD_DEPN,COD_SDEP,local"
        
        If Aplicacion.ObtenerRsDAO(SQL, RsDataVol) Then
            L_DecoVolados
            Aplicacion.CerrarDAO RsDataVol
        End If
    Else
        labInfo.caption = "Datos de pasajeros transitados Actualizados al día " & rs_pax!fch
    End If
        
End If


End Sub

Private Sub L_DecoVolados()
Dim valor As Variant

Do While Not RsDataVol.EOF
    Select Case RsDataVol!cod_depn
        Case DSLoc(1).Dep
            Select Case RsDataVol!local
                Case Left(DSLoc(1).locales(1), 3)
                    sprAep(0).MaxRows = sprAep(0).MaxRows + 1
                    sprAep(0).SetText 1, sprAep(0).MaxRows, "Pax Viajados"
                    sprAep(0).SetText 2, 4, str(RsDataVol!cant_v)
                   
                    sprAep(0).MaxRows = sprAep(0).MaxRows + 1
                    sprAep(0).SetText 1, sprAep(0).MaxRows, "Pr x Pax Viaj." '
                
                    sprAep(0).MaxRows = sprAep(0).MaxRows + 1
                    sprAep(0).SetText 1, sprAep(0).MaxRows, "Captación"
                    If RsDataVol!cant_v > 0 Then
                       sprAep(0).GetText 2, 1, valor
                       sprAep(0).SetText 2, 5, str(valor / RsDataVol!cant_v)
                       sprAep(0).GetText 2, 2, valor
                       sprAep(0).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
                    
                Case Left(DSLoc(1).locales(2), 3)
                    sprAep(1).MaxRows = sprAep(1).MaxRows + 1
                    sprAep(1).SetText 1, sprAep(1).MaxRows, "Pax Viajados"
                    sprAep(1).SetText 2, 4, str(RsDataVol!cant_v)
                    
                    sprAep(1).MaxRows = sprAep(1).MaxRows + 1
                    sprAep(1).SetText 1, sprAep(1).MaxRows, "Pr x Pax Viaj." '
                
                    sprAep(1).MaxRows = sprAep(1).MaxRows + 1
                    sprAep(1).SetText 1, sprAep(1).MaxRows, "Captación"
                    If RsDataVol!cant_v > 0 Then
                       sprAep(1).GetText 2, 1, valor
                       sprAep(1).SetText 2, 5, str(valor / RsDataVol!cant_v)
                       sprAep(1).GetText 2, 2, valor
                       sprAep(1).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
            
            End Select

        Case DSLoc(2).Dep
            Select Case RsDataVol!cod_sdep
                Case DSLoc(2).Sdep
                    Select Case RsDataVol!local
                        Case Left(DSLoc(2).locales(1), 3)
                            sprEzeA(0).MaxRows = sprEzeA(0).MaxRows + 1
                            sprEzeA(0).SetText 1, sprEzeA(0).MaxRows, "Pax Viajados"
                            sprEzeA(0).SetText 2, 4, str(RsDataVol!cant_v)
                            
                            sprEzeA(0).MaxRows = sprEzeA(0).MaxRows + 1
                            sprEzeA(0).SetText 1, sprEzeA(0).MaxRows, "Pr x Pax Viaj." '
                        
                            sprEzeA(0).MaxRows = sprEzeA(0).MaxRows + 1
                            sprEzeA(0).SetText 1, sprEzeA(0).MaxRows, "Captación"
                            If RsDataVol!cant_v > 0 Then
                               sprEzeA(0).GetText 2, 1, valor
                               sprEzeA(0).SetText 2, 5, str(valor / RsDataVol!cant_v)
                               sprEzeA(0).GetText 2, 2, valor
                               sprEzeA(0).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        
                        Case Left(DSLoc(2).locales(3), 3)
                            sprEzeA(1).MaxRows = sprEzeA(1).MaxRows + 1
                            sprEzeA(1).SetText 1, sprEzeA(1).MaxRows, "Pax Viajados"
                            sprEzeA(1).SetText 2, 4, str(RsDataVol!cant_v)
                            
                            sprEzeA(1).MaxRows = sprEzeA(1).MaxRows + 1
                            sprEzeA(1).SetText 1, sprEzeA(1).MaxRows, "Pr x Pax Viaj." '
                        
                            sprEzeA(1).MaxRows = sprEzeA(1).MaxRows + 1
                            sprEzeA(1).SetText 1, sprEzeA(1).MaxRows, "Captación"
                            If RsDataVol!cant_v > 0 Then
                               sprEzeA(1).GetText 2, 1, valor
                               sprEzeA(1).SetText 2, 5, str(valor / RsDataVol!cant_v)
                               sprEzeA(1).GetText 2, 2, valor
                               sprEzeA(1).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        
                        Case Left(DSLoc(2).locales(4), 3)
                            sprEzeA(2).MaxRows = sprEzeA(2).MaxRows + 1
                            sprEzeA(2).SetText 1, sprEzeA(2).MaxRows, "Pax Viajados"
                            sprEzeA(2).SetText 2, 4, str(RsDataVol!cant_v)
                            
                            sprEzeA(2).MaxRows = sprEzeA(2).MaxRows + 1
                            sprEzeA(2).SetText 1, sprEzeA(2).MaxRows, "Pr x Pax Viaj." '
                        
                            sprEzeA(2).MaxRows = sprEzeA(2).MaxRows + 1
                            sprEzeA(2).SetText 1, sprEzeA(2).MaxRows, "Captación"
                            If RsDataVol!cant_v > 0 Then
                               sprEzeA(2).GetText 2, 1, valor
                               sprEzeA(2).SetText 2, 5, str(valor / RsDataVol!cant_v)
                               sprEzeA(2).GetText 2, 2, valor
                               sprEzeA(2).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        
                        Case Left(DSLoc(2).locales(6), 3)
                            sprEzeA(3).MaxRows = sprEzeA(3).MaxRows + 1
                            sprEzeA(3).SetText 1, sprEzeA(3).MaxRows, "Pax Viajados"
                            sprEzeA(3).SetText 2, 4, str(RsDataVol!cant_v)
                            
                            sprEzeA(3).MaxRows = sprEzeA(3).MaxRows + 1
                            sprEzeA(3).SetText 1, sprEzeA(3).MaxRows, "Pr x Pax Viaj." '
                        
                            sprEzeA(3).MaxRows = sprEzeA(3).MaxRows + 1
                            sprEzeA(3).SetText 1, sprEzeA(3).MaxRows, "Captación"
                            If RsDataVol!cant_v > 0 Then
                               sprEzeA(3).GetText 2, 1, valor
                               sprEzeA(3).SetText 2, 5, str(valor / RsDataVol!cant_v)
                               sprEzeA(3).GetText 2, 2, valor
                               sprEzeA(3).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        
                    End Select
                        
                Case DSLoc(3).Sdep
                        Select Case RsDataVol!local
                            Case Left(DSLoc(3).locales(1), 3)
                                sprEzeB(0).MaxRows = sprEzeB(0).MaxRows + 1
                                sprEzeB(0).SetText 1, sprEzeB(0).MaxRows, "Pax Viajados"
                                sprEzeB(0).SetText 2, 4, str(RsDataVol!cant_v)
                                
                                sprEzeB(0).MaxRows = sprEzeB(0).MaxRows + 1
                                sprEzeB(0).SetText 1, sprEzeB(0).MaxRows, "Pr x Pax Viaj." '
                            
                                sprEzeB(0).MaxRows = sprEzeB(0).MaxRows + 1
                                sprEzeB(0).SetText 1, sprEzeB(0).MaxRows, "Captación"
                                If RsDataVol!cant_v > 0 Then
                                   sprEzeB(0).GetText 2, 1, valor
                                   sprEzeB(0).SetText 2, 5, str(valor / RsDataVol!cant_v)
                                   sprEzeB(0).GetText 2, 2, valor
                                   sprEzeB(0).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                                End If
                            Case Left(DSLoc(3).locales(3), 3)
                                sprEzeB(1).MaxRows = sprEzeB(1).MaxRows + 1
                                sprEzeB(1).SetText 1, sprEzeB(1).MaxRows, "Pax Viajados"
                                sprEzeB(1).SetText 2, 4, str(RsDataVol!cant_v)
                                
                                sprEzeB(1).MaxRows = sprEzeB(1).MaxRows + 1
                                sprEzeB(1).SetText 1, sprEzeB(1).MaxRows, "Pr x Pax Viaj." '
                            
                                sprEzeB(1).MaxRows = sprEzeB(1).MaxRows + 1
                                sprEzeB(1).SetText 1, sprEzeB(1).MaxRows, "Captación"
                                If RsDataVol!cant_v > 0 Then
                                   sprEzeB(1).GetText 2, 1, valor
                                   sprEzeB(1).SetText 2, 5, str(valor / RsDataVol!cant_v)
                                   sprEzeB(1).GetText 2, 2, valor
                                   sprEzeB(1).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                                End If
                            Case Left(DSLoc(3).locales(4), 3)
                                sprEzeB(2).MaxRows = sprEzeB(2).MaxRows + 1
                                sprEzeB(2).SetText 1, sprEzeB(2).MaxRows, "Pax Viajados"
                                sprEzeB(2).SetText 2, 4, str(RsDataVol!cant_v)
                                
                                sprEzeB(2).MaxRows = sprEzeB(2).MaxRows + 1
                                sprEzeB(2).SetText 1, sprEzeB(2).MaxRows, "Pr x Pax Viaj." '
                            
                                sprEzeB(2).MaxRows = sprEzeB(2).MaxRows + 1
                                sprEzeB(2).SetText 1, sprEzeB(2).MaxRows, "Captación"
                                If RsDataVol!cant_v > 0 Then
                                   sprEzeB(2).GetText 2, 1, valor
                                   sprEzeB(2).SetText 2, 5, str(valor / RsDataVol!cant_v)
                                   sprEzeB(2).GetText 2, 2, valor
                                   sprEzeB(2).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                                End If
                        
                        End Select
                 End Select
        ''''''''''''''''''''''''''''
       Case DSLoc(4).Dep
            Select Case RsDataVol!local
                Case Left(DSLoc(4).locales(1), 3)
                    sprInt(0).MaxRows = sprInt(0).MaxRows + 1
                    sprInt(0).SetText 1, sprInt(0).MaxRows, "Pax Viajados"
                    sprInt(0).SetText 2, 4, str(RsDataVol!cant_v)
                    
                    sprInt(0).MaxRows = sprInt(0).MaxRows + 1
                    sprInt(0).SetText 1, sprInt(0).MaxRows, "Pr x Pax Viaj." '
                
                    sprInt(0).MaxRows = sprInt(0).MaxRows + 1
                    sprInt(0).SetText 1, sprInt(0).MaxRows, "Captación"
                    If RsDataVol!cant_v > 0 Then
                       sprInt(0).GetText 2, 1, valor
                       sprInt(0).SetText 2, 5, str(valor / RsDataVol!cant_v)
                       sprInt(0).GetText 2, 2, valor
                       sprInt(0).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
              
                Case Left(DSLoc(5).locales(1), 3)
                    sprInt(2).MaxRows = sprInt(2).MaxRows + 1
                    sprInt(2).SetText 1, sprInt(2).MaxRows, "Pax Viajados"
                    sprInt(2).SetText 2, 4, str(RsDataVol!cant_v)
                    
                    sprInt(2).MaxRows = sprInt(2).MaxRows + 1
                    sprInt(2).SetText 1, sprInt(2).MaxRows, "Pr x Pax Viaj." '
                
                    sprInt(2).MaxRows = sprInt(2).MaxRows + 1
                    sprInt(2).SetText 1, sprInt(2).MaxRows, "Captación"
                    If RsDataVol!cant_v > 0 Then
                       sprInt(2).GetText 2, 1, valor
                       sprInt(2).SetText 2, 5, str(valor / RsDataVol!cant_v)
                       sprInt(2).GetText 2, 2, valor
                       sprInt(2).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
                
                Case Left(DSLoc(6).locales(1), 3)
                    sprInt(1).MaxRows = sprInt(1).MaxRows + 1
                    sprInt(1).SetText 1, sprInt(1).MaxRows, "Pax Viajados"
                    sprInt(1).SetText 2, 4, str(RsDataVol!cant_v)
                    
                    sprInt(1).MaxRows = sprInt(1).MaxRows + 1
                    sprInt(1).SetText 1, sprInt(1).MaxRows, "Pr x Pax Viaj." '
                
                    sprInt(1).MaxRows = sprInt(1).MaxRows + 1
                    sprInt(1).SetText 1, sprInt(1).MaxRows, "Captación"
                    If RsDataVol!cant_v > 0 Then
                       sprInt(1).GetText 2, 1, valor
                       sprInt(1).SetText 2, 5, str(valor / RsDataVol!cant_v)
                       sprInt(1).GetText 2, 2, valor
                       sprInt(1).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                     End If
                Case Left(DSLoc(6).locales(2), 3)
                    sprInt(4).MaxRows = sprInt(4).MaxRows + 1
                    sprInt(4).SetText 1, sprInt(4).MaxRows, "Pax Viajados"
                    sprInt(4).SetText 2, 4, str(RsDataVol!cant_v)
                    
                    sprInt(4).MaxRows = sprInt(4).MaxRows + 1
                    sprInt(4).SetText 1, sprInt(4).MaxRows, "Pr x Pax Viaj." '
                
                    sprInt(4).MaxRows = sprInt(4).MaxRows + 1
                    sprInt(4).SetText 1, sprInt(4).MaxRows, "Captación"
                    If RsDataVol!cant_v > 0 Then
                       sprInt(4).GetText 2, 1, valor
                       sprInt(4).SetText 2, 5, str(valor / RsDataVol!cant_v)
                       sprInt(4).GetText 2, 2, valor
                       sprInt(4).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
                Case Left(DSLoc(7).locales(1), 3)
                    sprInt(3).MaxRows = sprInt(3).MaxRows + 1
                    sprInt(3).SetText 1, sprInt(3).MaxRows, "Pax Viajados"
                    sprInt(3).SetText 2, 4, str(RsDataVol!cant_v)
                    
                    sprInt(3).MaxRows = sprInt(3).MaxRows + 1
                    sprInt(3).SetText 1, sprInt(3).MaxRows, "Pr x Pax Viaj." '
                
                    sprInt(3).MaxRows = sprInt(3).MaxRows + 1
                    sprInt(3).SetText 1, sprInt(3).MaxRows, "Captación"
                    If RsDataVol!cant_v > 0 Then
                       sprInt(3).GetText 2, 1, valor
                       sprInt(3).SetText 2, 5, str(valor / RsDataVol!cant_v)
                       sprInt(3).GetText 2, 2, valor
                       sprInt(3).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
                Case Left(DSLoc(8).locales(1), 3)
                    sprInt(5).MaxRows = sprInt(5).MaxRows + 1
                    sprInt(5).SetText 1, sprInt(5).MaxRows, "Pax Viajados"
                    sprInt(5).SetText 2, 4, str(RsDataVol!cant_v)
                    
                    sprInt(5).MaxRows = sprInt(5).MaxRows + 1
                    sprInt(5).SetText 1, sprInt(5).MaxRows, "Pr x Pax Viaj." '
                
                    sprInt(5).MaxRows = sprInt(5).MaxRows + 1
                    sprInt(5).SetText 1, sprInt(5).MaxRows, "Captación"
                    If RsDataVol!cant_v > 0 Then
                       sprInt(5).GetText 2, 1, valor
                       sprInt(5).SetText 2, 5, str(valor / RsDataVol!cant_v)
                       sprInt(5).GetText 2, 2, valor
                       sprInt(5).SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
            End Select
        ''''''''''''''''''''''''''''
    '''End Select
  End Select
        RsDataVol.MoveNext
Loop

End Sub



Private Sub L_DecoVolados_Ant()
Dim valor As Variant

Do While Not RsDataVol.EOF
    Select Case RsDataVol!cod_depn
        Case DSLoc(1).Dep
            Select Case RsDataVol!local
                Case Left(DSLoc(1).locales(1), 3)
                    sprAep(0).SetText 3, 4, str(RsDataVol!cant_v)
                    If RsDataVol!cant_v > 0 Then
                       sprAep(0).GetText 3, 1, valor
                       sprAep(0).SetText 3, 5, str(valor / RsDataVol!cant_v)
                       sprAep(0).GetText 3, 2, valor
                       sprAep(0).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
                    
                Case Left(DSLoc(1).locales(2), 3)
                    sprAep(1).SetText 3, 4, str(RsDataVol!cant_v)
                    If RsDataVol!cant_v > 0 Then
                       sprAep(1).GetText 3, 1, valor
                       sprAep(1).SetText 3, 5, str(valor / RsDataVol!cant_v)
                       sprAep(1).GetText 3, 2, valor
                       sprAep(1).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
            
            End Select

        Case DSLoc(2).Dep
            Select Case RsDataVol!cod_sdep
                Case DSLoc(2).Sdep
                    Select Case RsDataVol!local
                        Case Left(DSLoc(2).locales(1), 3)
                            sprEzeA(0).SetText 3, 4, str(RsDataVol!cant_v)
                            If RsDataVol!cant_v > 0 Then
                               sprEzeA(0).GetText 3, 1, valor
                               sprEzeA(0).SetText 3, 5, str(valor / RsDataVol!cant_v)
                               sprEzeA(0).GetText 3, 2, valor
                               sprEzeA(0).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        
                        Case Left(DSLoc(2).locales(3), 3)
                             sprEzeA(1).SetText 3, 4, str(RsDataVol!cant_v)
                            If RsDataVol!cant_v > 0 Then
                               sprEzeA(1).GetText 3, 1, valor
                               sprEzeA(1).SetText 3, 5, str(valor / RsDataVol!cant_v)
                               sprEzeA(1).GetText 3, 2, valor
                               sprEzeA(1).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        
                        Case Left(DSLoc(2).locales(4), 3)
                            sprEzeA(2).SetText 3, 4, str(RsDataVol!cant_v)
                            If RsDataVol!cant_v > 0 Then
                               sprEzeA(2).GetText 3, 1, valor
                               sprEzeA(2).SetText 3, 5, str(valor / RsDataVol!cant_v)
                               sprEzeA(2).GetText 3, 2, valor
                               sprEzeA(2).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        
                        Case Left(DSLoc(2).locales(6), 3)
                            sprEzeA(3).SetText 3, 4, str(RsDataVol!cant_v)
                            If RsDataVol!cant_v > 0 Then
                               sprEzeA(3).GetText 3, 1, valor
                               sprEzeA(3).SetText 3, 5, str(valor / RsDataVol!cant_v)
                               sprEzeA(3).GetText 3, 2, valor
                               sprEzeA(3).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                            End If
                        
                    End Select
                        
                Case DSLoc(3).Sdep
                        Select Case RsDataVol!local
                            Case Left(DSLoc(3).locales(1), 3)
                                sprEzeB(0).SetText 3, 4, str(RsDataVol!cant_v)
                                If RsDataVol!cant_v > 0 Then
                                   sprEzeB(0).GetText 3, 1, valor
                                   sprEzeB(0).SetText 3, 5, str(valor / RsDataVol!cant_v)
                                   sprEzeB(0).GetText 3, 2, valor
                                   sprEzeB(0).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                                End If
                            Case Left(DSLoc(3).locales(3), 3)
                                sprEzeB(1).SetText 3, 4, str(RsDataVol!cant_v)
                                If RsDataVol!cant_v > 0 Then
                                   sprEzeB(1).GetText 3, 1, valor
                                   sprEzeB(1).SetText 3, 5, str(valor / RsDataVol!cant_v)
                                   sprEzeB(1).GetText 3, 2, valor
                                   sprEzeB(1).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                                End If
                            Case Left(DSLoc(3).locales(4), 3)
                                sprEzeB(2).SetText 3, 4, str(RsDataVol!cant_v)
                                If RsDataVol!cant_v > 0 Then
                                   sprEzeB(2).GetText 3, 1, valor
                                   sprEzeB(2).SetText 3, 5, str(valor / RsDataVol!cant_v)
                                   sprEzeB(2).GetText 3, 2, valor
                                   sprEzeB(2).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                                End If
                        
                        End Select
                 End Select

       Case DSLoc(4).Dep
            Select Case RsDataVol!local
                Case Left(DSLoc(4).locales(1), 3)
                    sprInt(0).SetText 3, 4, str(RsDataVol!cant_v)
                    If RsDataVol!cant_v > 0 Then
                       sprInt(0).GetText 3, 1, valor
                       sprInt(0).SetText 3, 5, str(valor / RsDataVol!cant_v)
                       sprInt(0).GetText 3, 2, valor
                       sprInt(0).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
              
                Case Left(DSLoc(5).locales(1), 3)
                    sprInt(2).SetText 3, 4, str(RsDataVol!cant_v)
                    If RsDataVol!cant_v > 0 Then
                       sprInt(2).GetText 3, 1, valor
                       sprInt(2).SetText 3, 5, str(valor / RsDataVol!cant_v)
                       sprInt(2).GetText 3, 2, valor
                       sprInt(2).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
                
                Case Left(DSLoc(6).locales(1), 3)
                    sprInt(1).SetText 3, 4, str(RsDataVol!cant_v)
                    If RsDataVol!cant_v > 0 Then
                       sprInt(1).GetText 3, 1, valor
                       sprInt(1).SetText 3, 5, str(valor / RsDataVol!cant_v)
                       sprInt(1).GetText 3, 2, valor
                       sprInt(1).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                     End If
                Case Left(DSLoc(6).locales(2), 3)
                    sprInt(4).SetText 3, 4, str(RsDataVol!cant_v)
                    If RsDataVol!cant_v > 0 Then
                       sprInt(4).GetText 3, 1, valor
                       sprInt(4).SetText 3, 5, str(valor / RsDataVol!cant_v)
                       sprInt(4).GetText 3, 2, valor
                       sprInt(4).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
                Case Left(DSLoc(7).locales(1), 3)
                    sprInt(3).SetText 3, 4, str(RsDataVol!cant_v)
                    If RsDataVol!cant_v > 0 Then
                       sprInt(3).GetText 3, 1, valor
                       sprInt(3).SetText 3, 5, str(valor / RsDataVol!cant_v)
                       sprInt(3).GetText 3, 2, valor
                       sprInt(3).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
                Case Left(DSLoc(8).locales(1), 3)
                    sprInt(5).SetText 3, 4, str(RsDataVol!cant_v)
                    If RsDataVol!cant_v > 0 Then
                       sprInt(5).GetText 3, 1, valor
                       sprInt(5).SetText 3, 5, str(valor / RsDataVol!cant_v)
                       sprInt(5).GetText 3, 2, valor
                       sprInt(5).SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                    End If
            End Select

  End Select
        RsDataVol.MoveNext
Loop

End Sub


Private Sub L_TratarVoladosHIST()
Dim SQL As String
Dim FD As String, FH As String


If optFechas(0).Value Then
    FD = "01-01-" & Year(mskFDesde.FormattedText) - 1
    FH = Day(mskFDesde.FormattedText) & "-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText) - 1
Else
    FD = mskFDesdeAnt.FormattedText
    FH = mskFHastaAnt.FormattedText
End If
    
        SQL = " SELECT "
        SQL = SQL & "cod_depn,"
        SQL = SQL & "cod_sdep,"
        SQL = SQL & "local,"
        SQL = SQL & " SUM(pax_volados.cantidad) CANT_V "
        SQL = SQL & " From estadis.pax_volados "
        SQL = SQL & " Where "
        SQL = SQL & " fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
        SQL = SQL & " GROUP BY COD_DEPN,COD_SDEP,local"
        
        If Aplicacion.ObtenerRsDAO(SQL, RsDataVol) Then
            L_DecoVolados_Ant
            Aplicacion.CerrarDAO RsDataVol
        End If

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



Private Sub botExcelA_Click()
L_TratarExcel sprEzeA(0), EZEA, "ESPIGON INTERNACIONAL 'A'"
End Sub

Private Sub botExcelB_Click()
L_TratarExcel sprEzeB(0), EZEB, "ESPIGON INTERNACIONAL 'B'"
End Sub


Private Sub botExcelE_Click()
L_TratarExcel sprAep(0), AERO, "AEROPARQUE"
End Sub


Private Sub botExcelI_Click()
L_TratarExcel sprInt(0), INTE, "INTERIOR"
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
FuncLocal_SeteoTABS tabEspigon

End Sub

Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6300
Me.Width = 9300

If Day(Date) = 1 Then
    mskFDesde.Text = Format$(Func.func_Dia1SegunMes_Anio(Month(Date - 1), Year(Date - 1)), FTOFECHA)
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
    
    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio - 1, FTOFECHA)
    mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio - 1, FTOFECHA)
    
Else
    mskFDesde.Text = "01-" & Format$(Month(Date), "0#") & "-" & Format$(Year(Date), "####")
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
    
    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
    mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio, FTOFECHA)
    
End If
    

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
ElseIf CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
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

Private Sub optFechas_Click(Index As Integer)
Select Case Index
    Case 0
        mskFHasta.Visible = False
        botHelpFH.Visible = False
        Label1(0).caption = "Fecha Tope"
        Label1(1).Visible = False
        frPerAnt.Visible = False
    Case 1
        mskFHasta.Visible = True
        botHelpFH.Visible = True
        Label1(0).caption = "Fecha Desde"
        Label1(1).Visible = True
        frPerAnt.Visible = True
End Select
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
labInfo.Visible = Not labInfo.Visible
End Sub


