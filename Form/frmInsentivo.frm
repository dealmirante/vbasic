VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInsentivo 
   Caption         =   "Resumen de Estimados vs. Ventas"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   8235
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   3150
      TabIndex        =   34
      Top             =   -30
      Width           =   4830
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
         Height          =   435
         Index           =   2
         Left            =   1845
         Picture         =   "frmInsentivo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   330
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
         Left            =   1170
         Picture         =   "frmInsentivo.frx":0822
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   345
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
         Left            =   495
         Picture         =   "frmInsentivo.frx":0924
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   345
         Width           =   570
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4485
      Left            =   30
      TabIndex        =   7
      Top             =   1125
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   7911
      _Version        =   327680
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   423
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
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmInsentivo.frx":0A26
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "Int 'A'"
      TabPicture(1)   =   "frmInsentivo.frx":0A42
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabGrRub"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "AEP"
      TabPicture(2)   =   "frmInsentivo.frx":0A5E
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabGrIntB"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "Aep"
      TabPicture(3)   =   "frmInsentivo.frx":0A7A
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tabGrAep"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "Resultados"
      TabPicture(4)   =   "frmInsentivo.frx":0A96
      Tab(4).ControlCount=   4
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label1(13)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "txtdifcia"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame4"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "botPers(3)"
      Tab(4).Control(3).Enabled=   0   'False
      Begin VB.CommandButton botPers 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  Personal          CIA"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   6375
         TabIndex        =   57
         Top             =   300
         Width           =   1050
      End
      Begin VB.Frame Frame4 
         Caption         =   "Comisión correspondiente por Espigón-Grupo-Rubro"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3465
         Left            =   165
         TabIndex        =   49
         Top             =   885
         Width           =   7275
         Begin FPSpread.vaSpread sprInsAep 
            Height          =   1425
            Left            =   4920
            OleObjectBlob   =   "frmInsentivo.frx":0AB2
            TabIndex        =   53
            Top             =   1440
            Width           =   2235
         End
         Begin FPSpread.vaSpread sprInsB 
            Height          =   1425
            Left            =   2520
            OleObjectBlob   =   "frmInsentivo.frx":0E43
            TabIndex        =   52
            Top             =   1440
            Width           =   2235
         End
         Begin FPSpread.vaSpread sprInsA 
            Height          =   1425
            Left            =   90
            OleObjectBlob   =   "frmInsentivo.frx":1202
            TabIndex        =   50
            Top             =   1440
            Width           =   2235
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1095
            Left            =   150
            TabIndex        =   61
            Top             =   210
            Width           =   6975
            Begin VB.TextBox txtRI 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   6150
               TabIndex        =   79
               Top             =   765
               Width           =   720
            End
            Begin VB.TextBox txtRB 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   3645
               TabIndex        =   78
               Top             =   765
               Width           =   705
            End
            Begin VB.TextBox txtRA 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1290
               TabIndex        =   77
               Top             =   780
               Width           =   660
            End
            Begin VB.TextBox txtPA 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   165
               Locked          =   -1  'True
               TabIndex        =   64
               Top             =   795
               Width           =   705
            End
            Begin VB.TextBox txtPB 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   2490
               Locked          =   -1  'True
               TabIndex        =   63
               Top             =   750
               Width           =   705
            End
            Begin VB.TextBox txtPAep 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   5055
               Locked          =   -1  'True
               TabIndex        =   62
               Top             =   765
               Width           =   705
            End
            Begin MSMask.MaskEdBox txtB 
               Height          =   300
               Left            =   3930
               TabIndex        =   65
               Top             =   15
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   529
               _Version        =   327680
               ForeColor       =   4210752
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "% #,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtAep 
               Height          =   300
               Left            =   6330
               TabIndex        =   66
               Top             =   0
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   529
               _Version        =   327680
               ForeColor       =   4210752
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "% #,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtA 
               Height          =   300
               Left            =   1500
               TabIndex        =   76
               Top             =   15
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   529
               _Version        =   327680
               ForeColor       =   4210752
               PromptInclude   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "% #,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Inter ""A""  (40%)"
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
               Index           =   12
               Left            =   0
               TabIndex        =   75
               Top             =   15
               Width           =   1410
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "AEP  (40%)"
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
               Index           =   14
               Left            =   2430
               TabIndex        =   74
               Top             =   15
               Width           =   1410
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Interior (40%)"
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
               Index           =   15
               Left            =   4830
               TabIndex        =   73
               Top             =   0
               Width           =   1410
            End
            Begin VB.Label Label2 
               Caption         =   "General"
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
               Index           =   0
               Left            =   2265
               TabIndex        =   72
               Top             =   510
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "General"
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
               Index           =   1
               Left            =   0
               TabIndex        =   71
               Top             =   510
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "General"
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
               Index           =   2
               Left            =   4860
               TabIndex        =   70
               Top             =   450
               Width           =   855
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   3
               Left            =   3255
               TabIndex        =   69
               Top             =   735
               Width           =   330
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   435
               Index           =   4
               Left            =   960
               TabIndex        =   68
               Top             =   720
               Width           =   330
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   5
               Left            =   5805
               TabIndex        =   67
               Top             =   735
               Width           =   330
            End
         End
         Begin VB.CommandButton botPers 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Personal AEP"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   4920
            TabIndex        =   56
            Top             =   2910
            Width           =   2235
         End
         Begin VB.CommandButton botPers 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Personal Int ""B"""
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   2520
            TabIndex        =   55
            Top             =   2895
            Width           =   2235
         End
         Begin VB.CommandButton botPers 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Personal Int ""A"""
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   105
            TabIndex        =   54
            Top             =   2895
            Width           =   2235
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   1065
         Left            =   -74445
         TabIndex        =   24
         Top             =   405
         Width           =   6420
         Begin MSMask.MaskEdBox mskReal 
            Height          =   285
            Left            =   1545
            TabIndex        =   25
            Top             =   615
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   503
            _Version        =   327680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$ #,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskEst 
            Height          =   285
            Left            =   3225
            TabIndex        =   26
            Top             =   615
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   503
            _Version        =   327680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$ #,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDif 
            Height          =   300
            Left            =   4905
            TabIndex        =   27
            Top             =   600
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            _Version        =   327680
            ForeColor       =   16711680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "% #,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Venta Real"
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
            Left            =   1545
            TabIndex        =   31
            Top             =   240
            Width           =   1665
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Venta Estimada"
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
            Left            =   3225
            TabIndex        =   30
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diferencia"
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
            Left            =   4905
            TabIndex        =   29
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Compañía"
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
            Index           =   8
            Left            =   240
            TabIndex        =   28
            Top             =   600
            Width           =   1230
         End
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   1905
         Left            =   -74445
         TabIndex        =   8
         Top             =   1500
         Width           =   6435
         Begin MSMask.MaskEdBox mskRealEsp 
            Height          =   285
            Index           =   0
            Left            =   1530
            TabIndex        =   9
            Top             =   690
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   503
            _Version        =   327680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$ #,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskEstimEsp 
            Height          =   285
            Index           =   0
            Left            =   3210
            TabIndex        =   10
            Top             =   690
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   503
            _Version        =   327680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$ #,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDifEsp 
            Height          =   300
            Index           =   0
            Left            =   4905
            TabIndex        =   11
            Top             =   675
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            _Version        =   327680
            ForeColor       =   16711680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "% #,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskRealEsp 
            Height          =   285
            Index           =   1
            Left            =   1530
            TabIndex        =   12
            Top             =   1065
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   503
            _Version        =   327680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$ #,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskEstimEsp 
            Height          =   285
            Index           =   1
            Left            =   3225
            TabIndex        =   13
            Top             =   1065
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   503
            _Version        =   327680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$ #,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDifEsp 
            Height          =   300
            Index           =   1
            Left            =   4905
            TabIndex        =   14
            Top             =   1050
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   327680
            ForeColor       =   16711680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "% #,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskRealEsp 
            Height          =   285
            Index           =   2
            Left            =   1530
            TabIndex        =   15
            Top             =   1440
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   503
            _Version        =   327680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$ #,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskEstimEsp 
            Height          =   285
            Index           =   2
            Left            =   3225
            TabIndex        =   16
            Top             =   1440
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   503
            _Version        =   327680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$ #,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDifEsp 
            Height          =   300
            Index           =   2
            Left            =   4905
            TabIndex        =   17
            Top             =   1425
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   327680
            ForeColor       =   16711680
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "% #,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diferencia"
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
            Left            =   4890
            TabIndex        =   23
            Top             =   285
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Venta Estimada"
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
            Left            =   3210
            TabIndex        =   22
            Top             =   285
            Width           =   1650
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Venta Real"
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
            Left            =   1530
            TabIndex        =   21
            Top             =   285
            Width           =   1665
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "EZE"
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
            Index           =   9
            Left            =   255
            TabIndex        =   20
            Top             =   675
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AEP"
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
            Index           =   10
            Left            =   255
            TabIndex        =   19
            Top             =   1035
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "INT"
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
            Index           =   11
            Left            =   255
            TabIndex        =   18
            Top             =   1410
            Width           =   1185
         End
      End
      Begin TabDlg.SSTab tabGrRub 
         Height          =   2490
         Left            =   -74505
         TabIndex        =   32
         Top             =   660
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   4392
         _Version        =   327680
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   423
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Grupo ""A"""
         TabPicture(0)   =   "frmInsentivo.frx":15C1
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprGR(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Grupo ""B"""
         TabPicture(1)   =   "frmInsentivo.frx":15DD
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprGR(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Grupo ""C"""
         TabPicture(2)   =   "frmInsentivo.frx":15F9
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprGR(2)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "TOTAL"
         TabPicture(3)   =   "frmInsentivo.frx":1615
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "sprTot(0)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprGR 
            Height          =   1965
            Index           =   0
            Left            =   -74865
            OleObjectBlob   =   "frmInsentivo.frx":1631
            TabIndex        =   33
            Top             =   390
            Width           =   6315
         End
         Begin FPSpread.vaSpread sprGR 
            Height          =   1920
            Index           =   1
            Left            =   -74880
            OleObjectBlob   =   "frmInsentivo.frx":1A5C
            TabIndex        =   38
            Top             =   390
            Width           =   6315
         End
         Begin FPSpread.vaSpread sprGR 
            Height          =   1920
            Index           =   2
            Left            =   -74880
            OleObjectBlob   =   "frmInsentivo.frx":1E87
            TabIndex        =   39
            Top             =   405
            Width           =   6315
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   1920
            Index           =   0
            Left            =   105
            OleObjectBlob   =   "frmInsentivo.frx":22B2
            TabIndex        =   58
            Top             =   405
            Width           =   6315
         End
      End
      Begin TabDlg.SSTab tabGrIntB 
         Height          =   2490
         Left            =   -74550
         TabIndex        =   40
         Top             =   675
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   4392
         _Version        =   327680
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   423
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Grupo ""A"""
         TabPicture(0)   =   "frmInsentivo.frx":26DD
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprGR(3)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Grupo ""B"""
         TabPicture(1)   =   "frmInsentivo.frx":26F9
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprGR(4)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Grupo ""C"""
         TabPicture(2)   =   "frmInsentivo.frx":2715
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprGR(5)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "TOTAL"
         TabPicture(3)   =   "frmInsentivo.frx":2731
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "sprTot(1)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprGR 
            Height          =   1965
            Index           =   5
            Left            =   -74895
            OleObjectBlob   =   "frmInsentivo.frx":274D
            TabIndex        =   43
            Top             =   390
            Width           =   6315
         End
         Begin FPSpread.vaSpread sprGR 
            Height          =   1965
            Index           =   4
            Left            =   -74880
            OleObjectBlob   =   "frmInsentivo.frx":2B78
            TabIndex        =   42
            Top             =   390
            Width           =   6315
         End
         Begin FPSpread.vaSpread sprGR 
            Height          =   1965
            Index           =   3
            Left            =   -74865
            OleObjectBlob   =   "frmInsentivo.frx":2FA3
            TabIndex        =   41
            Top             =   390
            Width           =   6315
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   1920
            Index           =   1
            Left            =   105
            OleObjectBlob   =   "frmInsentivo.frx":33CE
            TabIndex        =   59
            Top             =   405
            Width           =   6315
         End
      End
      Begin TabDlg.SSTab tabGrAep 
         Height          =   2490
         Left            =   -74520
         TabIndex        =   44
         Top             =   735
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   4392
         _Version        =   327680
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   423
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Grupo ""A"""
         TabPicture(0)   =   "frmInsentivo.frx":37F9
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprGR(6)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "Grupo ""B"""
         TabPicture(1)   =   "frmInsentivo.frx":3815
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprGR(7)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Grupo ""C"""
         TabPicture(2)   =   "frmInsentivo.frx":3831
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprGR(8)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "TOTAL"
         TabPicture(3)   =   "frmInsentivo.frx":384D
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "sprTot(2)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprGR 
            Height          =   1965
            Index           =   8
            Left            =   -74880
            OleObjectBlob   =   "frmInsentivo.frx":3869
            TabIndex        =   47
            Top             =   375
            Width           =   6315
         End
         Begin FPSpread.vaSpread sprGR 
            Height          =   1965
            Index           =   7
            Left            =   -74880
            OleObjectBlob   =   "frmInsentivo.frx":3C94
            TabIndex        =   46
            Top             =   390
            Width           =   6315
         End
         Begin FPSpread.vaSpread sprGR 
            Height          =   1965
            Index           =   6
            Left            =   -74895
            OleObjectBlob   =   "frmInsentivo.frx":40BF
            TabIndex        =   45
            Top             =   390
            Width           =   6315
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   1920
            Index           =   2
            Left            =   105
            OleObjectBlob   =   "frmInsentivo.frx":44EA
            TabIndex        =   60
            Top             =   390
            Width           =   6315
         End
      End
      Begin MSMask.MaskEdBox txtdifcia 
         Height          =   315
         Left            =   5280
         TabIndex        =   51
         Top             =   420
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   327680
         ForeColor       =   16711680
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "% #,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diferencia entre Venta Real vs. Estimado a nivel COMPAÑIA"
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
         Index           =   13
         Left            =   150
         TabIndex        =   48
         Top             =   420
         Width           =   5100
      End
   End
   Begin VB.Frame frdatos 
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
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   3015
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2550
         Picture         =   "frmInsentivo.frx":4915
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   255
         Width           =   375
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   2535
         Picture         =   "frmInsentivo.frx":4A87
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   615
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1395
         TabIndex        =   3
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
         TabIndex        =   4
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
         Index           =   4
         Left            =   150
         TabIndex        =   6
         Top             =   270
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
         TabIndex        =   5
         Top             =   615
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmInsentivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsEstim As Recordset
Dim rsEstimGr As Recordset

Dim rsRealRub As Recordset

Private Sub L_CalculoDif()
Dim dif
Dim i As Integer, j

Spread.Spread_TotalesGrillas SprTot(0), SprTot(0).MaxRows - 1, 2
Spread.Spread_TotalesGrillas SprTot(1), SprTot(1).MaxRows - 1, 2
Spread.Spread_TotalesGrillas SprTot(2), SprTot(2).MaxRows - 1, 2

dif = Val(mskReal.Text) - Val(mskEst.Text)

If Val(mskEst.Text) > 0 Then
    mskDif.Text = dif * 100 / Val(mskEst.Text)
    txtdifcia.Text = mskDif.Text
End If

If Val(mskDif.Text) < 0 Then
    mskDif.ForeColor = &HFF&
    txtdifcia.ForeColor = &HFF&
Else
    mskDif.ForeColor = &HFF0000
    txtdifcia.ForeColor = &HFF0000
End If

For i = 0 To 2
    dif = Val(mskRealEsp(i).Text) - Val(mskEstimEsp(i).Text)
    If Val(mskEstimEsp(i).Text) > 0 Then
        mskDifEsp(i).Text = dif * 100 / Val(mskEstimEsp(i).Text)
    End If
    
    If Val(mskDifEsp(i).Text) < 0 Then
        mskDifEsp(i).ForeColor = &HFF&
    Else
        mskDifEsp(i).ForeColor = &HFF0000
    End If


Next

For i = 1 To sprGr(0).MaxRows
    For j = 0 To 8
        spread_ResaltarCelda sprGr(j), 4, i
        spread_ResaltarCelda sprGr(j), 5, i
    Next
Next
For i = 1 To SprTot(0).MaxRows
    For j = 0 To 2
     spread_ResaltarCelda SprTot(j), 4, i
     spread_ResaltarCelda SprTot(j), 5, i
    Next
Next

End Sub

Private Function L_InsentivoEsp(dato As Single) As Single

If dato >= -5 And dato < 1.5 Then
    L_InsentivoEsp = 5
ElseIf dato >= 1.51 And dato < 3 Then
    L_InsentivoEsp = 7
ElseIf dato >= 3.1 And dato < 5 Then
    L_InsentivoEsp = 9
ElseIf dato >= 5.1 And dato < 7 Then
    L_InsentivoEsp = 10
ElseIf dato >= 7.1 And dato < 10 Then
    L_InsentivoEsp = 15
ElseIf dato >= 10.1 And dato < 15 Then
    L_InsentivoEsp = 20
ElseIf dato >= 15.1 And dato < 20 Then
    L_InsentivoEsp = 25
ElseIf dato >= 20 Then
    L_InsentivoEsp = 30
Else
    L_InsentivoEsp = 0
End If

End Function
Private Function L_InsentivoRub(dato As Single) As Single

If dato >= 0 And dato < 3 Then
    L_InsentivoRub = 50
ElseIf dato >= 3.1 And dato < 6 Then
    L_InsentivoRub = 75
ElseIf dato >= 6.1 Then
    L_InsentivoRub = 100
Else
L_InsentivoRub = 0

End If

End Function

Private Sub L_LimpiarGrillas()
Dim i
    For i = 0 To 8
        sprGr(i).MaxRows = 0
    Next
    mskDif.Text = 0
    mskEst.Text = 0
    mskReal.Text = 0
    For i = 0 To 2
        mskDifEsp(i).Text = 0
        mskEstimEsp(i).Text = 0
        mskRealEsp(i).Text = 0
        SprTot(i).MaxRows = 0
    Next
    sprInsA.MaxRows = 0
    sprInsB.MaxRows = 0
    sprInsAep.MaxRows = 0
    txtA.Text = ""
    txtB.Text = ""
    txtAep.Text = ""
    txtPA.Text = ""
    txtPB.Text = ""
    txtPAep.Text = ""
    txtdifcia.Text = ""
End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim RS As Recordset

'On Error GoTo Err:

sql = " SELECT "
sql = sql & " cod_depn, "
'sql = sql & " cod_sdep, "
sql = sql & " sum(importe) imp "
sql = sql & " FROM " & funcLocal_Vista("venta_lgi", Year(mskFDesde.FormattedText))
sql = sql & " WHERE fch_ticket BETWEEN " & Func.func_ToDate(mskFDesde.FormattedText)
sql = sql & " And " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " And comitente = 'T' "
sql = sql & " group by cod_depn"
sql = sql & " order by cod_depn"

mskReal.Text = 0
mskRealEsp(0).Text = 0
mskRealEsp(1).Text = 0
mskRealEsp(2).Text = 0

If Aplicacion.ObtenerRsDAO(sql, RS) Then
    If Aplicacion.CantReg(RS) > 0 Then
        Do While Not RS.EOF
            Select Case RS!cod_depn
                Case "EZE"
                    mskRealEsp(0).Text = RS!IMP
                Case "AEP"
                    mskRealEsp(1).Text = RS!IMP
                Case "INT"
                    mskRealEsp(2).Text = RS!IMP
            End Select
            mskReal.Text = Val(mskReal.Text) + RS!IMP
            RS.MoveNext
        Loop
    
        sql = " SELECT "
        sql = sql & " cod_depn,grupo_venta grupo, "
        sql = sql & " decode (cod_rubr,'ACC','ACC/ELE', "
        sql = sql & " decode (cod_rubr,'ELE','ACC/ELE', "
        sql = sql & " decode(cod_rubr,'TAB','BEB/CIG/COM', "
        sql = sql & " decode(cod_rubr,'BEB','BEB/CIG/COM', "
        sql = sql & " decode(cod_rubr,'COM','BEB/CIG/COM', "
        sql = sql & " decode(cod_rubr,'CIG','BEB/CIG/COM','COS/PER')))))) cod_rubr , "
        sql = sql & " sum(importe) imp "
        sql = sql & " FROM " & funcLocal_Vista("venta_plg", Year(mskFDesde.FormattedText))
        sql = sql & " WHERE fch_ticket BETWEEN " & Func.func_ToDate(mskFDesde.FormattedText)
        sql = sql & " And " & Func.func_ToDate(mskFHasta.FormattedText)
        sql = sql & " group by cod_depn,grupo_venta, "
        sql = sql & " decode (cod_rubr,'ACC','ACC/ELE', "
        sql = sql & " decode (cod_rubr,'ELE','ACC/ELE', "
        sql = sql & " decode(cod_rubr,'TAB','BEB/CIG/COM', "
        sql = sql & " decode(cod_rubr,'BEB','BEB/CIG/COM', "
        sql = sql & " decode(cod_rubr,'COM','BEB/CIG/COM', "
        sql = sql & " decode(cod_rubr,'CIG','BEB/CIG/COM','COS/PER')))))) "
        sql = sql & " order by cod_rubr,cod_depn,grupo_venta "
        
        If Aplicacion.ObtenerRsDAO(sql, rsRealRub) Then
            
            If Aplicacion.CantReg(rsRealRub) > 0 Then
                L_LlenarGrillasRub
                'tabEst.Enabled = True
                'botEjecutar(0).Enabled = False
                'frCab.Enabled = False
                'frMod.Enabled = False
            End If
        
            Aplicacion.CerrarDAO rsRealRub
        
        End If
    End If
    
    Aplicacion.CerrarDAO RS

End If

Err:
    Exit Sub
End Sub


Private Sub L_Resultados__()
'No se usa pero por las dudas .....
Dim ind As Integer
Dim i, GR
Dim valor As Variant
Dim Porc As Single

sprInsA.MaxRows = 0
sprInsB.MaxRows = 0
sprInsAep.MaxRows = 0
'If Val(txtdifcia.Text) > 0 Then
   txtA.Text = L_InsentivoEsp(Val(mskDifEsp(0).Text))

    If Val(txtA.Text) > 0 Then
        GR = 65
        For ind = 0 To 2
            For i = 1 To sprGr(ind).MaxRows
                sprGr(ind).GetText 5, i, valor
                Porc = L_InsentivoRub(Val(valor))
                If Porc > 0 Then
                    sprInsA.MaxRows = sprInsA.MaxRows + 1
                    sprInsA.SetText 1, sprInsA.MaxRows, Chr(GR)
                    sprGr(ind).GetText 1, i, valor
                    sprInsA.SetText 2, sprInsA.MaxRows, Trim(valor)
                    sprInsA.SetText 3, sprInsA.MaxRows, str(Porc)
                End If
            Next
            GR = GR + 1
        Next
    End If
    '-----------------------
    txtB.Text = L_InsentivoEsp(Val(mskDifEsp(1).Text))
    If Val(txtB.Text) > 0 Then
        GR = 65
        For ind = 3 To 5
            For i = 1 To sprGr(ind).MaxRows
                sprGr(ind).GetText 5, i, valor
                Porc = L_InsentivoRub(Val(valor))
                If Porc > 0 Then
                    sprInsB.MaxRows = sprInsB.MaxRows + 1
                    sprInsB.SetText 1, sprInsB.MaxRows, Chr(GR)
                    sprGr(ind).GetText 1, i, valor
                    sprInsB.SetText 2, sprInsB.MaxRows, Trim(valor)
                    sprInsB.SetText 3, sprInsB.MaxRows, str(Porc)
                End If
                
            Next
            GR = GR + 1
        Next
    End If
    '-----------------------
    txtAep.Text = L_InsentivoEsp(Val(mskDifEsp(2).Text))
    If Val(txtAep.Text) > 0 Then
        GR = 65
        For ind = 6 To 8
            For i = 1 To sprGr(ind).MaxRows
                sprGr(ind).GetText 5, i, valor
                Porc = L_InsentivoRub(Val(valor))
                If Porc > 0 Then
                    sprInsAep.MaxRows = sprInsAep.MaxRows + 1
                    sprInsAep.SetText 1, sprInsAep.MaxRows, Chr(GR)
                    sprGr(ind).GetText 1, i, valor
                    sprInsAep.SetText 2, sprInsAep.MaxRows, Trim(valor)
                    sprInsAep.SetText 3, sprInsAep.MaxRows, str(Porc)
                End If
                
            Next
            GR = GR + 1
        Next
    End If

'Else
'    txtA.Text = 0
'    txtB.Text = 0
'    txtAep.Text = 0
'End If

End Sub

Private Sub L_Resultados()
Dim ind As Integer
Dim i
Dim valor As Variant
Dim Porc As Single, Prm As Single

sprInsA.MaxRows = 0
sprInsB.MaxRows = 0
sprInsAep.MaxRows = 0
If Val(txtdifcia.Text) > -10 Then
   txtA.Text = L_InsentivoEsp(Val(mskDifEsp(0).Text))
   txtPA.Text = 40 * txtA.Text / 100

    'If Val(txtA.Text) > 0 Then
        Prm = 0
        For i = 1 To SprTot(ind).MaxRows - 1
            SprTot(0).GetText 5, i, valor
            'Porc = L_InsentivoRub(Val(valor))
            Porc = Val(valor)
            Prm = Prm + Porc
            'If Porc > 0 Then
                sprInsA.MaxRows = sprInsA.MaxRows + 1
                SprTot(0).GetText 1, i, valor
                sprInsA.SetText 2, sprInsA.MaxRows, Trim(valor)
                sprInsA.SetText 4, sprInsA.MaxRows, str(Porc)
                'sprInsA.SetText 4, sprInsA.MaxRows, str(Porc * txtA.Text * 60 / 10000)
            'End If
        Next
        txtRA.Text = txtA.Text * 60 * L_InsentivoRub(Prm / 3) / 10000
        'sprInsA.MaxRows = sprInsA.MaxRows + 1
        'sprInsA.SetText 2, sprInsA.MaxRows, "CAJ"
        'sprInsA.SetText 3, sprInsA.MaxRows, L_InsentivoRub(Val(mskDifEsp(0).Text))
        'sprInsA.SetText 4, sprInsA.MaxRows, L_InsentivoRub(Val(mskDifEsp(0).Text)) * 60 * txtA.Text / 10000
    'End If
    '-----------------------
    txtB.Text = L_InsentivoEsp(Val(mskDifEsp(1).Text))
    txtPB.Text = 40 * txtB.Text / 100
    'If Val(txtB.Text) > 0 Then
        Prm = 0
        For i = 1 To SprTot(ind).MaxRows - 1
            SprTot(1).GetText 5, i, valor
            'Porc = L_InsentivoRub(Val(valor))
            Porc = Val(valor)
            Prm = Prm + Porc
            'If Porc > 0 Then
                sprInsB.MaxRows = sprInsB.MaxRows + 1
                SprTot(1).GetText 1, i, valor
                sprInsB.SetText 2, sprInsB.MaxRows, Trim(valor)
                sprInsB.SetText 4, sprInsB.MaxRows, str(Porc)
                'sprInsB.SetText 4, sprInsB.MaxRows, str(Porc * txtB.Text * 60 / 10000)
            'End If
        Next
        txtRB.Text = txtB.Text * 60 * L_InsentivoRub(Prm / 3) / 10000
        'sprInsB.MaxRows = sprInsB.MaxRows + 1
        'sprInsB.SetText 2, sprInsB.MaxRows, "CAJ"
        'sprInsB.SetText 3, sprInsB.MaxRows, L_InsentivoRub(Val(mskDifEsp(1).Text))
        'sprInsB.SetText 4, sprInsB.MaxRows, str(L_InsentivoRub(Val(mskDifEsp(1).Text)) * 60 * txtB.Text / 10000)
    'End If
    '----------------------- INTERIOR
    txtAep.Text = L_InsentivoEsp(Val(mskDifEsp(2).Text))
    txtPAep.Text = 40 * txtAep.Text / 100
    txtRI.Text = Val(txtAep.Text) - Val(txtPAep.Text)
    'If Val(txtAep.Text) > 0 Then
        'For i = 1 To SprTot(ind).MaxRows - 1
            'SprTot(2).GetText 5, i, valor
            'Porc = L_InsentivoRub(Val(valor))
            'If Porc > 0 Then
            '    sprInsAep.MaxRows = sprInsAep.MaxRows + 1
            '    SprTot(2).GetText 1, i, valor
            '    sprInsAep.SetText 2, sprInsAep.MaxRows, Trim(valor)
            '    sprInsAep.SetText 3, sprInsAep.MaxRows, str(Porc)
            '    sprInsAep.SetText 4, sprInsAep.MaxRows, str(Porc * txtAep.Text * 60 / 10000)
            'End If
        'Next
        'sprInsAep.MaxRows = sprInsAep.MaxRows + 1
        'sprInsAep.SetText 2, sprInsAep.MaxRows, "CAJ"
        'sprInsAep.SetText 3, sprInsAep.MaxRows, L_InsentivoRub(Val(mskDifEsp(2).Text))
        'sprInsAep.SetText 4, sprInsAep.MaxRows, str(L_InsentivoRub(Val(mskDifEsp(2).Text)) * 60 * txtAep.Text / 10000)
    'End If
Else
    txtA.Text = 0
    txtB.Text = 0
    txtAep.Text = 0
End If

End Sub


Private Sub L_ResultadosCajero()
Dim i
Dim valor As Variant


End Sub


Private Sub L_SumatoriaGrillas(ind As Integer, sprgrT As control)
Dim i
Dim valor As Variant, tot As Single

tot = 0
For i = 1 To sprGr(ind).MaxRows
    sprGr(ind).GetText 1, i, valor
    sprgrT.MaxRows = sprgrT.MaxRows + 1
    sprgrT.SetText 1, i, Trim(valor)
    
    sprGr(ind).GetText 2, i, valor
    tot = valor
    sprGr(ind + 1).GetText 2, i, valor
    tot = tot + valor
    sprGr(ind + 2).GetText 2, i, valor
    tot = tot + valor
    sprgrT.SetText 2, i, str(tot)
    
    sprGr(ind).GetText 3, i, valor
    tot = valor
    sprGr(ind + 1).GetText 3, i, valor
    tot = tot + valor
    sprGr(ind + 2).GetText 3, i, valor
    tot = tot + valor
    
    sprgrT.SetText 3, i, str(tot)

Next

End Sub

Private Sub botEjecutar_Click(Index As Integer)


Select Case Index
    Case 0
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        frmInsentivo.caption = Aplicacion.SeteoProceso(frmInsentivo.caption)
        L_RefrescarEstimRub
        L_Refrescar
        L_CalculoDif
        L_Resultados
        frmInsentivo.caption = Aplicacion.SeteoFin
    Case 1
        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        'tabEspigon.Enabled = False
        L_LimpiarGrillas
    Case 2
        Unload Me
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

Private Sub botPers_Click(Index As Integer)
Dim sql As String
Dim i, valor As Variant

Select Case Index
    Case 0
        If sprInsA.MaxRows > 0 Then
            frmGriPersonal.MostrarPersonal "INTA", Val(txtA.Text), sprInsA
        End If
    Case 1
        If sprInsB.MaxRows > 0 Then
            frmGriPersonal.MostrarPersonal "INTB", Val(txtB.Text), sprInsB
        End If
    Case 2
        If sprInsAep.MaxRows > 0 Then
            frmGriPersonal.MostrarPersonal "AEP", Val(txtAep.Text), sprInsAep
        End If
    
    Case 3
        frmGriPersonal.MostrarPersonalCIA "INTA", Val(txtA.Text), sprInsA, _
                                          "INTB", Val(txtB.Text), sprInsB, _
                                          "AEP", Val(txtAep.Text), sprInsAep


End Select
End Sub

Private Sub Form_Load()

Width = Screen.Width * 0.81
Height = Screen.Height * 0.78
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height - 500) / 2

mskFDesde.Text = "01-" & Format$(Month(Date), "00") & "-" & Format$(Year(Date), "####")
mskFHasta.Text = Format$(Date - 1, FTOFECHA)

SSTab1.TabVisible(3) = False
L_LimpiarGrillas

End Sub

Private Sub mskFDesde_LostFocus()

    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    'else
    'If (Year(mskFDesde.FormattedText) < Year(Date)) Then
    '(Month(mskFDesde.FormattedText) < Month(Date))
    '    mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    End If
    
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)

End Sub


Private Sub mskFHasta_LostFocus()

If Not IsDate(mskFHasta.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
ElseIf CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) _
        And Month(mskFHasta.FormattedText) <> Month(mskFHasta.FormattedText) Then
        
        mskFHasta.Text = mskFDesde.Text
End If

mskFHasta.Text = Format$(mskFHasta.FormattedText, FTOFECHA)

End Sub


Private Sub L_RefrescarEstimRub()
Dim sql As String

On Error GoTo ErrEstim:

sql = "Select depn,"
'sql = sql & " sdep,"
sql = sql & " sum(porc) PORC"
sql = sql & " From estadis.porcentaje_d "
sql = sql & L_ArmarcondicionEstim
sql = sql & " AND DIA BETWEEN " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText)
sql = sql & " GROUP BY depn "
sql = sql & " ORDER BY depn "

mskEst.Text = 0
mskEstimEsp(0).Text = 0
mskEstimEsp(1).Text = 0
mskEstimEsp(2).Text = 0

If Aplicacion.ObtenerRsDAO(sql, rsEstim) Then
    If Aplicacion.CantReg(rsEstim) > 0 Then
        Do While Not rsEstim.EOF
            Select Case rsEstim!depn
                Case "EZE"
                    mskEstimEsp(0).Text = rsEstim!Porc
                Case "AEP"
                    mskEstimEsp(1).Text = rsEstim!Porc
                Case "INT"
                    mskEstimEsp(2).Text = rsEstim!Porc
            End Select
            mskEst.Text = Val(mskEst.Text) + rsEstim!Porc
            rsEstim.MoveNext
        Loop
    End If
    
    Aplicacion.CerrarDAO rsEstim

End If

sql = "Select depn,"
'sql = sql & " sdep,"
sql = sql & " grupo,"
sql = sql & " decode (rubro,'ACC','ACC/ELE', "
sql = sql & " decode (rubro,'ELE','ACC/ELE', "
sql = sql & " decode(rubro,'TAB','BEB/CIG/COM', "
sql = sql & " decode(rubro,'BEB','BEB/CIG/COM', "
sql = sql & " decode(rubro,'COM','BEB/CIG/COM', "
sql = sql & " decode(rubro,'CIG','BEB/CIG/COM','COS/PER')))))) rubro , "
sql = sql & " sum(porc) PORC"
sql = sql & " From estadis.porcentaje_dgr "
sql = sql & L_ArmarcondicionEstim
sql = sql & " AND DIA BETWEEN " & func_ToDate(mskFDesde.FormattedText) & " AND " & func_ToDate(mskFHasta.FormattedText)
sql = sql & " GROUP BY depn,grupo,"
sql = sql & " decode (rubro,'ACC','ACC/ELE', "
sql = sql & " decode (rubro,'ELE','ACC/ELE', "
sql = sql & " decode(rubro,'TAB','BEB/CIG/COM', "
sql = sql & " decode(rubro,'BEB','BEB/CIG/COM', "
sql = sql & " decode(rubro,'COM','BEB/CIG/COM', "
sql = sql & " decode(rubro,'CIG','BEB/CIG/COM','COS/PER')))))) "

sql = sql & " ORDER BY rubro,depn,grupo"

If Aplicacion.ObtenerRsDAO(sql, rsEstimGr) Then

    If Aplicacion.CantReg(rsEstimGr) > 0 Then
        L_LlenarGrillasEstimRub
    End If

    Aplicacion.CerrarDAO rsEstimGr

End If

ErrEstim:
    Exit Sub
End Sub

Private Sub L_LlenarGrillasEstimRub()
Dim rub As String, valor As Variant
Dim rsGr As Recordset
Dim sql As String
Dim i

'No hago esta consulta junto con Dia porque hay que separar los
'casos de aquellos modelos que no tienen datos por grupo

Do While Not rsEstimGr.EOF
    rub = rsEstimGr!Rubro
    For i = 0 To 8
        sprGr(i).MaxRows = sprGr(i).MaxRows + 1
        sprGr(i).SetText 1, sprGr(i).MaxRows, rub
    Next
    Do While rub = rsEstimGr!Rubro
    Select Case rsEstimGr.depn
        Case "EZE"
            Select Case rsEstimGr!Grupo
                Case "A"
                    sprGr(0).GetText 1, sprGr(0).MaxRows, valor
                    If valor = rub Then
                        sprGr(0).SetText 3, sprGr(0).MaxRows, str(rsEstimGr!Porc)
                    End If
                Case "B"
                    sprGr(1).GetText 1, sprGr(1).MaxRows, valor
                    If valor = rub Then
                        sprGr(1).SetText 3, sprGr(1).MaxRows, str(rsEstimGr!Porc)
                    End If
                Case "C"
                    sprGr(2).GetText 1, sprGr(2).MaxRows, valor
                    If valor = rub Then
                        sprGr(2).SetText 3, sprGr(2).MaxRows, str(rsEstimGr!Porc)
                    End If
            End Select
        Case "AEP"
            Select Case rsEstimGr!Grupo
                Case "A"
                    sprGr(3).GetText 1, sprGr(3).MaxRows, valor
                    If valor = rub Then
                        sprGr(3).SetText 3, sprGr(3).MaxRows, str(rsEstimGr!Porc)
                    End If
                Case "B"
                    sprGr(4).GetText 1, sprGr(4).MaxRows, valor
                    If valor = rub Then
                        sprGr(4).SetText 3, sprGr(4).MaxRows, str(rsEstimGr!Porc)
                    End If
                Case "C"
                    sprGr(5).GetText 1, sprGr(5).MaxRows, valor
                    If valor = rub Then
                        sprGr(5).SetText 3, sprGr(5).MaxRows, str(rsEstimGr!Porc)
                    End If
            End Select
        Case "II"
            Select Case rsEstimGr!Grupo
                Case "A"
                    sprGr(6).GetText 1, sprGr(6).MaxRows, valor
                    If valor = rub Then
                        sprGr(6).SetText 3, sprGr(6).MaxRows, str(rsEstimGr!Porc)
                    End If
                Case "B"
                    sprGr(7).GetText 1, sprGr(7).MaxRows, valor
                    If valor = rub Then
                        sprGr(7).SetText 3, sprGr(7).MaxRows, str(rsEstimGr!Porc)
                    End If
                Case "C"
                    sprGr(8).GetText 1, sprGr(8).MaxRows, valor
                    If valor = rub Then
                        sprGr(8).SetText 3, sprGr(8).MaxRows, str(rsEstimGr!Porc)
                    End If
            End Select
    
    End Select
    rsEstimGr.MoveNext
    Loop
Loop

End Sub

Private Sub L_LlenarGrillasRub()
Dim rub As String, valor As Variant, tot As Single
Dim rsGr As Recordset
Dim sql As String
Dim i

'No hago esta consulta junto con Dia porque hay que separar los
'casos de aquellos modelos que no tienen datos por grupo
i = 1
Do While Not rsRealRub.EOF
    rub = rsRealRub!cod_rubr
    Do While rub = rsRealRub!cod_rubr
    Select Case rsRealRub.cod_depn
        
        Case "EZE"
            Select Case rsRealRub!Grupo
                Case "A"
                    sprGr(0).GetText 1, i, valor
                    If valor = rub Then
                        sprGr(0).SetText 2, i, str(rsRealRub!IMP)
                    End If
                Case "B"
                    sprGr(1).GetText 1, i, valor
                    If valor = rub Then
                        sprGr(1).SetText 2, i, str(rsRealRub!IMP)
                    End If
                Case "C"
                    sprGr(2).GetText 1, i, valor
                    If valor = rub Then
                        sprGr(2).SetText 2, i, str(rsRealRub!IMP)
                    End If
            End Select
        Case "AEP"
            Select Case rsRealRub!Grupo
                Case "A"
                    sprGr(3).GetText 1, i, valor
                    If valor = rub Then
                        sprGr(3).SetText 2, i, str(rsRealRub!IMP)
                    End If
                Case "B"
                    sprGr(4).GetText 1, i, valor
                    If valor = rub Then
                        sprGr(4).SetText 2, i, str(rsRealRub!IMP)
                    End If
                Case "C"
                    sprGr(5).GetText 1, i, valor
                    If valor = rub Then
                        sprGr(5).SetText 2, i, str(rsRealRub!IMP)
                    End If
            End Select
        Case "II"
            Select Case rsRealRub!Grupo
                Case "A"
                    sprGr(6).GetText 1, i, valor
                    If valor = rub Then
                        sprGr(6).SetText 2, i, str(rsRealRub!IMP)
                    End If
                Case "B"
                    sprGr(7).GetText 1, i, valor
                    If valor = rub Then
                        sprGr(7).SetText 2, i, str(rsRealRub!IMP)
                    End If
                Case "C"
                    sprGr(8).GetText 1, i, valor
                    If valor = rub Then
                        sprGr(8).SetText 2, i, str(rsRealRub!IMP)
                    End If
            End Select
        
    End Select
    rsRealRub.MoveNext
    If rsRealRub.EOF Then
        Exit Do
    End If
    Loop
    i = i + 1

Loop
L_SumatoriaGrillas 0, SprTot(0)
L_SumatoriaGrillas 3, SprTot(1)
L_SumatoriaGrillas 6, SprTot(2)
End Sub

Private Function L_ArmarcondicionEstim()
Dim Cond, i
Dim auxCond

Cond = " WHERE anio = " & Year(mskFDesde.FormattedText)
Cond = Cond & " And Mes = " & Month(mskFDesde.FormattedText)
Cond = Cond & " And Tipo_porc = 'I' "
Cond = Cond & " And Nivel = 0 "

L_ArmarcondicionEstim = Cond

End Function

