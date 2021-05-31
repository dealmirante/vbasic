VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmResumenDiario 
   Caption         =   "Estimados vs Ventas-Tickets-Pasajeros"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   11070
   Begin VB.CheckBox chkNeto 
      Caption         =   "Venta sin Duty Check"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4485
      TabIndex        =   79
      Top             =   225
      Visible         =   0   'False
      Width           =   2085
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
      Height          =   525
      Index           =   2
      Left            =   10140
      Picture         =   "frmResumenDiario.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   150
      Width           =   600
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
      Height          =   525
      Index           =   1
      Left            =   8865
      Picture         =   "frmResumenDiario.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   150
      Width           =   600
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
      Height          =   525
      Index           =   0
      Left            =   8235
      Picture         =   "frmResumenDiario.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   150
      Width           =   615
   End
   Begin VB.CommandButton botExcel 
      Caption         =   "Excel"
      Height          =   525
      Left            =   9510
      Picture         =   "frmResumenDiario.frx":0A26
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   150
      Width           =   600
   End
   Begin VB.Frame frCab 
      Caption         =   "Datos de Cabecera"
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
      Height          =   780
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   10920
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskMes 
         Height          =   300
         Left            =   2985
         TabIndex        =   2
         Top             =   270
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Año"
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
         Left            =   135
         TabIndex        =   4
         Top             =   315
         Width           =   780
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mes"
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
         Left            =   2160
         TabIndex        =   3
         Top             =   270
         Width           =   780
      End
   End
   Begin TabDlg.SSTab tabEst 
      Height          =   5520
      Left            =   45
      TabIndex        =   9
      Top             =   870
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   9737
      _Version        =   327680
      Tabs            =   7
      TabsPerRow      =   8
      TabHeight       =   441
      ForeColor       =   255
      TabCaption(0)   =   "Total Bs. As."
      TabPicture(0)   =   "frmResumenDiario.frx":0FB8
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tabTotal"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "INTER A Lleg"
      TabPicture(1)   =   "frmResumenDiario.frx":0FD4
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabGA"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "INTER A Sal"
      TabPicture(2)   =   "frmResumenDiario.frx":0FF0
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabGB"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "INTER B"
      TabPicture(3)   =   "frmResumenDiario.frx":100C
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tabGC"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "AEP"
      TabPicture(4)   =   "frmResumenDiario.frx":1028
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TabCIA"
      Tab(4).Control(0).Enabled=   0   'False
      TabCaption(5)   =   "Total CIA"
      TabPicture(5)   =   "frmResumenDiario.frx":1044
      Tab(5).ControlCount=   1
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "TabGD"
      Tab(5).Control(0).Enabled=   0   'False
      TabCaption(6)   =   "INTERIOR"
      TabPicture(6)   =   "frmResumenDiario.frx":1060
      Tab(6).ControlCount=   1
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "SSTab1"
      Tab(6).Control(0).Enabled=   0   'False
      Begin TabDlg.SSTab tabTotal 
         Height          =   5025
         Left            =   135
         TabIndex        =   10
         Top             =   315
         Width           =   10710
         _ExtentX        =   18891
         _ExtentY        =   8864
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   6
         Tab             =   1
         TabsPerRow      =   6
         TabHeight       =   882
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "$ Real vs. Estimado"
         TabPicture(0)   =   "frmResumenDiario.frx":107C
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprTot(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "$ Actual vs. Anterior"
         TabPicture(1)   =   "frmResumenDiario.frx":1098
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "sprTot(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pax Transitado Real vs Estim."
         TabPicture(2)   =   "frmResumenDiario.frx":10B4
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprTot(4)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Pr. Pax Trans. Real vs Estim."
         TabPicture(3)   =   "frmResumenDiario.frx":10D0
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprTot(5)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Pax Transitado Actual vs. Ant"
         TabPicture(4)   =   "frmResumenDiario.frx":10EC
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprTot(2)"
         Tab(4).Control(0).Enabled=   0   'False
         TabCaption(5)   =   "Pr. Pax Trans Actual vs Ant"
         TabPicture(5)   =   "frmResumenDiario.frx":1108
         Tab(5).ControlCount=   1
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "sprTot(3)"
         Tab(5).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   0
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":1124
            TabIndex        =   34
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   5
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":19B4
            TabIndex        =   33
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4245
            Index           =   4
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":219C
            TabIndex        =   32
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4230
            Index           =   3
            Left            =   -74910
            OleObjectBlob   =   "frmResumenDiario.frx":294C
            TabIndex        =   19
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   2
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":3127
            TabIndex        =   18
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   1
            Left            =   105
            OleObjectBlob   =   "frmResumenDiario.frx":3902
            TabIndex        =   17
            Top             =   150
            Width           =   10470
         End
      End
      Begin TabDlg.SSTab tabGA 
         Height          =   5040
         Left            =   -74895
         TabIndex        =   11
         Top             =   315
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   8890
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   6
         Tab             =   1
         TabsPerRow      =   6
         TabHeight       =   882
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "$ Real vs. Estimado"
         TabPicture(0)   =   "frmResumenDiario.frx":40E5
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprA(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "$ Actual vs. Anterior"
         TabPicture(1)   =   "frmResumenDiario.frx":4101
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "sprA(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pax Transitado Real vs Estim"
         TabPicture(2)   =   "frmResumenDiario.frx":411D
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprA(4)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Pr Pax Trans Real vs Estim"
         TabPicture(3)   =   "frmResumenDiario.frx":4139
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprA(5)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Pax Transitado Actual vs Ant"
         TabPicture(4)   =   "frmResumenDiario.frx":4155
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprA(2)"
         Tab(4).Control(0).Enabled=   0   'False
         TabCaption(5)   =   "Pr Pax Trans Actual vs Ant"
         TabPicture(5)   =   "frmResumenDiario.frx":4171
         Tab(5).ControlCount=   1
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "sprA(3)"
         Tab(5).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprA 
            Height          =   4260
            Index           =   0
            Left            =   -74865
            OleObjectBlob   =   "frmResumenDiario.frx":418D
            TabIndex        =   35
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprA 
            Height          =   4260
            Index           =   5
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":4A1D
            TabIndex        =   31
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprA 
            Height          =   4260
            Index           =   4
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":5205
            TabIndex        =   30
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprA 
            Height          =   4260
            Index           =   1
            Left            =   135
            OleObjectBlob   =   "frmResumenDiario.frx":59B5
            TabIndex        =   14
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprA 
            Height          =   4230
            Index           =   2
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":6198
            TabIndex        =   20
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprA 
            Height          =   4230
            Index           =   3
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":6973
            TabIndex        =   21
            Top             =   135
            Width           =   10470
         End
      End
      Begin TabDlg.SSTab tabGB 
         Height          =   5055
         Left            =   -74895
         TabIndex        =   12
         Top             =   315
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8916
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   6
         Tab             =   1
         TabsPerRow      =   6
         TabHeight       =   882
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "$ Real vs. Estimado"
         TabPicture(0)   =   "frmResumenDiario.frx":714E
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprAS(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "$ Actual vs. Anterior"
         TabPicture(1)   =   "frmResumenDiario.frx":716A
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "sprAS(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pax Transitado Real vs Estim"
         TabPicture(2)   =   "frmResumenDiario.frx":7186
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprAS(4)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Pr Pax Trans Real vs Estim"
         TabPicture(3)   =   "frmResumenDiario.frx":71A2
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprAS(5)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Pax Transitado Actual vs Ant"
         TabPicture(4)   =   "frmResumenDiario.frx":71BE
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprAS(2)"
         Tab(4).Control(0).Enabled=   0   'False
         TabCaption(5)   =   "Pr Pax Trans Actual vs Ant"
         TabPicture(5)   =   "frmResumenDiario.frx":71DA
         Tab(5).ControlCount=   1
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "sprAS(3)"
         Tab(5).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprAS 
            Height          =   4260
            Index           =   0
            Left            =   -74835
            OleObjectBlob   =   "frmResumenDiario.frx":71F6
            TabIndex        =   36
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprAS 
            Height          =   4260
            Index           =   5
            Left            =   -74865
            OleObjectBlob   =   "frmResumenDiario.frx":7A86
            TabIndex        =   29
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprAS 
            Height          =   4260
            Index           =   4
            Left            =   -74850
            OleObjectBlob   =   "frmResumenDiario.frx":826E
            TabIndex        =   28
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprAS 
            Height          =   4260
            Index           =   1
            Left            =   150
            OleObjectBlob   =   "frmResumenDiario.frx":8A1E
            TabIndex        =   15
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprAS 
            Height          =   4230
            Index           =   2
            Left            =   -74865
            OleObjectBlob   =   "frmResumenDiario.frx":9201
            TabIndex        =   22
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprAS 
            Height          =   4230
            Index           =   3
            Left            =   -74865
            OleObjectBlob   =   "frmResumenDiario.frx":99A4
            TabIndex        =   23
            Top             =   135
            Width           =   10470
         End
      End
      Begin TabDlg.SSTab tabGC 
         Height          =   5040
         Left            =   -74895
         TabIndex        =   13
         Top             =   330
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   8890
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   6
         Tab             =   1
         TabsPerRow      =   6
         TabHeight       =   882
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "$ Real vs. Estimado"
         TabPicture(0)   =   "frmResumenDiario.frx":A17F
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprB(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "$ Actual vs. Anterior"
         TabPicture(1)   =   "frmResumenDiario.frx":A19B
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "sprB(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pax Transitado Real vs. Estim"
         TabPicture(2)   =   "frmResumenDiario.frx":A1B7
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprB(4)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Pr. Pax Trans Real vs Estim"
         TabPicture(3)   =   "frmResumenDiario.frx":A1D3
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprB(5)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Pax Transitado Actual vs Ant"
         TabPicture(4)   =   "frmResumenDiario.frx":A1EF
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprB(2)"
         Tab(4).Control(0).Enabled=   0   'False
         TabCaption(5)   =   "Pr. Pax Trans Actual vs Ant"
         TabPicture(5)   =   "frmResumenDiario.frx":A20B
         Tab(5).ControlCount=   1
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "sprB(3)"
         Tab(5).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprB 
            Height          =   4260
            Index           =   0
            Left            =   -74865
            OleObjectBlob   =   "frmResumenDiario.frx":A227
            TabIndex        =   37
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprB 
            Height          =   4260
            Index           =   5
            Left            =   -74865
            OleObjectBlob   =   "frmResumenDiario.frx":AAB7
            TabIndex        =   27
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprB 
            Height          =   4260
            Index           =   4
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":B29F
            TabIndex        =   26
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprB 
            Height          =   4260
            Index           =   1
            Left            =   120
            OleObjectBlob   =   "frmResumenDiario.frx":BA4F
            TabIndex        =   16
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprB 
            Height          =   4230
            Index           =   2
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":C232
            TabIndex        =   24
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprB 
            Height          =   4230
            Index           =   3
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":C9D5
            TabIndex        =   25
            Top             =   135
            Width           =   10470
         End
      End
      Begin TabDlg.SSTab TabCIA 
         Height          =   5025
         Left            =   -74850
         TabIndex        =   38
         Top             =   360
         Width           =   10710
         _ExtentX        =   18891
         _ExtentY        =   8864
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   6
         Tab             =   1
         TabsPerRow      =   6
         TabHeight       =   882
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "$ Real vs. Estimado"
         TabPicture(0)   =   "frmResumenDiario.frx":D1B0
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprC(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "$ Actual vs. Anterior"
         TabPicture(1)   =   "frmResumenDiario.frx":D1CC
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "sprC(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pax Transitado Real vs Estim."
         TabPicture(2)   =   "frmResumenDiario.frx":D1E8
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprC(4)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Pr. Pax Trans. Real vs Estim."
         TabPicture(3)   =   "frmResumenDiario.frx":D204
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprC(5)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Pax Transitado Actual vs. Ant"
         TabPicture(4)   =   "frmResumenDiario.frx":D220
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprC(2)"
         Tab(4).Control(0).Enabled=   0   'False
         TabCaption(5)   =   "Pr. Pax Trans Actual vs Ant"
         TabPicture(5)   =   "frmResumenDiario.frx":D23C
         Tab(5).ControlCount=   1
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "sprC(3)"
         Tab(5).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprC 
            Height          =   4230
            Index           =   3
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":D258
            TabIndex        =   56
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprC 
            Height          =   4230
            Index           =   2
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":DA33
            TabIndex        =   55
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprC 
            Height          =   4260
            Index           =   5
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":E1D6
            TabIndex        =   54
            Top             =   120
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprC 
            Height          =   4260
            Index           =   4
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":E9BE
            TabIndex        =   53
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprC 
            Height          =   4260
            Index           =   1
            Left            =   135
            OleObjectBlob   =   "frmResumenDiario.frx":F16E
            TabIndex        =   52
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   11
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":F951
            TabIndex        =   44
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   10
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":10134
            TabIndex        =   43
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4230
            Index           =   9
            Left            =   -74910
            OleObjectBlob   =   "frmResumenDiario.frx":1090F
            TabIndex        =   42
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4245
            Index           =   8
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":110EA
            TabIndex        =   41
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   7
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":1189A
            TabIndex        =   40
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprC 
            Height          =   4260
            Index           =   0
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":12082
            TabIndex        =   39
            Top             =   120
            Width           =   10470
         End
      End
      Begin TabDlg.SSTab TabGD 
         Height          =   5025
         Left            =   -74865
         TabIndex        =   45
         Top             =   375
         Width           =   10710
         _ExtentX        =   18891
         _ExtentY        =   8864
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   6
         Tab             =   1
         TabsPerRow      =   6
         TabHeight       =   882
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "$ Real vs. Estimado"
         TabPicture(0)   =   "frmResumenDiario.frx":12912
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprCIA(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "$ Actual vs. Anterior"
         TabPicture(1)   =   "frmResumenDiario.frx":1292E
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "sprCIA(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pax Transitado Real vs Estim."
         TabPicture(2)   =   "frmResumenDiario.frx":1294A
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprCIA(4)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Pr. Pax Trans. Real vs Estim."
         TabPicture(3)   =   "frmResumenDiario.frx":12966
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprCIA(5)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Pax Transitado Actual vs. Ant"
         TabPicture(4)   =   "frmResumenDiario.frx":12982
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprCIA(2)"
         Tab(4).Control(0).Enabled=   0   'False
         TabCaption(5)   =   "Pr. Pax Trans Actual vs Ant"
         TabPicture(5)   =   "frmResumenDiario.frx":1299E
         Tab(5).ControlCount=   1
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "sprCIA(3)"
         Tab(5).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprCIA 
            Height          =   4230
            Index           =   3
            Left            =   -74865
            OleObjectBlob   =   "frmResumenDiario.frx":129BA
            TabIndex        =   61
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprCIA 
            Height          =   4230
            Index           =   2
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":13195
            TabIndex        =   60
            Top             =   120
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprCIA 
            Height          =   4260
            Index           =   5
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":13938
            TabIndex        =   59
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprCIA 
            Height          =   4260
            Index           =   4
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":14120
            TabIndex        =   58
            Top             =   120
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprCIA 
            Height          =   4260
            HelpContextID   =   1
            Index           =   1
            Left            =   120
            OleObjectBlob   =   "frmResumenDiario.frx":148D0
            TabIndex        =   57
            Top             =   90
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   17
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":150B3
            TabIndex        =   51
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   16
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":15896
            TabIndex        =   50
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4230
            Index           =   15
            Left            =   -74910
            OleObjectBlob   =   "frmResumenDiario.frx":16071
            TabIndex        =   49
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4245
            Index           =   14
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":1684C
            TabIndex        =   48
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   13
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":16FFC
            TabIndex        =   47
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprCIA 
            Height          =   4260
            Index           =   0
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":177E4
            TabIndex        =   46
            Top             =   150
            Width           =   10470
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5025
         Left            =   -74835
         TabIndex        =   62
         Top             =   390
         Width           =   10710
         _ExtentX        =   18891
         _ExtentY        =   8864
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   6
         Tab             =   1
         TabsPerRow      =   6
         TabHeight       =   882
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "$ Real vs. Estimado"
         TabPicture(0)   =   "frmResumenDiario.frx":18074
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprD(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "$ Actual vs. Anterior"
         TabPicture(1)   =   "frmResumenDiario.frx":18090
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "sprD(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "Pax Transitado Real vs Estim."
         TabPicture(2)   =   "frmResumenDiario.frx":180AC
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprD(4)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "Pr. Pax Trans. Real vs Estim."
         TabPicture(3)   =   "frmResumenDiario.frx":180C8
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprD(5)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "Pax Transitado Actual vs. Ant"
         TabPicture(4)   =   "frmResumenDiario.frx":180E4
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "sprD(2)"
         Tab(4).Control(0).Enabled=   0   'False
         TabCaption(5)   =   "Pr. Pax Trans Actual vs Ant"
         TabPicture(5)   =   "frmResumenDiario.frx":18100
         Tab(5).ControlCount=   1
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "sprD(3)"
         Tab(5).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprCIA 
            Height          =   4260
            Index           =   6
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":1811C
            TabIndex        =   63
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   6
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":189AC
            TabIndex        =   64
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4245
            Index           =   12
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":19194
            TabIndex        =   65
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4230
            Index           =   18
            Left            =   -74910
            OleObjectBlob   =   "frmResumenDiario.frx":19944
            TabIndex        =   66
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   19
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":1A11F
            TabIndex        =   67
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   4260
            Index           =   20
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":1A8FA
            TabIndex        =   68
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprD 
            Height          =   4260
            Index           =   1
            Left            =   135
            OleObjectBlob   =   "frmResumenDiario.frx":1B0DD
            TabIndex        =   69
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprCIA 
            Height          =   4260
            Index           =   8
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":1B8C0
            TabIndex        =   70
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprCIA 
            Height          =   4260
            Index           =   9
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":1C070
            TabIndex        =   71
            Top             =   120
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprCIA 
            Height          =   4230
            Index           =   10
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":1C858
            TabIndex        =   72
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprCIA 
            Height          =   4230
            Index           =   11
            Left            =   -74895
            OleObjectBlob   =   "frmResumenDiario.frx":1CFFB
            TabIndex        =   73
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprD 
            Height          =   4260
            Index           =   0
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":1D7D6
            TabIndex        =   74
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprD 
            Height          =   4260
            Index           =   4
            Left            =   -74865
            OleObjectBlob   =   "frmResumenDiario.frx":1E066
            TabIndex        =   75
            Top             =   150
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprD 
            Height          =   4260
            Index           =   5
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":1E816
            TabIndex        =   76
            Top             =   120
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprD 
            Height          =   4230
            Index           =   2
            Left            =   -74865
            OleObjectBlob   =   "frmResumenDiario.frx":1EFFE
            TabIndex        =   77
            Top             =   135
            Width           =   10470
         End
         Begin FPSpread.vaSpread sprD 
            Height          =   4230
            Index           =   3
            Left            =   -74880
            OleObjectBlob   =   "frmResumenDiario.frx":1F7A1
            TabIndex        =   78
            Top             =   135
            Width           =   10470
         End
      End
   End
End
Attribute VB_Name = "frmResumenDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsData As Recordset
Dim RsActual As Recordset
Dim RsAnterior As Recordset

Dim siImp As Boolean, siTic As Boolean, siPax As Boolean

Dim Control_Int As Boolean
Dim Control_BsAs As Boolean

Dim Fch_Int As String
Dim Fch_BsAs As String
Private Sub L_CorregirEstimado(spr As Object)
Dim valorPor As Variant
Dim valorReal As Variant, valorAcum As Variant
Dim tope As Integer, fila As Integer
Dim Porc As Single

If mskMes.Text = Month(Now) And mskAnio.Text = Year(Now) Then
   If Day(Now) = 1 Then
    tope = Day(Now)
   Else
    tope = Day(Now) - 1
   End If
Else
    tope = Day(Func.func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text))
End If

spr.GetText 9, tope, valorPor

Porc = Format$(valorPor / 100, "#.000")

For fila = tope + 1 To spr.MaxRows
  spr.GetText 3, fila, valorReal
  spr.GetText 10, fila - 1, valorAcum
  
  spr.SetText 10, fila, valorAcum + (valorReal + (Porc * valorReal))
  
Next

End Sub

Private Sub L_Grilla(spr As control, valor As Single)
    spr.MaxRows = spr.MaxRows + 1
    spr.SetText 1, spr.MaxRows, Trim(RsData!DM)
    spr.SetText 2, spr.MaxRows, str(RsData!Act)
    spr.SetText 3, spr.MaxRows, str(valor)

    Spread_TotalesGrillaAcum spr, 2, 6, spr.MaxRows
    Spread_TotalesGrillaAcum spr, 3, 7, spr.MaxRows
    
    spread_ResaltarCelda spr, 4, spr.MaxRows
    spread_ResaltarCelda spr, 5, spr.MaxRows
    spread_ResaltarCelda spr, 8, spr.MaxRows
    spread_ResaltarCelda spr, 9, spr.MaxRows

End Sub
Private Sub L_LlenarGrillas()
Dim SumImp As Double
Dim Tipo As String
Dim Estimado As Single
Dim Anterior_A As Single
Dim Anterior_AS As Single
Dim Anterior_B As Single
Dim Anterior_AEP As Single
Dim Anterior_TOT As Single
Dim Anterior_CIA As Single
Dim Anterior_INT As Single

Dim i As Integer

On Error GoTo ErrGLC:

Do While Not RsData.EOF
    Tipo = RsData!Tipo
    Select Case Tipo
        Case "INTAL"
            Do While Tipo = RsData!Tipo
             If mskMes.Text = Mid(RsData!DM, 4) Then
                L_Grilla sprA(0), RsData!etim
                L_Grilla sprA(1), RsData!ant
               Else
                Anterior_A = RsData!ant
             End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If
                            
            Loop
        Case "INTAS"
            Do While Tipo = RsData!Tipo
             If mskMes.Text = Mid(RsData!DM, 4) Then
                L_Grilla sprAS(0), RsData!etim
                L_Grilla sprAS(1), RsData!ant
               Else
                Anterior_AS = RsData!ant
             End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If
                            
            Loop
        Case "INTB"
            Do While Tipo = RsData!Tipo
             If mskMes.Text = Mid(RsData!DM, 4) Then
                L_Grilla sprB(0), RsData!etim
                L_Grilla sprB(1), RsData!ant
               Else
                Anterior_B = RsData!ant
             End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If
                            
            Loop
        Case "AEP"
            Do While Tipo = RsData!Tipo
             If mskMes.Text = Mid(RsData!DM, 4) Then
                L_Grilla sprC(0), RsData!etim
                L_Grilla sprC(1), RsData!ant
               Else
                Anterior_AEP = RsData!ant
             End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If
                            
            Loop
        Case "TOT"
            Do While Tipo = RsData!Tipo
             If mskMes.Text = Mid(RsData!DM, 4) Then
                L_Grilla SprTot(0), RsData!etim
                L_Grilla SprTot(1), RsData!ant
               Else
                Anterior_TOT = RsData!ant
             End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If
                            
            Loop
            
        Case "CIA"
            Do While Tipo = RsData!Tipo
             If mskMes.Text = Mid(RsData!DM, 4) Then
                L_Grilla sprCIA(0), RsData!etim
                L_Grilla sprCIA(1), RsData!ant
               Else
                Anterior_CIA = RsData!ant
             End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If
                            
            Loop
            
        Case "INT"
            Do While Tipo = RsData!Tipo
             If mskMes.Text = Mid(RsData!DM, 4) Then
                L_Grilla sprD(0), RsData!etim
                L_Grilla sprD(1), RsData!ant
               Else
                Anterior_INT = RsData!ant
             End If
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

If mskMes.Text = 2 Then
  sprA(1).SetText 3, sprA(1).MaxRows, str(Anterior_A)
  sprAS(1).SetText 3, sprA(1).MaxRows, str(Anterior_AS)
  sprB(1).SetText 3, sprB(1).MaxRows, str(Anterior_B)
  sprC(1).SetText 3, sprC(1).MaxRows, str(Anterior_AEP)
  SprTot(1).SetText 3, SprTot(1).MaxRows, str(Anterior_TOT)
  sprCIA(1).SetText 3, sprCIA(1).MaxRows, str(Anterior_CIA)
  sprD(1).SetText 3, sprD(1).MaxRows, str(Anterior_INT)
End If

For i = 0 To 1
    If Val(mskMes.Text) = Month(Now) And Val(mskAnio.Text) = Year(Now) Then
    Spread.Spead_LimpiarGrilla sprA(i), Day(Now), sprA(i).MaxRows, 4, 6
    Spread.Spead_LimpiarGrilla sprA(i), Day(Now), sprA(i).MaxRows, 8, 9
    Spread.Spead_LimpiarGrilla sprAS(i), Day(Now), sprA(i).MaxRows, 4, 6
    Spread.Spead_LimpiarGrilla sprAS(i), Day(Now), sprA(i).MaxRows, 8, 9
    Spread.Spead_LimpiarGrilla sprB(i), Day(Now), sprB(i).MaxRows, 4, 6
    Spread.Spead_LimpiarGrilla sprB(i), Day(Now), sprB(i).MaxRows, 8, 9
    Spread.Spead_LimpiarGrilla sprC(i), Day(Now), sprC(i).MaxRows, 4, 6
    Spread.Spead_LimpiarGrilla sprC(i), Day(Now), sprC(i).MaxRows, 8, 9
    Spread.Spead_LimpiarGrilla sprD(i), Day(Now), sprD(i).MaxRows, 4, 6
    Spread.Spead_LimpiarGrilla sprD(i), Day(Now), sprD(i).MaxRows, 8, 9
    Spread.Spead_LimpiarGrilla SprTot(i), Day(Now), SprTot(i).MaxRows, 4, 6
    Spread.Spead_LimpiarGrilla SprTot(i), Day(Now), SprTot(i).MaxRows, 8, 9
    Spread.Spead_LimpiarGrilla sprCIA(i), Day(Now), sprCIA(i).MaxRows, 4, 6
    Spread.Spead_LimpiarGrilla sprCIA(i), Day(Now), sprCIA(i).MaxRows, 8, 9
    End If
Next

ErrGLC:
     Exit Sub
End Sub
Private Sub AjustarVenta(Tipo As Integer, rs As Recordset, col As Integer)
Dim fch As Variant
Dim fila As Integer

On Error GoTo err1:

rs.MoveFirst
Do While Not rs.EOF
    Select Case rs!cod_sdep
        Case "AEP"
            fila = 1
            Do While rs!cod_sdep = "AEP"
               'sprC(1).GetText 1, FILA, FCH
               'If RS!DM = FCH Then
                If Tipo = 0 Then
                  sprC(1).SetText col, fila, str(rs!impvta)
                  If col = 2 Then
                  sprC(0).SetText col, fila, str(rs!impvta)
                  End If
                Else
                  sprC(1).SetText col, fila, str(rs!impvta - rs!impdc)
                  If col = 2 Then
                  sprC(0).SetText col, fila, str(rs!impvta - rs!impdc)
                  End If
                End If
               'Else
               'End If
               fila = fila + 1
               rs.MoveNext
               If rs.EOF Then
                  Exit Do
               End If
            Loop
        Case "INTAL"
            fila = 1
            Do While rs!cod_sdep = "INTAL"
                If Tipo = 0 Then
                  sprA(1).SetText col, fila, str(rs!impvta)
                  If col = 2 Then
                  sprA(0).SetText col, fila, str(rs!impvta)
                  End If
                Else
                  sprA(1).SetText col, fila, str(rs!impvta - rs!impdc)
                  If col = 2 Then
                  sprA(0).SetText col, fila, str(rs!impvta - rs!impdc)
                  End If
                End If
               fila = fila + 1
               rs.MoveNext
               If rs.EOF Then
                  Exit Do
               End If
            Loop
        Case "INTAS"
            fila = 1
            Do While rs!cod_sdep = "INTAS"
                If Tipo = 0 Then
                  sprAS(1).SetText col, fila, str(rs!impvta)
                  If col = 2 Then
                  sprAS(0).SetText col, fila, str(rs!impvta)
                  End If
                Else
                  sprAS(1).SetText col, fila, str(rs!impvta - rs!impdc)
                  If col = 2 Then
                  sprAS(0).SetText col, fila, str(rs!impvta - rs!impdc)
                  End If
                End If
               fila = fila + 1
               rs.MoveNext
               If rs.EOF Then
                  Exit Do
               End If
            Loop
        Case "INTB"
            fila = 1
            Do While rs!cod_sdep = "INTB"
                If Tipo = 0 Then
                  sprB(1).SetText col, fila, str(rs!impvta)
                  If col = 2 Then
                  sprB(0).SetText col, fila, str(rs!impvta)
                  End If
                Else
                  sprB(1).SetText col, fila, str(rs!impvta - rs!impdc)
                  If col = 2 Then
                  sprB(0).SetText col, fila, str(rs!impvta - rs!impdc)
                  End If
                End If
               fila = fila + 1
               rs.MoveNext
               If rs.EOF Then
                  Exit Do
               End If
            Loop
        Case "INT"
            fila = 1
            Do While rs!cod_sdep = "INT"
                If Tipo = 0 Then
                  sprD(1).SetText col, fila, str(rs!impvta)
                  If col = 2 Then
                  sprD(0).SetText col, fila, str(rs!impvta)
                  End If
                Else
                  sprD(1).SetText col, fila, str(rs!impvta - rs!impdc)
                  If col = 2 Then
                  sprD(0).SetText col, fila, str(rs!impvta - rs!impdc)
                  End If
                End If
               fila = fila + 1
               rs.MoveNext
               If rs.EOF Then
                  Exit Do
               End If
            Loop
        Case "TOT"
            fila = 1
            Do While rs!cod_sdep = "TOT"
                If Tipo = 0 Then
                  SprTot(1).SetText col, fila, str(rs!impvta)
                  If col = 2 Then
                  SprTot(0).SetText col, fila, str(rs!impvta)
                  End If
                Else
                  SprTot(1).SetText col, fila, str(rs!impvta - rs!impdc)
                  If col = 2 Then
                  SprTot(0).SetText col, fila, str(rs!impvta - rs!impdc)
                  End If
                End If
               fila = fila + 1
               rs.MoveNext
               If rs.EOF Then
                  Exit Do
               End If
            Loop
        Case "CIA"
            fila = 1
            Do While rs!cod_sdep = "CIA"
                If Tipo = 0 Then
                  sprCIA(1).SetText col, fila, str(rs!impvta)
                  If col = 2 Then
                  sprCIA(0).SetText col, fila, str(rs!impvta)
                  End If
                Else
                  sprCIA(1).SetText col, fila, str(rs!impvta - rs!impdc)
                  If col = 2 Then
                  sprCIA(0).SetText col, fila, str(rs!impvta - rs!impdc)
                  End If
                End If
               fila = fila + 1
               rs.MoveNext
               If rs.EOF Then
                  Exit Do
               End If
            Loop
    End Select
Loop

err1:

End Sub

Private Sub L_LlenarGrillasPax()
Dim SumImp As Double
Dim Tipo As String
Dim Anterior_A As Single
Dim Anterior_AS As Single
Dim Anterior_B As Single
Dim Anterior_AEP As Single
Dim Anterior_TOT As Single
Dim Anterior_CIA As Single
Dim Anterior_INT As Single


On Error GoTo ErrGLC:

Do While Not RsData.EOF
    Tipo = RsData!Tipo
    Select Case Tipo
        Case "INTAL"
            Do While Tipo = RsData!Tipo
                'If CDate((RsData!DM) & "-" & mskAnio.Text) < (Format$(Now, "dd-mm"))
                If CDate((RsData!DM) & "-" & mskAnio.Text) <= (Format$(Fch_BsAs, "dd-mm-yyyy")) _
                Then
                ' Or mskMes.Text < Month(Now)
                
                   If mskMes.Text = Mid(RsData!DM, 4) Then
                    L_Grilla sprA(2), RsData!ant
                    L_Grilla sprA(4), RsData!estim
                   Else
                    Anterior_A = RsData!ant
                   End If
                End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If

            Loop
        Case "INTAS"
            Do While Tipo = RsData!Tipo
                'If CDate((RsData!DM) & "-" & mskAnio.Text) < (Format$(Now, "dd-mm"))
                If CDate((RsData!DM) & "-" & mskAnio.Text) <= (Format$(Fch_BsAs, "dd-mm-yyyy")) _
                Then
                'Or mskMes.Text < Month(Now)
                   If mskMes.Text = Mid(RsData!DM, 4) Then
                    L_Grilla sprAS(2), RsData!ant
                    L_Grilla sprAS(4), RsData!estim
                   Else
                    Anterior_AS = RsData!ant
                   End If
                End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If

            Loop
        Case "INTB"
            Do While Tipo = RsData!Tipo
                'If CDate((RsData!DM) & "-" & mskAnio.Text) < (Format$(Now, "dd-mm"))
                If CDate((RsData!DM) & "-" & mskAnio.Text) <= (Format$(Fch_BsAs, "dd-mm-yyyy")) _
                Then
                'Or mskMes.Text < Month(Now)
                  If mskMes.Text = Mid(RsData!DM, 4) Then
                    L_Grilla sprB(2), RsData!ant
                    L_Grilla sprB(4), RsData!estim
                  Else
                    Anterior_B = RsData!ant
                  End If
                End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If

            Loop
        Case "AEP"
            Do While Tipo = RsData!Tipo
                'If CDate((RsData!DM) & "-" & mskAnio.Text) < (Format$(Now, "dd-mm"))
                If CDate((RsData!DM) & "-" & mskAnio.Text) <= (Format$(Fch_BsAs, "dd-mm-yyyy")) _
                Then
                'Or mskMes.Text < Month(Now)
                  If mskMes.Text = Mid(RsData!DM, 4) Then
                    L_Grilla sprC(2), RsData!ant
                    L_Grilla sprC(4), RsData!estim
                  Else
                    Anterior_AEP = RsData!ant
                  End If
                End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If

            Loop
            
        Case "INT"
            Do While Tipo = RsData!Tipo
                'If CDate((RsData!DM) & "-" & mskAnio.Text) < (Format$(Now, "dd-mm"))
                If CDate((RsData!DM) & "-" & mskAnio.Text) <= (Format$(Fch_Int, "dd-mm-yyyy")) _
                Then
                'Or mskMes.Text < Month(Now)
                  If mskMes.Text = Mid(RsData!DM, 4) Then
                    L_Grilla sprD(2), RsData!ant
                    L_Grilla sprD(4), RsData!estim
                  Else
                    Anterior_INT = RsData!ant
                  End If
                End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If

            Loop
            
        Case "TOT"
            Do While Tipo = RsData!Tipo
                'If CDate((RsData!DM) & "-" & mskAnio.Text) < (Format$(Now - 1, "dd-mm"))
                If (CDate((RsData!DM) & "-" & mskAnio.Text) <= (Format$(Fch_BsAs, "dd-mm-yyyy"))) _
                 Then
                 'Or mskMes.Text < Month(Now)
                'Or (CDate((RsData!DM) & "-" & mskAnio.Text) = (Format$(Now - 1, "dd-mm")) And Control_BsAs)
                  If mskMes.Text = Mid(RsData!DM, 4) Then
                    L_Grilla SprTot(2), RsData!ant
                    L_Grilla SprTot(4), RsData!estim
                  Else
                    Anterior_TOT = RsData!ant
                  End If
                End If
                RsData.MoveNext
                If RsData.EOF Then
                    Exit Do
                End If

            Loop
            
        Case "CIA"
            Do While Tipo = RsData!Tipo
                'If CDate((RsData!DM) & "-" & mskAnio.Text) < (Format$(Now - 1, "dd-mm")) _
                'Or mskMes.Text < Month(Now) _
                'Or (CDate((RsData!DM) & "-" & mskAnio.Text) = (Format$(Now - 1, "dd-mm")) And Control_Int) Then
                
                If (CDate((RsData!DM) & "-" & mskAnio.Text) <= (Format$(Fch_BsAs, "dd-mm-yyyy")) And _
                    CDate((RsData!DM) & "-" & mskAnio.Text) <= (Format$(Fch_Int, "dd-mm-yyyy"))) _
                Then
                'Or mskMes.Text < Month(Now)

                  If mskMes.Text = Mid(RsData!DM, 4) Then
                    L_Grilla sprCIA(2), RsData!ant
                    L_Grilla sprCIA(4), RsData!estim
                  Else
                    Anterior_CIA = RsData!ant
                  End If
                End If
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
If mskMes.Text = 2 Then
  sprA(2).SetText 3, sprA(2).MaxRows, str(Anterior_A)
  sprB(2).SetText 3, sprB(2).MaxRows, str(Anterior_B)
  sprC(2).SetText 3, sprC(2).MaxRows, str(Anterior_AEP)
  SprTot(2).SetText 3, SprTot(2).MaxRows, str(Anterior_TOT)
  sprCIA(2).SetText 3, sprCIA(2).MaxRows, str(Anterior_CIA)
  sprD(2).SetText 3, sprD(2).MaxRows, str(Anterior_INT)
End If

ErrGLC:
    Exit Sub

End Sub
Private Sub L_ControlFecha()
Dim FH As String
'Dim FCH_BsAs As String
'Dim FCH_Int As String
Dim rs_pax As Recordset
Dim SQL As String

FH = Format$(Now - 1, FTOFECHA)

SQL = "select min(fch_actualizado) Fch from estadis.control_carga Where Cod_sdep <> 'INT' "

If Aplicacion.ObtenerRsDAO(SQL, rs_pax) Then
   Fch_BsAs = CDate(rs_pax!fch)
End If


SQL = "select min(fch_actualizado) Fch from estadis.control_carga  "

If Aplicacion.ObtenerRsDAO(SQL, rs_pax) Then
   Fch_Int = CDate(rs_pax!fch)
End If

If CDate(FH) <= CDate(Fch_BsAs) Then
   Control_BsAs = True
 Else
   Control_BsAs = False
End If

If CDate(FH) <= CDate(Fch_Int) Then
   Control_Int = True
 Else
   Control_Int = False
End If

End Sub


Private Sub L_LimpiarGrillas()
Dim i

For i = 0 To 5
    sprA(i).MaxRows = 0
    sprAS(i).MaxRows = 0
    sprB(i).MaxRows = 0
    sprC(i).MaxRows = 0
    sprD(i).MaxRows = 0
    SprTot(i).MaxRows = 0
    sprCIA(i).MaxRows = 0
Next
 
End Sub




Private Sub L_PlanillaExcel(Titulo As String, Tipo As String)
Dim AppExcel As Excel.Application
Dim libroExcel As Excel.Workbook
Dim hojaExcel As Excel.Worksheet
Dim i As Integer
Dim nombre As String
Dim tit(1 To 2) As String

On Error GoTo ErrorExl:

tit(1) = " ACTUAL vs ANTERIOR "
tit(2) = " REAL vs. ESTIMADO "

nombre = frmDir.NombreArchivo()
DoEvents

If nombre <> "" Then
    frmResumenDiario.caption = Aplicacion.SeteoProceso(frmResumenDiario.caption)
    Set AppExcel = CreateObject("Excel.application")
    
    Set libroExcel = AppExcel.Workbooks.Add
    
    'AppExcel.Visible = True
        
    Select Case Tipo
        Case "TOTAL"
        For i = 1 To 2
            Set hojaExcel = libroExcel.Worksheets(i)
            hojaExcel.Activate
            L_TratarExel hojaExcel, SprTot(2 - i), SprTot(2 + ((i - 1) * 2)), SprTot(3 + ((i - 1) * 2)), Titulo & tit(i), i, Tipo
        Next
        Case "INTAL"
        For i = 1 To 2
            Set hojaExcel = libroExcel.Worksheets(i)
            hojaExcel.Activate
            L_TratarExel hojaExcel, sprA(2 - i), sprA(2 + ((i - 1) * 2)), sprA(3 + ((i - 1) * 2)), Titulo & tit(i), i, Tipo
        Next
        Case "INTAS"
        For i = 1 To 2
            Set hojaExcel = libroExcel.Worksheets(i)
            hojaExcel.Activate
            L_TratarExel hojaExcel, sprAS(2 - i), sprAS(2 + ((i - 1) * 2)), sprAS(3 + ((i - 1) * 2)), Titulo & tit(i), i, Tipo
        Next
        Case "INTB"
        For i = 1 To 2
            Set hojaExcel = libroExcel.Worksheets(i)
            hojaExcel.Activate
            L_TratarExel hojaExcel, sprB(2 - i), sprB(2 + ((i - 1) * 2)), sprB(3 + ((i - 1) * 2)), Titulo & tit(i), i, Tipo
        Next
        Case "AEP"
        For i = 1 To 2
            Set hojaExcel = libroExcel.Worksheets(i)
            hojaExcel.Activate
            L_TratarExel hojaExcel, sprC(2 - i), sprC(2 + ((i - 1) * 2)), sprC(3 + ((i - 1) * 2)), Titulo & tit(i), i, Tipo
        Next
        
        Case "INT"
        For i = 1 To 2
            Set hojaExcel = libroExcel.Worksheets(i)
            hojaExcel.Activate
            L_TratarExel hojaExcel, sprD(2 - i), sprD(2 + ((i - 1) * 2)), sprD(3 + ((i - 1) * 2)), Titulo & tit(i), i, Tipo
        Next
        
        Case "CIA"
        For i = 1 To 2
            Set hojaExcel = libroExcel.Worksheets(i)
            hojaExcel.Activate
            L_TratarExel hojaExcel, sprCIA(2 - i), sprCIA(2 + ((i - 1) * 2)), sprCIA(3 + ((i - 1) * 2)), Titulo & tit(i), i, Tipo
        Next
        
    End Select
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
      libroExcel.PrintOut
    End If
    
    libroExcel.SaveAs (nombre & ".xls")
    AppExcel.Quit
    Set AppExcel = Nothing
    frmResumenDiario.caption = Aplicacion.SeteoFin
End If
Exit Sub

ErrorExl:
    frmResumenDiario.caption = Aplicacion.SeteoFin
    MsgBox "Planilla no generada", vbOKOnly + vbCritical, "ATENCION"
    Exit Sub
End Sub

Private Sub L_TitulosyFormatos(xls As Object, filaD As Integer, filaH As Integer, hoja As Integer, Tipo As String)
Dim rango As String
Dim i

If hoja = 1 Then
    Exl_PonerValor xls, filaD, 2, "$ Actual "
    Exl.Exl_PonerValor xls, filaD, 3, "$ Anterior"
    Exl.Exl_PonerValor xls, filaD, 4, "Dif %"
    Exl.Exl_PonerValor xls, filaD, 5, "$ Actual Acumulado"
    Exl.Exl_PonerValor xls, filaD, 6, "$ Ant. Acumulado"
    Exl.Exl_PonerValor xls, filaD, 7, "Dif Acum %"
    Exl.Exl_PonerValor xls, filaD, 8, "Pax Actual "
    Exl.Exl_PonerValor xls, filaD, 9, "Pax Anterior"
    Exl.Exl_PonerValor xls, filaD, 10, "Dif %"
    Exl.Exl_PonerValor xls, filaD, 11, "Pax Actual Acumulado"
    Exl.Exl_PonerValor xls, filaD, 12, "Pax Ant. Acumulado"
    Exl.Exl_PonerValor xls, filaD, 13, "Dif Acum %"
    Exl.Exl_PonerValor xls, filaD, 14, "Pr. Pax Actual"
    Exl.Exl_PonerValor xls, filaD, 15, "Pr. Pax Anterior"
    Exl.Exl_PonerValor xls, filaD, 16, "Dif %"
    Exl.Exl_PonerValor xls, filaD, 17, "Pr. Pax Act Acum"
    Exl.Exl_PonerValor xls, filaD, 18, "Pr. Pax Ant Acum"
    Exl.Exl_PonerValor xls, filaD, 19, "Dif Pr. Acum %"

    rango = Exl_rangos(filaD, filaH, 2, 2)
    Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LIzq


    For i = 0 To 2
        rango = Exl_rangos(filaD, filaD, 2 + (i * 6), 7 + (i * 6))
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LArr
        rango = Exl_rangos(filaD, filaH, 7 + (i * 6), 7 + (i * 6))
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LDer
        rango = Exl_rangos(filaH, filaH, 2 + (i * 6), 7 + (i * 6))
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LAba
    Next
    'Lineas y formatos de letras
        rango = Exl_rangos(filaD, filaD, 2, 19)
        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LAba
        Exl.Exl_Justificacion xls, rango, Exl_CentroCol, Exl_DownVert, True
        Exl.Exl_Letra xls, rango, NEGRITA, 8, "Ms Serif"
        rango = Exl_rangos(filaD, filaH, 1, 1)
        Exl.Exl_Letra xls, rango, NEGRITA, 8, "Ms Serif"
        rango = Exl_rangos(filaD, filaH, 4, 4)
        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LDer
        rango = Exl_rangos(filaD, filaH, 10, 10)
        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LDer
        rango = Exl_rangos(filaD, filaH, 16, 16)
        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LDer
    'Achos de columnas
        Exl.Exl_AnchoCol xls, 1, 1, 13
        Exl.Exl_AnchoCol xls, 4, 4, 6
        Exl.Exl_AnchoCol xls, 7, 7, 6
        Exl.Exl_AnchoCol xls, 10, 10, 6
        Exl.Exl_AnchoCol xls, 13, 13, 6
    '    Exl.Exl_AnchoCol xls, 16, 16, 6
    '    Exl.Exl_AnchoCol xls, 19, 19, 6
        Exl.Exl_AnchoCol xls, 8, 9, 9
        Exl.Exl_AnchoCol xls, 11, 12, 9
        Exl.Exl_AnchoCol xls, 14, 19, 7
    
        rango = Exl_rangos(filaD + 1, filaH, 2, 19)
        xls.Range(rango).Font.Size = 8
    'formato de columnas de difencias
        rango = Exl_rangos(filaD + 1, filaH, 4, 4)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 7, 7)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 10, 10)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 13, 13)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 16, 16)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 19, 19)
        Exl.Exl_Format xls, rango
    'formato de las columnas de importes
        rango = Exl_rangos(filaD + 1, filaH, 2, 3)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 5, 6)
        Exl.Exl_Format xls, rango

Else
    Exl_PonerValor xls, filaD, 2, "$ Actual "
    Exl.Exl_PonerValor xls, filaD, 3, "$ Estimado"
    Exl.Exl_PonerValor xls, filaD, 4, "Dif %"
    Exl.Exl_PonerValor xls, filaD, 5, "$ Actual Acumulado"
    Exl.Exl_PonerValor xls, filaD, 6, "$ Estim. Acumulado"
    Exl.Exl_PonerValor xls, filaD, 7, "Dif Acum %"
    Exl.Exl_PonerValor xls, filaD, 8, "Estm. Corregido"

    Exl.Exl_PonerValor xls, filaD, 9, "Pax Actual"
    
    Exl.Exl_PonerValor xls, filaD, 10, "Pax Estimado"
    Exl.Exl_PonerValor xls, filaD, 11, "Dif %"
    Exl.Exl_PonerValor xls, filaD, 12, "Pax Actual Acumulado"
    Exl.Exl_PonerValor xls, filaD, 13, "Pax Estim. Acumulado"
    Exl.Exl_PonerValor xls, filaD, 14, "Dif Acum %"
    Exl.Exl_PonerValor xls, filaD, 15, "Pr. Pax Actual"
    Exl.Exl_PonerValor xls, filaD, 16, "Pr. Pax Estim."
    Exl.Exl_PonerValor xls, filaD, 17, "Dif %"
    Exl.Exl_PonerValor xls, filaD, 18, "Pr. Pax Act Acum"
    Exl.Exl_PonerValor xls, filaD, 19, "Pr. Pax Estim Acum"
    Exl.Exl_PonerValor xls, filaD, 20, "Dif Pr. Acum %"

    rango = Exl_rangos(filaD, filaH, 2, 2)
    Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LIzq

    rango = Exl_rangos(filaD, filaD, 2, 8)
    Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LArr
    rango = Exl_rangos(filaD, filaH, 8, 8)
    Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LDer
    rango = Exl_rangos(filaH, filaH, 2, 8)
    Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LAba

    For i = 1 To 2
        rango = Exl_rangos(filaD, filaD, 3 + (i * 6), 8 + (i * 6))
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LArr
        rango = Exl_rangos(filaD, filaH, 8 + (i * 6), 8 + (i * 6))
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LDer
        rango = Exl_rangos(filaH, filaH, 3 + (i * 6), 8 + (i * 6))
        Exl_LineasPart xls, rango, Exl_LinDoble, Exl_LAba
    Next
    'Lineas y formatos de letras
        rango = Exl_rangos(filaD, filaD, 2, 20)
        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LAba
        Exl.Exl_Justificacion xls, rango, Exl_CentroCol, Exl_DownVert, True
        Exl.Exl_Letra xls, rango, NEGRITA, 8, "Ms Serif"
        rango = Exl_rangos(filaD, filaH, 1, 1)
        Exl.Exl_Letra xls, rango, NEGRITA, 8, "Ms Serif"
        rango = Exl_rangos(filaD, filaH, 4, 4)
        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LDer
        rango = Exl_rangos(filaD, filaH, 11, 11)
        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LDer
        rango = Exl_rangos(filaD, filaH, 17, 17)
        Exl_LineasPart xls, rango, Exl_Linsimple, Exl_LDer
    'Achos de columnas
        Exl.Exl_AnchoCol xls, 1, 1, 13
        Exl.Exl_AnchoCol xls, 4, 4, 6
        Exl.Exl_AnchoCol xls, 7, 7, 6
        Exl.Exl_AnchoCol xls, 11, 11, 6
        Exl.Exl_AnchoCol xls, 14, 14, 6
    '    Exl.Exl_AnchoCol xls, 16, 16, 6
    '    Exl.Exl_AnchoCol xls, 19, 19, 6
        Exl.Exl_AnchoCol xls, 9, 10, 9
        Exl.Exl_AnchoCol xls, 12, 13, 9
        Exl.Exl_AnchoCol xls, 15, 19, 6
    
        rango = Exl_rangos(filaD + 1, filaH, 2, 20)
        xls.Range(rango).Font.Size = 8
    'formato de columnas de difencias
        rango = Exl_rangos(filaD + 1, filaH, 4, 4)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 8, 8)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 11, 11)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 14, 14)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 17, 17)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 20, 20)
        Exl.Exl_Format xls, rango
    'formato de las columnas de importes
        rango = Exl_rangos(filaD + 1, filaH, 2, 3)
        Exl.Exl_Format xls, rango
        rango = Exl_rangos(filaD + 1, filaH, 5, 7)
        Exl.Exl_Format xls, rango

End If

End Sub

Private Sub L_TratarExel(hojaExcel As Object, spr As control, sprP As control, sprPr As control, Titulo As String, hoja As Integer, Tipo As String)
'Dim AppExcel As Excel.Application
'Dim libroExcel As Excel.Workbook
'Dim hojaExcel As Excel.Worksheet
Dim i As Integer, fila As Integer
Dim rango As String
Dim valor As Variant

'Set AppExcel = CreateObject("Excel.application")'

'Set libroExcel = AppExcel.Workbooks.Add

'Set hojaExcel = libroExcel.Worksheets(ind)

'AppExcel.Visible = True

fila = 2
Exl_PonerValor hojaExcel, fila, 1, Titulo
rango = Exl_rangos(fila, fila, 1, 19)
Exl_Justificacion hojaExcel, rango, Exl_Centro, Exl_CentroVert, False
rango = Exl_rangos(fila, fila, 1, 1)
Exl.Exl_Letra hojaExcel, rango, NEGRITA, 12, "Ms Serif"

fila = fila + 2

For i = 1 To spr.MaxRows
    spr.GetText 1, i, valor
    hojaExcel.Cells(fila + i, 1).Value = "'" & Format(valor & "-" & mskAnio.Text, "dd-mmm, dddd")
Next
L_TitulosyFormatos hojaExcel, fila, i + fila - 1, hoja, Tipo

For i = 1 To spr.MaxRows
    spr.GetText 2, i, valor
    Exl_PonerValor hojaExcel, fila + i, 2, valor
    spr.GetText 3, i, valor
    Exl_PonerValor hojaExcel, fila + i, 3, valor
    spr.GetText 5, i, valor
    Exl_PonerValor hojaExcel, fila + i, 4, Format$(valor, "#.00")
    spr.GetText 6, i, valor
    Exl_PonerValor hojaExcel, fila + i, 5, valor
    spr.GetText 7, i, valor
    Exl_PonerValor hojaExcel, fila + i, 6, valor
    spr.GetText 9, i, valor
    Exl_PonerValor hojaExcel, fila + i, 7, Format$(valor, "#.00")
    
    If hoja = 2 Then
        spr.GetText 10, i, valor
        Exl_PonerValor hojaExcel, fila + i, 8, Format$(valor, "#.00")
    End If

Next
For i = 1 To sprP.MaxRows
    sprP.GetText 2, i, valor
    Exl_PonerValor hojaExcel, fila + i, 8 + (hoja - 1), valor
    sprP.GetText 3, i, valor
    Exl_PonerValor hojaExcel, fila + i, 9 + (hoja - 1), valor
    sprP.GetText 5, i, valor
    Exl_PonerValor hojaExcel, fila + i, 10 + (hoja - 1), Format$(valor, "#.00")
    sprP.GetText 6, i, valor
    Exl_PonerValor hojaExcel, fila + i, 11 + (hoja - 1), valor
    sprP.GetText 7, i, valor
    Exl_PonerValor hojaExcel, fila + i, 12 + (hoja - 1), valor
    sprP.GetText 9, i, valor
    Exl_PonerValor hojaExcel, fila + i, 13 + (hoja - 1), Format$(valor, "#.00")
Next
For i = 1 To sprPr.MaxRows
    sprPr.GetText 2, i, valor
    Exl_PonerValor hojaExcel, fila + i, 14 + (hoja - 1), Format$(valor, "#.00")
    sprPr.GetText 3, i, valor
    Exl_PonerValor hojaExcel, fila + i, 15 + (hoja - 1), Format$(valor, "#.00")
    sprPr.GetText 5, i, valor
    Exl_PonerValor hojaExcel, fila + i, 16 + (hoja - 1), Format$(valor, "#.00")
    sprPr.GetText 6, i, valor
    Exl_PonerValor hojaExcel, fila + i, 17 + (hoja - 1), Format$(valor, "#.00")
    sprPr.GetText 7, i, valor
    Exl_PonerValor hojaExcel, fila + i, 18 + (hoja - 1), Format$(valor, "#.00")
    sprPr.GetText 9, i, valor
    Exl_PonerValor hojaExcel, fila + i, 19 + (hoja - 1), Format$(valor, "#.00")
Next
'hojaExcel.PageSetup.PaperSize = Exl_Legal
hojaExcel.PageSetup.Orientation = 2
hojaExcel.PageSetup.Zoom = False
'AppExcel.Quit
'Set AppExcel = Nothing
        '.FitToPagesWide = 1
        '.FitToPagesTall = 1
End Sub

Private Sub L_TratarPromedios(sprS As control, sprD As control, sprR As control)
Dim fila As Integer, filaD As Integer, tope As Integer
Dim fchs As Variant, fchD As Variant
Dim Sor As Variant, Dendo As Variant
Dim i

fila = 1
filaD = 1
sprR.MaxRows = 0
If mskMes.Text = Month(Now) And mskAnio.Text = Year(Now) Then
    tope = Day(Now) - 1
Else
    tope = Day(Func.func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text))
End If

Do While fila <= tope
    sprR.MaxRows = sprR.MaxRows + 1
    sprS.GetText 1, fila, fchs
    sprR.SetText 1, sprR.MaxRows, Trim(fchs)
    
    sprD.GetText 1, filaD, fchD
    
    If fchs = fchD Then
        sprS.GetText 2, fila, Sor
        sprD.GetText 2, fila, Dendo
        If Val(Dendo) > 0 Then
            sprR.SetText 2, fila, Format$((Sor / Dendo), "#.00")
        End If
        sprS.GetText 3, fila, Sor
        sprD.GetText 3, fila, Dendo
        If Val(Dendo) > 0 Then
            sprR.SetText 3, fila, Format$((Sor / Dendo), "#.00")
        End If
        sprS.GetText 6, fila, Sor
        sprD.GetText 6, fila, Dendo
        If Val(Dendo) > 0 Then
            sprR.SetText 6, fila, Format$((Sor / Dendo), "#.00")
        End If
        sprS.GetText 7, fila, Sor
        sprD.GetText 7, fila, Dendo
        If Val(Dendo) > 0 Then
            sprR.SetText 7, fila, Format$((Sor / Dendo), "#.00")
        End If
        
        filaD = filaD + 1
    End If
    spread_ResaltarCelda sprR, 4, fila
    spread_ResaltarCelda sprR, 5, fila
    spread_ResaltarCelda sprR, 8, fila
    spread_ResaltarCelda sprR, 9, fila
    
    fila = fila + 1

Loop
    
End Sub

Private Sub L_VentaNeta(anio As Integer, Per As String)
Dim SQL As String

'CONSULTA DE VENTAS Y DE DUY CHEK
SQL = " select TO_CHAR(V.fch_ticket,'DD-MM') dm ,V.cod_depn,v.cod_sdep,impVta,NVL(impDC,0) impDC "
SQL = SQL & " From "
SQL = SQL & " (select V.fch_ticket,V.cod_depn,decode(v.cod_depn,'INT',v.cod_depn,cod_ssdep) cod_sdep ,sum(V.importe) impVta"
SQL = SQL & " from " & funcLocal_Vista("venta_lgi", anio) & " V , ventas.apertura_sdep A"
SQL = SQL & " Where "
SQL = SQL & L_Armarcondicion("fch_ticket", anio)
SQL = SQL & " and V.cod_depn=A.cod_depn and V.cod_sdep=A.cod_sdep  and V.cod_local=A.cod_local and comitente = 'T'"
SQL = SQL & " group by V.fch_ticket,V.cod_depn,decode(v.cod_depn,'INT',v.cod_depn,cod_ssdep)"
SQL = SQL & " Union"
SQL = SQL & " select V.fch_ticket,'TOT' cod_depn,'TOT' cod_sdep ,sum(V.importe) impVta"
SQL = SQL & " from  " & funcLocal_Vista("venta_lgi", anio) & "  V , ventas.apertura_sdep A"
SQL = SQL & " Where "
SQL = SQL & L_Armarcondicion("fch_ticket", anio)
SQL = SQL & " and V.cod_depn=A.cod_depn and V.cod_sdep=A.cod_sdep and V.cod_local=A.cod_local and comitente = 'T' and v.cod_depn <> 'INT'"
SQL = SQL & " group by V.fch_ticket"
SQL = SQL & " Union"
SQL = SQL & " select V.fch_ticket,'CIA' cod_depn,'CIA' cod_sdep ,sum(V.importe) impVta"
SQL = SQL & " from  " & funcLocal_Vista("venta_lgi", anio) & "  V , ventas.apertura_sdep A"
SQL = SQL & " Where "
SQL = SQL & L_Armarcondicion("fch_ticket", anio)
SQL = SQL & " and V.cod_depn=A.cod_depn and V.cod_sdep=A.cod_sdep  and V.cod_local=A.cod_local and comitente = 'T'"
SQL = SQL & " group by V.fch_ticket ) V ,"
SQL = SQL & "    (select fch_ticket,P.cod_depn,cod_ssdep cod_sdep ,sum(P.importe) impDC"
SQL = SQL & "     from  " & funcLocal_Vista("pago_lpt", anio) & "  P , ventas.apertura_sdep A"
SQL = SQL & "     Where  "
SQL = SQL & L_Armarcondicion("fch_ticket", anio)
SQL = SQL & "     and p.cod_depn=A.cod_depn and P.cod_sdep=A.cod_sdep  and P.cod_local=A.cod_local and tipo_pago = 14"
SQL = SQL & "     group by fch_ticket,P.cod_depn,cod_ssdep"
SQL = SQL & "     Union"
SQL = SQL & "     select fch_ticket,'TOT' cod_depn,'TOT' cod_sdep ,sum(P.importe) impDC"
SQL = SQL & "     from  " & funcLocal_Vista("pago_lpt", anio) & "  P , ventas.apertura_sdep A"
SQL = SQL & "     Where  "
SQL = SQL & L_Armarcondicion("fch_ticket", anio)
SQL = SQL & "     and p.cod_depn=A.cod_depn and P.cod_sdep=A.cod_sdep and P.cod_local=A.cod_local and P.cod_depn <> 'INT' and tipo_pago = 14"
SQL = SQL & "     group by fch_ticket"
SQL = SQL & "     Union"
SQL = SQL & "     select fch_ticket,'CIA' cod_depn,'CIA' cod_sdep ,sum(P.importe) impDC"
SQL = SQL & "     from  " & funcLocal_Vista("pago_lpt", anio) & "  P , ventas.apertura_sdep A"
SQL = SQL & "     Where "
SQL = SQL & L_Armarcondicion("fch_ticket", anio)
SQL = SQL & "     and p.cod_depn=A.cod_depn and P.cod_sdep=A.cod_sdep and P.cod_local=A.cod_local and tipo_pago = 14"
SQL = SQL & "     group by fch_ticket) P"
SQL = SQL & " Where "
SQL = SQL & " v.fch_ticket = p.fch_ticket(+) "
SQL = SQL & " and v.cod_depn=p.cod_depn(+) "
SQL = SQL & " and v.cod_sdep=p.cod_sdep(+) "
SQL = SQL & " order by v.cod_sdep,v.fch_ticket "

If Per = "ACTUAL" Then
    If Not Aplicacion.ObtenerRsDAO(SQL, RsActual) Then
        MsgBox "No se pudo cargar datos de Duty Check ", vbOKOnly + vbInformation, "Atencion"
    End If
Else
    If Not Aplicacion.ObtenerRsDAO(SQL, RsAnterior) Then
        MsgBox "No se pudo cargar datos de Duty Check ", vbOKOnly + vbInformation, "Atencion"
    End If
End If


End Sub

Private Sub botEjecutar_Click(Index As Integer)
Dim i



Select Case Index
    Case 0
    frmResumenDiario.caption = Aplicacion.SeteoProceso(frmResumenDiario.caption)
        'Para cada dependencia
        'L_VentaNeta mskAnio.Text, "ACTUAL"
        'L_VentaNeta mskAnio.Text - 1, "ANTERIOR"
        
        L_Refrescar
        L_RefrescarPax
                
        L_TratarPromedios sprA(1), sprA(2), sprA(3)
        L_TratarPromedios sprAS(1), sprAS(2), sprAS(3)
        L_TratarPromedios sprB(1), sprB(2), sprB(3)
        L_TratarPromedios sprC(1), sprC(2), sprC(3)
        L_TratarPromedios sprD(1), sprD(2), sprD(3)
        L_TratarPromedios sprD(0), sprD(4), sprD(5)
        L_TratarPromedios sprC(0), sprC(4), sprC(5)
        L_TratarPromedios sprB(0), sprB(4), sprB(5)
        L_TratarPromedios sprA(0), sprA(4), sprA(5)
        L_TratarPromedios sprAS(0), sprAS(4), sprAS(5)
        L_TratarPromedios SprTot(1), SprTot(2), SprTot(3)
        L_TratarPromedios SprTot(0), SprTot(4), SprTot(5)
        L_TratarPromedios sprCIA(1), sprCIA(2), sprCIA(3)
        L_TratarPromedios sprCIA(0), sprCIA(4), sprCIA(5)
        For i = 0 To 5
            L_PintarfinSemana_Esp sprA(i)
            L_PintarfinSemana_Esp sprAS(i)
            L_PintarfinSemana_Esp sprB(i)
            L_PintarfinSemana_Esp sprC(i)
            L_PintarfinSemana_Esp sprD(i)
            L_PintarfinSemana_Esp SprTot(i)
            L_PintarfinSemana_Esp sprCIA(i)
        Next
        L_CorregirEstimado SprTot(0)
        L_CorregirEstimado sprA(0)
        L_CorregirEstimado sprAS(0)
        L_CorregirEstimado sprB(0)
        L_CorregirEstimado sprC(0)
        L_CorregirEstimado sprD(0)
        L_CorregirEstimado sprCIA(0)
        botEjecutar(0).Enabled = False
    frmResumenDiario.caption = Aplicacion.SeteoFin
    Case 1
        frCab.Enabled = True
        botEjecutar(0).Enabled = True
'        tabEst.Enabled = False
        L_LimpiarGrillas
    Case 2
        Unload Me
End Select


End Sub

Private Function L_Armarcondicion(campo As String, anio As Integer)
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String

If mskAnio.Text = anio Then
'  If Not (mskMes.Text) = 2 Then
    fechaDesde = Format$(func_Dia1SegunMes_Anio(mskMes.Text, mskAnio.Text), FTOFECHA)
    fechaHasta = Format$(func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text), FTOFECHA)
'  Else
'    fechaDesde = Format$(func_Dia1SegunMes_Anio(mskMes.Text, mskAnio.Text), FTOFECHA)
'    fechaHasta = Format$(func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text), FTOFECHA) - 1
'  End If
Else
'  If Not (mskMes.Text) = 2 Then
    fechaDesde = Format$(CDate(func_Dia1SegunMes_Anio(mskMes.Text, mskAnio.Text)) - Aplicacion.anio, FTOFECHA)
    fechaHasta = Format$(CDate(func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text)) - Aplicacion.anio, FTOFECHA)
'   Else
'    fechaDesde = Format$(CDate(func_Dia1SegunMes_Anio(mskMes.Text, mskAnio.Text)) - Aplicacion.anio, FTOFECHA)
'    fechaHasta = Format$(func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text), FTOFECHA) - 1
'  End If
End If

L_Armarcondicion = " " & campo & " between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

End Function



Private Sub L_Refrescar()
Dim SQL As String, sqlT As String

On Error GoTo ErrEstim:

'frmResumenDiario.caption = Aplicacion.SeteoProceso(frmResumenDiario.caption)

sqlT = " select DM,decode(cod_depn,'INT',cod_depn,cod_sdep)tipo,sum(ImpActual) Act,sum(ImpAnt) Ant,sum(ImpEst) Etim from "
sqlT = sqlT & " (select to_char(fch_ticket,'dd-mm') DM,cod_depn,cod_sdep,"
sqlT = sqlT & " decode(v.tipo,'A',imp,0) ImpActual,"
sqlT = sqlT & " decode(v.tipo,'H',imp,0) ImpAnt,"
sqlT = sqlT & " decode(v.tipo,'E',imp,0) ImpEst"
sqlT = sqlT & " From"
sqlT = sqlT & " (select 'H' tipo,estadis.mismo_dia (fch_ticket," & mskAnio.Text & ") fch_ticket,v.cod_depn,cod_ssdep cod_sdep,sum(importe) imp"
sqlT = sqlT & " From  estadis.venta_lgi_hist v, ventas.apertura_sdep A"
sqlT = sqlT & " where comitente = 'T' AND a.cod_depn=v.cod_depn AND a.cod_sdep=v.cod_sdep and a.cod_local = v.cod_local and "
sqlT = sqlT & L_Armarcondicion("fch_ticket", mskAnio.Text - 1)
sqlT = sqlT & " group by fch_ticket,v.cod_depn,cod_ssdep"
sqlT = sqlT & " Union "
sqlT = sqlT & " select 'A' tipo, fch_ticket,v.cod_depn,cod_ssdep cod_sdep,sum(importe) imp"
sqlT = sqlT & " From " & funcLocal_Vista("venta_lgi", mskAnio.Text) & " v , ventas.apertura_sdep A"
sqlT = sqlT & " where comitente = 'T' AND a.cod_depn=v.cod_depn AND a.cod_sdep=v.cod_sdep and a.cod_local = v.cod_local and "
sqlT = sqlT & L_Armarcondicion("fch_ticket", mskAnio.Text)
sqlT = sqlT & " group by fch_ticket,v.cod_depn,cod_ssdep"
sqlT = sqlT & " Union"
sqlT = sqlT & " Select 'E' tipo,dia,depn cod_depn, sdep, trunc(porc,2) imp"
sqlT = sqlT & " From estadis.porcentaje_d"
sqlT = sqlT & " WHERE anio = " & mskAnio.Text & " And Mes = " & mskMes.Text
sqlT = sqlT & " And Tipo_porc = 'I'  And Nivel = 0"
sqlT = sqlT & " Union "
sqlT = sqlT & " select 'H' tipo,estadis.mismo_dia (fch_ticket," & mskAnio.Text & ") fch_ticket,'TOT' cod_depn,'TOT' cod_sdep,sum(importe) imp"
sqlT = sqlT & " From estadis.venta_lgi_hist "
sqlT = sqlT & " where comitente = 'T' AND " & L_Armarcondicion("fch_ticket", mskAnio.Text - 1)
sqlT = sqlT & " and cod_depn <> 'INT' "
sqlT = sqlT & " group by fch_ticket "
sqlT = sqlT & " Union "
sqlT = sqlT & " select 'A' tipo, fch_ticket,'TOT' cod_depn,'TOT' cod_sdep,sum(importe) imp"
sqlT = sqlT & " From " & funcLocal_Vista("venta_lgi", mskAnio.Text)
sqlT = sqlT & " where comitente = 'T' AND " & L_Armarcondicion("fch_ticket", mskAnio.Text)
sqlT = sqlT & " and cod_depn <> 'INT' "
sqlT = sqlT & " group by fch_ticket "
sqlT = sqlT & " Union"
sqlT = sqlT & " Select 'E' tipo,dia,'TOT' cod_depn, 'TOT' sdep, trunc(porc,2) imp"
sqlT = sqlT & " From estadis.porcentaje_d"
sqlT = sqlT & " WHERE anio = " & mskAnio.Text & " And Mes = " & mskMes.Text
sqlT = sqlT & " And Tipo_porc = 'I'  And Nivel = 0 and depn <> 'INT' and sdep <> 'INTA' "

sqlT = sqlT & " Union "
sqlT = sqlT & " select 'H' tipo,estadis.mismo_dia (fch_ticket," & mskAnio.Text & ") fch_ticket,'CIA' cod_depn,'CIA' cod_sdep,sum(importe) imp"
sqlT = sqlT & " From estadis.venta_lgi_hist "
sqlT = sqlT & " where comitente = 'T' AND " & L_Armarcondicion("fch_ticket", mskAnio.Text - 1)
sqlT = sqlT & " group by fch_ticket "
sqlT = sqlT & " Union "
sqlT = sqlT & " select 'A' tipo, fch_ticket,'CIA' cod_depn,'CIA' cod_sdep,sum(importe) imp"
sqlT = sqlT & " From " & funcLocal_Vista("venta_lgi", mskAnio.Text)
sqlT = sqlT & " where comitente = 'T' AND " & L_Armarcondicion("fch_ticket", mskAnio.Text)
sqlT = sqlT & " group by fch_ticket "
sqlT = sqlT & " Union"
sqlT = sqlT & " Select 'E' tipo,dia,'CIA' cod_depn, 'CIA' sdep, trunc(porc,2) imp"
sqlT = sqlT & " From estadis.porcentaje_d"
sqlT = sqlT & " WHERE anio = " & mskAnio.Text & " And Mes = " & mskMes.Text
sqlT = sqlT & " And Tipo_porc = 'I'  And Nivel = 0  and sdep <> 'INTA' "

sqlT = sqlT & "  ) V"
sqlT = sqlT & " ) U GROUP BY  DM,decode(cod_depn,'INT',cod_depn,cod_sdep) "
sqlT = sqlT & " ORDER BY decode(cod_depn,'INT',cod_depn,cod_sdep),dm "

If Aplicacion.ObtenerRsDAO(sqlT, RsData) Then
    
    L_LlenarGrillas
    Aplicacion.CerrarDAO RsData

End If

ErrEstim:
    'frmResumenDiario.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub

Private Sub L_RefrescarPax()
Dim SQL As String, sqlT As String

On Error GoTo ErrEstim:

'frmResumenDiario.caption = Aplicacion.SeteoProceso(frmResumenDiario.caption)

sqlT = " select to_char(fch_vuelo,'dd-mm') DM,decode(cod_depn,'INT',cod_depn,cod_sdep) tipo ,"
sqlT = sqlT & " sum(paxActual) Act,"
sqlT = sqlT & " sum(paxAnt) Ant,"
sqlT = sqlT & " sum(paxEstim) Estim"
sqlT = sqlT & " from("
sqlT = sqlT & " select fch_vuelo,cod_depn,cod_sdep,"
sqlT = sqlT & " decode(tipo,'A',cant,0) PaxActual,"
sqlT = sqlT & " decode(tipo,'H',cant,0) PaxAnt,"
sqlT = sqlT & " decode(tipo,'E',cant,0) PaxEstim"
sqlT = sqlT & " From"
sqlT = sqlT & " (select 'A' Tipo ,fch_vuelo,v.cod_depn,cod_ssdep cod_sdep,sum(cantidad) cant"
'sqlT = sqlT & " From estadis.pax_volados"
sqlT = sqlT & " From estadis.pax_volados v, ventas.apertura_sdep A"
'sqlT = sqlT & " where " & L_Armarcondicion("fch_vuelo", mskAnio.FormattedText)
sqlT = sqlT & " where a.cod_depn=v.cod_depn AND a.cod_sdep=v.cod_sdep and a.cod_local = v.local and "
sqlT = sqlT & L_Armarcondicion("fch_vuelo", mskAnio.FormattedText)
sqlT = sqlT & " group by fch_vuelo,v.cod_depn,cod_ssdep"
sqlT = sqlT & " Union"
sqlT = sqlT & " select 'H' tipo, estadis.mismo_dia (fch_vuelo," & mskAnio.Text & ") fch_vuelo,v.cod_depn,cod_ssdep cod_sdep,sum(cantidad) cant"
sqlT = sqlT & " From estadis.pax_volados v, ventas.apertura_sdep A "
sqlT = sqlT & " where a.cod_depn=v.cod_depn AND a.cod_sdep=v.cod_sdep and a.cod_local = v.local and "
sqlT = sqlT & L_Armarcondicion("fch_vuelo", mskAnio.FormattedText - 1)
sqlT = sqlT & " group by fch_vuelo,v.cod_depn,cod_ssdep "
sqlT = sqlT & " Union"
sqlT = sqlT & " Select 'E' tipo,dia,depn cod_depn,sdep, trunc(porc,2) imp"
sqlT = sqlT & " From estadis.porcentaje_d"
sqlT = sqlT & " WHERE anio = " & mskAnio.Text & " And Mes = " & mskMes.Text
sqlT = sqlT & " And Tipo_porc = 'P'  And Nivel = 0  and sdep <> 'INTA' "
sqlT = sqlT & " Union"
sqlT = sqlT & " select 'A' Tipo ,fch_vuelo,'TOT' cod_depn,'TOT' cod_sdep,sum(cantidad) cant"
sqlT = sqlT & " From estadis.pax_volados"
sqlT = sqlT & " where " & L_Armarcondicion("fch_vuelo", mskAnio.FormattedText)
sqlT = sqlT & " And cod_depn <> 'INT' "
sqlT = sqlT & " group by fch_vuelo "
sqlT = sqlT & " Union"
sqlT = sqlT & " select 'H' tipo, estadis.mismo_dia (fch_vuelo," & mskAnio.Text & ") fch_vuelo,'TOT' cod_depn ,'TOT' cod_sdep,sum(cantidad) cant"
sqlT = sqlT & " From estadis.pax_volados"
sqlT = sqlT & " where " & L_Armarcondicion("fch_vuelo", mskAnio.FormattedText - 1)
sqlT = sqlT & " And cod_depn <> 'INT' "
sqlT = sqlT & " group by fch_vuelo "
sqlT = sqlT & " Union"
sqlT = sqlT & " Select 'E' tipo,dia,'TOT' cod_depn, 'TOT' sdep, trunc(porc,2) imp"
sqlT = sqlT & " From estadis.porcentaje_d"
sqlT = sqlT & " WHERE anio = " & mskAnio.Text & " And Mes = " & mskMes.Text
sqlT = sqlT & " And Tipo_porc = 'P'  And Nivel = 0 and depn <> 'INT'  and sdep <> 'INTA' "
sqlT = sqlT & " Union"
sqlT = sqlT & " select 'A' Tipo ,fch_vuelo,'CIA' cod_depn,'CIA' cod_sdep,sum(cantidad) cant"
sqlT = sqlT & " From estadis.pax_volados"
sqlT = sqlT & " where " & L_Armarcondicion("fch_vuelo", mskAnio.FormattedText)
sqlT = sqlT & " group by fch_vuelo "
sqlT = sqlT & " Union"
sqlT = sqlT & " select 'H' tipo, estadis.mismo_dia (fch_vuelo," & mskAnio.Text & ")  fch_vuelo,'CIA' cod_depn ,'CIA' cod_sdep,sum(cantidad) cant"
sqlT = sqlT & " From estadis.pax_volados"
sqlT = sqlT & " where " & L_Armarcondicion("fch_vuelo", mskAnio.FormattedText - 1)
sqlT = sqlT & " group by fch_vuelo "
sqlT = sqlT & " Union"
sqlT = sqlT & " Select 'E' tipo,dia,'CIA' cod_depn,'CIA' sdep, trunc(porc,2) imp"
sqlT = sqlT & " From estadis.porcentaje_d"
sqlT = sqlT & " WHERE anio = " & mskAnio.Text & " And Mes = " & mskMes.Text
sqlT = sqlT & " And Tipo_porc = 'P'  And Nivel = 0  and sdep <> 'INTA' )"
sqlT = sqlT & " ) group by to_char(fch_vuelo,'dd-mm'),decode(cod_depn,'INT',cod_depn,cod_sdep) "
sqlT = sqlT & " order by decode(cod_depn,'INT',cod_depn,cod_sdep),dm"

If Aplicacion.ObtenerRsDAO(sqlT, RsData) Then
    
    L_LlenarGrillasPax
    Aplicacion.CerrarDAO RsData

End If

ErrEstim:
'    frmResumenDiario.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub

Private Sub botExcel_Click()

Select Case tabEst.Tab
    Case 0
        L_PlanillaExcel " TOTAL BUENOS AIRES - " & UCase(Format$(mskMes.Text & "-" & mskAnio.Text, "MMMM") & " " & mskAnio.Text), "TOTAL"
    Case 1
        L_PlanillaExcel " INTERNACIONAL 'A'  Llegadas - " & UCase(Format$(mskMes.Text & "-" & mskAnio.Text, "MMMM") & " " & mskAnio.Text), "INTAL"
    Case 2
        L_PlanillaExcel " INTERNACIONAL 'A'  Salidas - " & UCase(Format$(mskMes.Text & "-" & mskAnio.Text, "MMMM") & " " & mskAnio.Text), "INTAS"
    Case 3
        L_PlanillaExcel " INTERNACIONAL 'B' - " & UCase(Format$(mskMes.Text & "-" & mskAnio.Text, "MMMM") & " " & mskAnio.Text), "INTB"
    Case 4
        L_PlanillaExcel " AEROPARQUE - " & UCase(Format$(mskMes.Text & "-" & mskAnio.Text, "MMMM") & " " & mskAnio.Text), "AEP"
    Case 5
        L_PlanillaExcel " TOTAL COMPANIA - " & UCase(Format$(mskMes.Text & "-" & mskAnio.Text, "MMMM") & " " & mskAnio.Text), "CIA"
    Case 6
        L_PlanillaExcel " INTERIOR - " & UCase(Format$(mskMes.Text & "-" & mskAnio.Text, "MMMM") & " " & mskAnio.Text), "INT"
End Select

End Sub



Private Sub chkNeto_Click()
Dim i

AjustarVenta chkNeto.Value, RsActual, 2
AjustarVenta chkNeto.Value, RsAnterior, 3

'poner un L_Resaltar

        L_TratarPromedios sprA(1), sprA(2), sprA(3)
        L_TratarPromedios sprAS(1), sprAS(2), sprAS(3)
        L_TratarPromedios sprB(1), sprB(2), sprB(3)
        L_TratarPromedios sprC(1), sprC(2), sprC(3)
        L_TratarPromedios sprD(1), sprD(2), sprD(3)
        L_TratarPromedios sprD(0), sprD(4), sprD(5)
        L_TratarPromedios sprC(0), sprC(4), sprC(5)
        L_TratarPromedios sprB(0), sprB(4), sprB(5)
        L_TratarPromedios sprA(0), sprA(4), sprA(5)
        L_TratarPromedios sprAS(0), sprAS(4), sprAS(5)
        L_TratarPromedios SprTot(1), SprTot(2), SprTot(3)
        L_TratarPromedios SprTot(0), SprTot(4), SprTot(5)
        L_TratarPromedios sprCIA(1), sprCIA(2), sprCIA(3)
        L_TratarPromedios sprCIA(0), sprCIA(4), sprCIA(5)
    
For i = 1 To 5
    spread_ResaltarCelda sprA(1), 4, i
    spread_ResaltarCelda sprA(1), 5, i
    spread_ResaltarCelda sprA(1), 8, i
    spread_ResaltarCelda sprA(1), 9, i
    spread_ResaltarCelda sprA(0), 4, i
    spread_ResaltarCelda sprA(0), 5, i
    spread_ResaltarCelda sprA(0), 8, i
    spread_ResaltarCelda sprA(0), 9, i
    spread_ResaltarCelda sprAS(1), 4, i
    spread_ResaltarCelda sprAS(1), 5, i
    spread_ResaltarCelda sprAS(1), 8, i
    spread_ResaltarCelda sprAS(1), 9, i
    spread_ResaltarCelda sprAS(0), 4, i
    spread_ResaltarCelda sprAS(0), 5, i
    spread_ResaltarCelda sprAS(0), 8, i
    spread_ResaltarCelda sprAS(0), 9, i
    spread_ResaltarCelda sprB(1), 4, i
    spread_ResaltarCelda sprB(1), 5, i
    spread_ResaltarCelda sprB(1), 8, i
    spread_ResaltarCelda sprB(1), 9, i
    spread_ResaltarCelda sprB(0), 4, i
    spread_ResaltarCelda sprB(0), 5, i
    spread_ResaltarCelda sprB(0), 8, i
    spread_ResaltarCelda sprB(0), 9, i
    spread_ResaltarCelda sprC(1), 4, i
    spread_ResaltarCelda sprC(1), 5, i
    spread_ResaltarCelda sprC(1), 8, i
    spread_ResaltarCelda sprC(1), 9, i
    spread_ResaltarCelda sprC(0), 4, i
    spread_ResaltarCelda sprC(0), 5, i
    spread_ResaltarCelda sprC(0), 8, i
    spread_ResaltarCelda sprC(0), 9, i
    spread_ResaltarCelda sprD(1), 4, i
    spread_ResaltarCelda sprD(1), 5, i
    spread_ResaltarCelda sprD(1), 8, i
    spread_ResaltarCelda sprD(1), 9, i
    spread_ResaltarCelda sprD(0), 4, i
    spread_ResaltarCelda sprD(0), 5, i
    spread_ResaltarCelda sprD(0), 8, i
    spread_ResaltarCelda sprD(0), 9, i
    spread_ResaltarCelda sprCIA(1), 4, i
    spread_ResaltarCelda sprCIA(1), 5, i
    spread_ResaltarCelda sprCIA(1), 8, i
    spread_ResaltarCelda sprCIA(1), 9, i
    spread_ResaltarCelda sprCIA(0), 4, i
    spread_ResaltarCelda sprCIA(0), 5, i
    spread_ResaltarCelda sprCIA(0), 8, i
    spread_ResaltarCelda sprCIA(0), 9, i
    spread_ResaltarCelda SprTot(1), 4, i
    spread_ResaltarCelda SprTot(1), 5, i
    spread_ResaltarCelda SprTot(1), 8, i
    spread_ResaltarCelda SprTot(1), 9, i
    spread_ResaltarCelda SprTot(0), 4, i
    spread_ResaltarCelda SprTot(0), 5, i
    spread_ResaltarCelda SprTot(0), 8, i
    spread_ResaltarCelda SprTot(0), 9, i
Next

End Sub

Private Sub Form_Load()
Dim SQL As String

Top = 30
Left = 210
Height = 7000
Width = 11300

SQL = " SELECT cod_depn,descrip FROM baires.dependencia "
SQL = SQL & " ORDER BY cod_depn"

mskAnio.Text = Year(Date)
mskMes.Text = Format$(Month(Date), "0#")

'txtMesMod.Text = Month(Date)
L_LimpiarGrillas
'frmPrincipal.lstForms.AddItem "frmVsDia"
L_ControlFecha

End Sub


Public Sub L_PintarfinSemana_Esp(ByRef spr As control)
Dim i As Integer
Dim valor As Variant

For i = 1 To spr.MaxRows
    spr.GetText 1, i, valor
    If IsDate(valor & "-" & mskAnio.Text) Then
        If WeekDay(CDate(valor & "-" & mskAnio.Text)) = DOMINGO Or WeekDay(CDate(valor & "-" & mskAnio.Text)) = SABADO Then
            Spread_PintaLinea spr, i, 1, i, spr.MaxCols
        End If
    End If
Next

End Sub


Private Sub Form_Unload(Cancel As Integer)
    FuncLocal_SacarForm "frmVsDia"
End Sub


Private Sub mskAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub mskAnio_LostFocus()
If Val(mskAnio.Text) < 1996 Or Val(mskAnio) > 2050 Then
    mskAnio.Text = Year(Date)
End If

End Sub


Private Sub mskMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub mskMes_LostFocus()
If Val(mskMes.Text) < 1 Or Val(mskMes.Text) > 12 Then
    mskMes.Text = Month(Date)
End If
mskMes.Text = Format$(mskMes.Text, "0#")
End Sub


Private Sub tabEst_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    Select Case tabEst.Tab
        Case 0
'            sprTotal(tabTotal.Tab).SetFocus
        Case 1
'            sprGA(tabGA.Tab).SetFocus
        Case 2
'            sprGB(tabGB.Tab).SetFocus
        Case 3
'            sprGC(tabGC.Tab).SetFocus
    End Select
    
    
ErrT:
    Exit Sub

End Sub

Private Sub tabGA_Click(PreviousTab As Integer)
On Error GoTo ErrT:

'    sprGA(tabGA.Tab).SetFocus
    
ErrT:
    Exit Sub

End Sub

Private Sub tabGB_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    'sprGB(tabGB.Tab).SetFocus
    
ErrT:
    Exit Sub

End Sub

Private Sub tabGC_Click(PreviousTab As Integer)
On Error GoTo ErrT:

   ' sprGC(tabGC.Tab).SetFocus
    
ErrT:
    Exit Sub

End Sub

Private Sub tabTotal_Click(PreviousTab As Integer)

On Error GoTo ErrT:

'    sprTotal(tabTotal.Tab).SetFocus
    
ErrT:
    Exit Sub
End Sub

