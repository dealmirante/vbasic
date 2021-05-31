VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIndic_t 
   Caption         =   "INDICADORES"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   10815
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
      Left            =   255
      TabIndex        =   39
      Top             =   1065
      Width           =   1590
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   8955
      Top             =   1965
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   3810
      Left            =   120
      TabIndex        =   2
      Top             =   1545
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   6720
      _Version        =   327680
      Tabs            =   10
      TabsPerRow      =   10
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
      TabCaption(0)   =   "BS.AS."
      TabPicture(0)   =   "frmIndic_T.frx":0000
      Tab(0).ControlCount=   3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "borGrBar(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "botExcel(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "sprTotalBsAs"
      Tab(0).Control(2).Enabled=   0   'False
      TabCaption(1)   =   "INTA-L"
      TabPicture(1)   =   "frmIndic_T.frx":001C
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sprEzeA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "borGrBar(1)"
      Tab(1).Control(1).Enabled=   -1  'True
      TabCaption(2)   =   "INTA-S"
      TabPicture(2)   =   "frmIndic_T.frx":0038
      Tab(2).ControlCount=   2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "sprEzeAS"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "borGrBar(6)"
      Tab(2).Control(1).Enabled=   -1  'True
      TabCaption(3)   =   "INTER B"
      TabPicture(3)   =   "frmIndic_T.frx":0054
      Tab(3).ControlCount=   2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "sprEzeB"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "borGrBar(2)"
      Tab(3).Control(1).Enabled=   -1  'True
      TabCaption(4)   =   "AEP"
      TabPicture(4)   =   "frmIndic_T.frx":0070
      Tab(4).ControlCount=   2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "sprAep"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "borGrBar(3)"
      Tab(4).Control(1).Enabled=   -1  'True
      TabCaption(5)   =   "TOT CIA"
      TabPicture(5)   =   "frmIndic_T.frx":008C
      Tab(5).ControlCount=   3
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "sprTotal"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "borGrBar(5)"
      Tab(5).Control(1).Enabled=   -1  'True
      Tab(5).Control(2)=   "botExcel(1)"
      Tab(5).Control(2).Enabled=   -1  'True
      TabCaption(6)   =   "INTERIOR"
      TabPicture(6)   =   "frmIndic_T.frx":00A8
      Tab(6).ControlCount=   2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "SprInt"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "borGrBar(4)"
      Tab(6).Control(1).Enabled=   -1  'True
      TabCaption(7)   =   "IN FLIGHT"
      TabPicture(7)   =   "frmIndic_T.frx":00C4
      Tab(7).ControlCount=   1
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "SprFlight"
      Tab(7).Control(0).Enabled=   0   'False
      TabCaption(8)   =   "EZEIZA"
      TabPicture(8)   =   "frmIndic_T.frx":00E0
      Tab(8).ControlCount=   3
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "SprEZE"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "botExcel(2)"
      Tab(8).Control(1).Enabled=   -1  'True
      Tab(8).Control(2)=   "borGrBar(7)"
      Tab(8).Control(2).Enabled=   -1  'True
      TabCaption(9)   =   "CIA + IF"
      TabPicture(9)   =   "frmIndic_T.frx":00FC
      Tab(9).ControlCount=   1
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "SprCIA_IF"
      Tab(9).Control(0).Enabled=   0   'False
      Begin FPSpread.vaSpread SprEZE 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic_T.frx":0118
         TabIndex        =   42
         Top             =   480
         Width           =   8685
      End
      Begin FPSpread.vaSpread SprFlight 
         Height          =   3105
         Left            =   -74775
         OleObjectBlob   =   "frmIndic_T.frx":086B
         TabIndex        =   41
         Top             =   465
         Width           =   8655
      End
      Begin FPSpread.vaSpread sprTotalBsAs 
         Height          =   3105
         Left            =   165
         OleObjectBlob   =   "frmIndic_T.frx":0FBE
         TabIndex        =   40
         Top             =   480
         Width           =   8565
      End
      Begin FPSpread.vaSpread sprEzeB 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic_T.frx":16FF
         TabIndex        =   35
         Top             =   465
         Width           =   8640
      End
      Begin FPSpread.vaSpread sprAep 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic_T.frx":1E40
         TabIndex        =   33
         Top             =   450
         Width           =   8655
      End
      Begin FPSpread.vaSpread sprTotal 
         Height          =   3105
         Left            =   -74775
         OleObjectBlob   =   "frmIndic_T.frx":2593
         TabIndex        =   30
         Top             =   450
         Width           =   8580
      End
      Begin FPSpread.vaSpread sprEzeA 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic_T.frx":2CD4
         TabIndex        =   14
         Top             =   450
         Width           =   8565
      End
      Begin FPSpread.vaSpread sprEzeAS 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic_T.frx":3415
         TabIndex        =   37
         Top             =   450
         Width           =   8580
      End
      Begin FPSpread.vaSpread SprInt 
         Height          =   3105
         Left            =   -74820
         OleObjectBlob   =   "frmIndic_T.frx":3B56
         TabIndex        =   28
         Top             =   450
         Width           =   8670
      End
      Begin FPSpread.vaSpread SprCIA_IF 
         Height          =   3105
         Left            =   -74850
         OleObjectBlob   =   "frmIndic_T.frx":42A9
         TabIndex        =   45
         Top             =   420
         Width           =   8580
      End
      Begin VB.CommandButton botExcel 
         Caption         =   "Excel"
         Height          =   510
         Index           =   2
         Left            =   -65940
         Picture         =   "frmIndic_T.frx":49EA
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3045
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   7
         Left            =   -65955
         Picture         =   "frmIndic_T.frx":4F7C
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2355
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   6
         Left            =   -66000
         Picture         =   "frmIndic_T.frx":53BE
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3030
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   2
         Left            =   -65985
         Picture         =   "frmIndic_T.frx":5800
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3030
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   3
         Left            =   -65985
         Picture         =   "frmIndic_T.frx":5C42
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3030
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   5
         Left            =   -65970
         Picture         =   "frmIndic_T.frx":6084
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2370
         Width           =   780
      End
      Begin VB.CommandButton botExcel 
         Caption         =   "Excel"
         Height          =   510
         Index           =   1
         Left            =   -65955
         Picture         =   "frmIndic_T.frx":64C6
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3060
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   4
         Left            =   -66000
         Picture         =   "frmIndic_T.frx":6A58
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3045
         Width           =   780
      End
      Begin VB.CommandButton botExcel 
         Caption         =   "Excel"
         Height          =   510
         Index           =   0
         Left            =   8925
         Picture         =   "frmIndic_T.frx":6E9A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3030
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   0
         Left            =   8925
         Picture         =   "frmIndic_T.frx":742C
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2385
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   1
         Left            =   -66015
         Picture         =   "frmIndic_T.frx":786E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3030
         Width           =   780
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
         Picture         =   "frmIndic_T.frx":7CB0
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
         Picture         =   "frmIndic_T.frx":7DB2
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
         Picture         =   "frmIndic_T.frx":7EB4
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
            TabIndex        =   18
            Top             =   210
            Width           =   3015
            Begin VB.CommandButton botHelpFDAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmIndic_T.frx":86D6
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmIndic_T.frx":8848
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   600
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesdeAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   21
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
               TabIndex        =   22
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
               TabIndex        =   24
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
               TabIndex        =   23
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
            TabIndex        =   16
            Top             =   375
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.OptionButton optFechas 
            Caption         =   "Acumulado Rango de Fechas"
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
            TabIndex        =   15
            Top             =   525
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
               Picture         =   "frmIndic_T.frx":89BA
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2550
               Picture         =   "frmIndic_T.frx":8B2C
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
      Height          =   420
      Left            =   150
      TabIndex        =   25
      Top             =   5400
      Width           =   8730
   End
End
Attribute VB_Name = "frmIndic_t"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset
Dim RsDataAnt  As Recordset
Dim RsDataEstim  As Recordset
Dim RsDataVol  As Recordset

Dim RsActual  As Recordset
Dim RsAnterior  As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas
                       'con porcentajes
Dim UsoPorIntB As Boolean
Dim UsoPorAero As Boolean
Dim UsoPorTotal As Boolean
Dim UsoPorInte As Boolean
Dim UsoPorFlight As Boolean

Dim Control_BsAs As Boolean
Dim Control_Int As Boolean

Dim Fch_BsAs As String
Dim Fch_Int As String

Public Vol_Flight As Variant
Dim Vol_Flight_Ant As Long
Private Sub AjustarVenta(Tipo As Integer)
Dim valorTicket As Variant, valorPax As Variant

On Error GoTo err1:

RsActual.MoveFirst
Do While Not RsActual.EOF
    Select Case RsActual!cod_sdep
        Case "AEP"
            SprAep.GetText 2, 2, valorTicket
            SprAep.GetText 2, 4, valorPax
            If Tipo = 0 Then
             SprAep.SetText 2, 1, str(RsActual!impvta)
             SprAep.SetText 2, 3, str(RsActual!impvta / valorTicket)
             If valorPax > 0 Then
                SprAep.SetText 2, 5, str(RsActual!impvta / valorPax)
             End If
            Else
             SprAep.SetText 2, 1, str(RsActual!impvta - RsActual!impdc)
             SprAep.SetText 2, 3, str((RsActual!impvta - RsActual!impdc) / valorTicket)
             If valorPax > 0 Then
                SprAep.SetText 2, 5, str((RsActual!impvta - RsActual!impdc) / valorPax)
             End If
            End If
            
        Case "INTAL"
            SprEzeA.GetText 2, 2, valorTicket
            SprEzeA.GetText 2, 4, valorPax
            If Tipo = 0 Then
             SprEzeA.SetText 2, 1, str(RsActual!impvta)
             SprEzeA.SetText 2, 3, str(RsActual!impvta / valorTicket)
             If valorPax > 0 Then
                SprEzeA.SetText 2, 5, str(RsActual!impvta / valorPax)
             End If
            Else
             SprEzeA.SetText 2, 1, str(RsActual!impvta - RsActual!impdc)
             SprEzeA.SetText 2, 3, str((RsActual!impvta - RsActual!impdc) / valorTicket)
             If valorPax > 0 Then
                SprEzeA.SetText 2, 5, str((RsActual!impvta - RsActual!impdc) / valorPax)
             End If
            End If
        Case "INTAS"
            SprEzeAS.GetText 2, 2, valorTicket
            SprEzeAS.GetText 2, 4, valorPax
            If Tipo = 0 Then
             SprEzeAS.SetText 2, 1, str(RsActual!impvta)
             SprEzeAS.SetText 2, 3, str(RsActual!impvta / valorTicket)
             If valorPax > 0 Then
                SprEzeAS.SetText 2, 5, str(RsActual!impvta / valorPax)
             End If
            Else
             SprEzeAS.SetText 2, 1, str(RsActual!impvta - RsActual!impdc)
             SprEzeAS.SetText 2, 3, str((RsActual!impvta - RsActual!impdc) / valorTicket)
             If valorPax > 0 Then
                SprEzeAS.SetText 2, 5, str((RsActual!impvta - RsActual!impdc) / valorPax)
             End If
            End If
        Case "INTB"
            SprEzeB.GetText 2, 2, valorTicket
            SprEzeB.GetText 2, 4, valorPax
            If Tipo = 0 Then
             SprEzeB.SetText 2, 1, str(RsActual!impvta)
             SprEzeB.SetText 2, 3, str(RsActual!impvta / valorTicket)
             If valorPax > 0 Then
                SprEzeB.SetText 2, 5, str(RsActual!impvta / valorPax)
             End If
            Else
             SprEzeB.SetText 2, 1, str(RsActual!impvta - RsActual!impdc)
             SprEzeB.SetText 2, 3, str((RsActual!impvta - RsActual!impdc) / valorTicket)
             If valorPax > 0 Then
                SprEzeB.SetText 2, 5, str((RsActual!impvta - RsActual!impdc) / valorPax)
             End If
            End If
        
        Case "FLI"
            SprFlight.GetText 2, 2, valorTicket
            SprFlight.GetText 2, 4, valorPax
            If Tipo = 0 Then
             SprFlight.SetText 2, 1, str(RsActual!impvta)
             SprFlight.SetText 2, 3, str(RsActual!impvta / valorTicket)
             If valorPax > 0 Then
                SprFlight.SetText 2, 5, str(RsActual!impvta / valorPax)
             End If
            Else
             SprFlight.SetText 2, 1, str(RsActual!impvta - RsActual!impdc)
             SprFlight.SetText 2, 3, str((RsActual!impvta - RsActual!impdc) / valorTicket)
             If valorPax > 0 Then
                SprFlight.SetText 2, 5, str((RsActual!impvta - RsActual!impdc) / valorPax)
             End If
            End If
       
        Case "INT"
            If Tipo = 0 Then
             SprInt.SetText 2, 1, str(RsActual!impvta)
            Else
             SprInt.SetText 2, 1, str(RsActual!impvta - RsActual!impdc)
            End If
    End Select
    RsActual.MoveNext
Loop
L_RellenarDatos 1, 2, ""
L_RellenarDatos 3, 2, ""
L_RellenarDatos 5, 2, ""
L_RellenarDatos_BsAs 1, 2, ""
L_RellenarDatos_BsAs 3, 2, ""
L_RellenarDatos_BsAs 5, 2, ""
L_RellenarDatos_EZE 1, 2, ""
L_RellenarDatos_EZE 3, 2, ""
L_RellenarDatos_EZE 5, 2, ""
L_RellenarDatos_CIA_IF 1, 2, ""
L_RellenarDatos_CIA_IF 3, 2, ""
L_RellenarDatos_CIA_IF 5, 2, ""


err1:
  Resume err12
  
err12:
On Error GoTo err2:

RsAnterior.MoveFirst
Do While Not RsAnterior.EOF
    Select Case RsAnterior!cod_sdep
        Case "AEP"
            SprAep.GetText 3, 2, valorTicket
            SprAep.GetText 3, 4, valorPax
            If Tipo = 0 Then
             SprAep.SetText 3, 1, str(RsAnterior!impvta)
             SprAep.SetText 3, 3, str(RsAnterior!impvta / valorTicket)
             If valorPax > 0 Then
                SprAep.SetText 3, 5, str(RsAnterior!impvta / valorPax)
             End If
            Else
             SprAep.SetText 3, 1, str(RsAnterior!impvta - RsAnterior!impdc)
             SprAep.SetText 3, 3, str((RsAnterior!impvta - RsAnterior!impdc) / valorTicket)
             If valorPax > 0 Then
                SprAep.SetText 3, 5, str((RsAnterior!impvta - RsAnterior!impdc) / valorPax)
             End If
            End If
        Case "INTAL"
            SprEzeA.GetText 3, 2, valorTicket
            SprEzeA.GetText 3, 4, valorPax
            If Tipo = 0 Then
             SprEzeA.SetText 3, 1, str(RsAnterior!impvta)
             SprEzeA.SetText 3, 3, str(RsAnterior!impvta / valorTicket)
             If valorPax > 0 Then
                SprEzeA.SetText 3, 5, str(RsAnterior!impvta / valorPax)
             End If
            Else
             SprEzeA.SetText 3, 1, str(RsAnterior!impvta - RsAnterior!impdc)
             SprEzeA.SetText 3, 3, str((RsAnterior!impvta - RsAnterior!impdc) / valorTicket)
             If valorPax > 0 Then
                SprEzeA.SetText 3, 5, str((RsAnterior!impvta - RsAnterior!impdc) / valorPax)
             End If
            End If
        Case "INTAS"
            SprEzeAS.GetText 3, 2, valorTicket
            SprEzeAS.GetText 3, 4, valorPax
            If Tipo = 0 Then
             SprEzeAS.SetText 3, 1, str(RsAnterior!impvta)
             SprEzeAS.SetText 3, 3, str(RsAnterior!impvta / valorTicket)
             If valorPax > 0 Then
                SprEzeAS.SetText 3, 5, str(RsAnterior!impvta / valorPax)
             End If
            Else
             SprEzeAS.SetText 3, 1, str(RsAnterior!impvta - RsAnterior!impdc)
             SprEzeAS.SetText 3, 3, str((RsAnterior!impvta - RsAnterior!impdc) / valorTicket)
             If valorPax > 0 Then
                SprEzeAS.SetText 3, 5, str((RsAnterior!impvta - RsAnterior!impdc) / valorPax)
             End If
            End If
        Case "INTB"
            SprEzeB.GetText 3, 2, valorTicket
            SprEzeB.GetText 3, 4, valorPax
            If Tipo = 0 Then
             SprEzeB.SetText 3, 1, str(RsAnterior!impvta)
             SprEzeB.SetText 3, 3, str(RsAnterior!impvta / valorTicket)
             If valorPax > 0 Then
                SprEzeB.SetText 3, 5, str(RsAnterior!impvta / valorPax)
             End If
            Else
             SprEzeB.SetText 3, 1, str(RsAnterior!impvta - RsAnterior!impdc)
             SprEzeB.SetText 3, 3, str((RsAnterior!impvta - RsAnterior!impdc) / valorTicket)
             If valorPax > 0 Then
                SprEzeB.SetText 3, 5, str((RsAnterior!impvta - RsAnterior!impdc) / valorPax)
             End If
            End If
            
        Case "FLI"
            SprFlight.GetText 3, 2, valorTicket
            SprFlight.GetText 3, 4, valorPax
            If Tipo = 0 Then
             SprFlight.SetText 3, 1, str(RsAnterior!impvta)
             SprFlight.SetText 3, 3, str(RsAnterior!impvta / valorTicket)
             If valorPax > 0 Then
                SprFlight.SetText 3, 5, str(RsAnterior!impvta / valorPax)
             End If
            Else
             SprFlight.SetText 3, 1, str(RsAnterior!impvta - RsAnterior!impdc)
             SprFlight.SetText 3, 3, str((RsAnterior!impvta - RsAnterior!impdc) / valorTicket)
             If valorPax > 0 Then
                SprFlight.SetText 3, 5, str((RsAnterior!impvta - RsAnterior!impdc) / valorPax)
             End If
            End If
            
        Case "INT"
            If Tipo = 0 Then
             SprInt.SetText 3, 1, str(RsAnterior!impvta)
            Else
             SprInt.SetText 3, 1, str(RsAnterior!impvta - RsAnterior!impdc)
            End If
    End Select
    RsAnterior.MoveNext
Loop
L_RellenarDatos 1, 3, ""
L_RellenarDatos 3, 3, ""
L_RellenarDatos 5, 3, ""
L_RellenarDatos_BsAs 1, 3, ""
L_RellenarDatos_BsAs 3, 3, ""
L_RellenarDatos_BsAs 5, 3, ""
L_RellenarDatos_EZE 1, 3, ""
L_RellenarDatos_EZE 3, 3, ""
L_RellenarDatos_EZE 5, 3, ""
L_RellenarDatos_CIA_IF 1, 3, ""
L_RellenarDatos_CIA_IF 3, 3, ""
L_RellenarDatos_CIA_IF 5, 3, ""

err2:
   Exit Sub
End Sub

Private Sub L_ControlFecha()
Dim FH As String
'Dim Fch_BsAs As String
'Dim Fch_Int As String
Dim rs_pax As Recordset
Dim SQL As String

FH = mskFHasta.FormattedText

SQL = "select min(fch_actualizado) Fch from estadis.control_carga Where Cod_sdep <> 'INT' "

If Aplicacion.ObtenerRsDAO(SQL, rs_pax) Then
   Fch_BsAs = CDate(rs_pax!fch)
End If


SQL = "select min(fch_actualizado) Fch from estadis.control_carga Where Cod_sdep = 'INT' "

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
Cond = Cond & " AND A.COD_LOCAL = S.COD_LOCAL and  A.COD_depn = S.COD_depn and  A.COD_sdep = S.COD_sdep "
L_Armarcondicion = Cond

End Function
Private Function L_ArmarcondicionIFL(anio As Integer) As String
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

Cond = " WHERE fch_vuelo between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)
L_ArmarcondicionIFL = Cond

End Function

Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant

Do While Not RsData.EOF
        Select Case RsData!cod_depn
            Case DSLoc(1).Dep

                SprAep.SetText 2, 1, str(RsData!imp)
                
                SprAep.SetText 2, 2, str(RsData!cant_t)
                
                SprAep.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                                   
            
            Case DSLoc(2).Dep
                Select Case RsData!cod_sdep
                    Case DSLoc(2).Sdep

                        SprEzeA.SetText 2, 1, str(RsData!imp)
                        
                        SprEzeA.SetText 2, 2, str(RsData!cant_t)
                        
                        SprEzeA.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                        

                    Case DSLoc(3).Sdep
                        SprEzeB.SetText 2, 1, str(RsData!imp)
                        
                        SprEzeB.SetText 2, 2, str(RsData!cant_t)
                        
                        SprEzeB.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                    
                    Case DSLoc(9).Sdep

                        SprEzeA.SetText 2, 1, str(RsData!imp)
                        
                        SprEzeA.SetText 2, 2, str(RsData!cant_t)
                        
                        SprEzeA.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                        

                    Case DSLoc(10).Sdep
                        SprEzeAS.SetText 2, 1, str(RsData!imp)
                        
                        SprEzeAS.SetText 2, 2, str(RsData!cant_t)
                        
                        SprEzeAS.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                                        
                End Select
                
            Case DSLoc(4).Dep

                        SprInt.SetText 2, 1, str(RsData!imp)
                        
                        SprInt.SetText 2, 2, str(RsData!cant_t)
                        
                        SprInt.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                        
            Case DSLoc(11).Dep
                SprFlight.SetText 2, 1, str(RsData!imp)
                
                SprFlight.SetText 2, 2, str(RsData!cant_t)
                SprFlight.SetText 2, 3, "0"
                'SprFlight.SetText 2, 3, str(RsData!imp / RsData!cant_t)

        End Select
                
        RsData.MoveNext

Loop


End Sub
Private Sub L_DecoEspigonAnt()
Dim fecha As String
Dim i As Integer
Dim dato As Variant

Do While Not RsDataAnt.EOF
        Select Case RsDataAnt!cod_depn
            Case DSLoc(1).Dep
                SprAep.SetText 3, 1, str(RsDataAnt!imp)
                SprAep.SetText 3, 2, str(RsDataAnt!cant_t)
                SprAep.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
            Case DSLoc(2).Dep
                Select Case RsDataAnt!cod_sdep
                    Case DSLoc(2).Sdep
                        SprEzeA.SetText 3, 1, str(RsDataAnt!imp)
                        SprEzeA.SetText 3, 2, str(RsDataAnt!cant_t)
                        SprEzeA.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                    Case DSLoc(3).Sdep
                        SprEzeB.SetText 3, 1, str(RsDataAnt!imp)
                        SprEzeB.SetText 3, 2, str(RsDataAnt!cant_t)
                        SprEzeB.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                    Case DSLoc(9).Sdep
                        SprEzeA.SetText 3, 1, str(RsDataAnt!imp)
                        SprEzeA.SetText 3, 2, str(RsDataAnt!cant_t)
                        SprEzeA.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                    Case DSLoc(10).Sdep
                        SprEzeAS.SetText 3, 1, str(RsDataAnt!imp)
                        SprEzeAS.SetText 3, 2, str(RsDataAnt!cant_t)
                        SprEzeAS.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                End Select
                
             Case DSLoc(11).Dep
                  SprFlight.SetText 3, 1, str(RsDataAnt!imp)
                  SprFlight.SetText 3, 2, str(RsDataAnt!cant_t)
                  SprFlight.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                
            Case DSLoc(4).Dep
                SprInt.SetText 3, 1, str(RsDataAnt!imp)
                SprInt.SetText 3, 2, str(RsDataAnt!cant_t)
                SprInt.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
             
        End Select
                
        RsDataAnt.MoveNext

Loop

End Sub
Private Sub L_DecoEstim()
Dim ImpAep As Double, ImpA As Double, ImpAS As Double, ImpB As Double, ImpFL As Double
Dim valor As Variant

Do While Not RsDataEstim.EOF
   Select Case RsDataEstim!cod_depn
          Case DSLoc(1).Dep
               Select Case RsDataEstim!tipo_porc
                      Case "I"
                        SprAep.SetText 6, 1, str(RsDataEstim!valor)
                        ImpAep = RsDataEstim!valor
                      Case "T"
                        SprAep.SetText 6, 2, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            SprAep.SetText 6, 3, str(ImpAep / RsDataEstim!valor)
                        End If
                        SprAep.GetText 6, 4, valor
                        If Val(valor) > 0 Then
                            SprAep.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                        End If
                    
                      Case "P"
                        SprAep.SetText 6, 4, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            SprAep.SetText 6, 5, str(ImpAep / RsDataEstim!valor)
                        End If
                
               End Select

            Case DSLoc(2).Dep
                Select Case RsDataEstim!cod_sdep
                    Case DSLoc(3).Sdep
                            Select Case RsDataEstim!tipo_porc
                                Case "I"
                                    SprEzeB.SetText 6, 1, str(RsDataEstim!valor)
                                    ImpB = RsDataEstim!valor
                                Case "T"
                                    SprEzeB.SetText 6, 2, str(RsDataEstim!valor)
                                    If RsDataEstim!valor > 0 Then
                                        SprEzeB.SetText 6, 3, str(ImpB / RsDataEstim!valor)
                                    End If
                                    SprEzeB.GetText 6, 4, valor
                                    If Val(valor) > 0 Then
                                        SprEzeB.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                                    End If
                                    
                                Case "P"
                                    SprEzeB.SetText 6, 4, str(RsDataEstim!valor)
                                    If RsDataEstim!valor > 0 Then
                                        SprEzeB.SetText 6, 5, str(ImpB / RsDataEstim!valor)
                                    End If
                            End Select

                    Case DSLoc(9).Sdep
                        Select Case RsDataEstim!tipo_porc
                            Case "I"
                                SprEzeA.SetText 6, 1, str(RsDataEstim!valor)
                                ImpA = RsDataEstim!valor
                            Case "T"
                                SprEzeA.SetText 6, 2, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    SprEzeA.SetText 6, 3, str(ImpA / RsDataEstim!valor)
                                End If
                                SprEzeA.GetText 6, 4, valor
                                If Val(valor) > 0 Then
                                    SprEzeA.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                                End If
                            Case "P"
                                SprEzeA.SetText 6, 4, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    SprEzeA.SetText 6, 5, str(ImpA / RsDataEstim!valor)
                                End If
                            
                        End Select
                            
                    Case DSLoc(10).Sdep
                        Select Case RsDataEstim!tipo_porc
                            Case "I"
                                SprEzeAS.SetText 6, 1, str(RsDataEstim!valor)
                                ImpAS = RsDataEstim!valor
                            Case "T"
                                SprEzeAS.SetText 6, 2, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    SprEzeAS.SetText 6, 3, str(ImpAS / RsDataEstim!valor)
                                End If
                                SprEzeAS.GetText 6, 4, valor
                                If Val(valor) > 0 Then
                                    SprEzeAS.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                                End If
                            Case "P"
                                SprEzeAS.SetText 6, 4, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    SprEzeAS.SetText 6, 5, str(ImpAS / RsDataEstim!valor)
                                End If
                        End Select
                End Select
            Case DSLoc(11).Dep
                Select Case RsDataEstim!tipo_porc
                    Case "I"
                        SprFlight.SetText 6, 1, str(RsDataEstim!valor)
                        ImpFL = RsDataEstim!valor
                    Case "T"
                        SprFlight.SetText 6, 2, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            SprFlight.SetText 6, 3, str(ImpFL / RsDataEstim!valor)
                        End If
                        SprFlight.GetText 6, 4, valor
                        If Val(valor) > 0 Then
                            SprFlight.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                        End If
                    Case "P"
                        SprFlight.SetText 6, 4, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            SprFlight.SetText 6, 5, str(ImpFL / RsDataEstim!valor)
                        End If
                    
                End Select
                                        
            Case DSLoc(4).Dep
                Select Case RsDataEstim!tipo_porc
                    Case "I"
                        SprInt.SetText 6, 1, str(RsDataEstim!valor)
                        ImpAep = RsDataEstim!valor
                    Case "T"
                        SprInt.SetText 6, 2, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            SprInt.SetText 6, 3, str(ImpAep / RsDataEstim!valor)
                        End If
                        SprInt.GetText 6, 4, valor
                        If Val(valor) > 0 Then
                            SprInt.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                        End If
                    
                    Case "P"
                        SprInt.SetText 6, 4, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            SprInt.SetText 6, 5, str(ImpAep / RsDataEstim!valor)
                        End If
                                                            
                End Select
        End Select
                
        RsDataEstim.MoveNext

Loop
End Sub


Private Sub L_DecoVolados(col As Integer)
Dim valor As Variant

Do While Not RsDataVol.EOF
        Select Case RsDataVol!cod_depn
            Case DSLoc(1).Dep
                   SprAep.SetText col, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprAep.GetText col, 1, valor
                      SprAep.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                      SprAep.GetText col, 2, valor
                      SprAep.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
            Case DSLoc(2).Dep
                Select Case RsDataVol!cod_sdep
                    Case DSLoc(2).Sdep
                        SprEzeA.SetText col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprEzeA.GetText col, 1, valor
                            SprEzeA.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            SprEzeA.GetText col, 2, valor
                            SprEzeA.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(3).Sdep
                        SprEzeB.SetText col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprEzeB.GetText col, 1, valor
                            SprEzeB.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            SprEzeB.GetText col, 2, valor
                            SprEzeB.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(9).Sdep
                        SprEzeA.SetText col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprEzeA.GetText col, 1, valor
                            SprEzeA.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            SprEzeA.GetText col, 2, valor
                            SprEzeA.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(10).Sdep
                        SprEzeAS.SetText col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprEzeAS.GetText col, 1, valor
                            SprEzeAS.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            SprEzeAS.GetText col, 2, valor
                            SprEzeAS.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                
                End Select
        
            Case DSLoc(4).Dep
               If (Control_Int And col = 2) Or col = 3 Then
                   SprInt.SetText col, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprInt.GetText col, 1, valor
                      SprInt.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                      SprInt.GetText col, 2, valor
                      SprInt.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
                 Else
                   labInfo.caption = "Datos de pasajeros transitados del Interior al día " & Fch_Int
                   Timer1.Enabled = True
                   Exit Do
               End If
        End Select
                
        RsDataVol.MoveNext
Loop

SprFlight.SetText col, 4, str(Vol_Flight)

If Vol_Flight > 0 Then
   SprFlight.GetText col, 1, valor
   SprFlight.SetText col, 5, Format$((valor / Vol_Flight), "#.00")
   SprFlight.GetText col, 2, valor
   SprFlight.SetText col, 6, Format$((valor / Vol_Flight * 100), "00.00")
End If

End Sub
Private Sub L_DecoVoladosHIST_(col As Integer)
Dim valor As Variant

Do While Not RsDataVol.EOF
        Select Case RsDataVol!cod_depn
            Case DSLoc(1).Dep
                   SprAep.SetText 3, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprAep.GetText 3, 1, valor
                      SprAep.SetText 3, 5, str(valor / RsDataVol!cant_v)
                      SprAep.GetText 3, 2, valor
                      SprAep.SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
            Case DSLoc(2).Dep
                Select Case RsDataVol!cod_sdep
                    Case DSLoc(2).Sdep
                        SprEzeA.SetText 3, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprEzeA.GetText 3, 1, valor
                            SprEzeA.SetText 3, 5, str(valor / RsDataVol!cant_v)
                            SprEzeA.GetText 3, 2, valor
                            SprEzeA.SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(3).Sdep
                        SprEzeB.SetText 2, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprEzeB.GetText 2, 1, valor
                            SprEzeB.SetText 2, 5, str(valor / RsDataVol!cant_v)
                            SprEzeB.GetText 2, 2, valor
                            SprEzeB.SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(11).Sdep
                        SprFlight.SetText 2, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprFlight.GetText 2, 1, valor
                            SprFlight.SetText 2, 5, str(valor / RsDataVol!cant_v)
                            SprFlight.GetText 2, 2, valor
                            SprFlight.SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                End Select
                
            Case DSLoc(4).Dep
                   SprInt.SetText 3, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprInt.GetText 3, 1, valor
                      SprInt.SetText 3, 5, str(valor / RsDataVol!cant_v)
                      SprInt.GetText 3, 2, valor
                      SprInt.SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
                
        End Select
                
        RsDataVol.MoveNext
Loop

    If Vol_Flight_Ant > 0 Then
       SprFlight.GetText col, 1, valor
       SprFlight.SetText col, 5, Format$((valor / Vol_Flight_Ant), "#.00")
       SprFlight.GetText col, 2, valor
       SprFlight.SetText col, 6, Format$((valor / Vol_Flight_Ant * 100), "00.00")
    End If


End Sub


Private Sub L_LimpiarGrillas()
Dim i
    
    SprEzeA.MaxRows = 0
    SprEzeAS.MaxRows = 0
    SprEzeB.MaxRows = 0
    SprAep.MaxRows = 0
    SprInt.MaxRows = 0
    SprFlight.MaxRows = 0
    SprTotal.MaxRows = 0
    SprTotalBsAs.MaxRows = 0
    SprEZE.MaxRows = 0
    SprCIA_IF.MaxRows = 0
    labInfo.caption = ""
    Timer1.Enabled = False
End Sub

Private Sub L_LlenarTotales()
Dim Nom As String
Dim i As Integer

For i = 1 To 6
    SprTotal.MaxRows = SprTotal.MaxRows + 1
    
    Select Case i
        Case 1
            Nom = "Facturación"
        Case 2
            Nom = "Tickets"
        Case 3
            Nom = "Pr. x Tickets"
        Case 4
            Nom = "Pax Viajados"
       Case 5
            Nom = "Pr x Pax Viaj."
        Case 6
            Nom = "Captación"
    End Select
    
    SprTotal.SetText 1, i, Nom
    
     If i < 4 Or Control_Int Then
        L_RellenarDatos i, 2, Nom
     End If
        L_RellenarDatos i, 3, Nom
        L_RellenarDatos i, 6, Nom

Next
End Sub
Private Sub L_LlenarTotalesBsAs()
Dim Nom As String
Dim i As Integer

For i = 1 To 6
    SprTotalBsAs.MaxRows = SprTotalBsAs.MaxRows + 1
    
    Select Case i
        Case 1
            Nom = "Facturación"
        Case 2
            Nom = "Tickets"
        Case 3
            Nom = "Pr. x Tickets"
        Case 4
            Nom = "Pax Viajados"
        Case 5
            Nom = "Pr x Pax Viaj."
        Case 6
            Nom = "Captación"
    End Select
    
    SprTotalBsAs.SetText 1, i, Nom
    
        L_RellenarDatos_BsAs i, 2, Nom
        L_RellenarDatos_BsAs i, 3, Nom
        L_RellenarDatos_BsAs i, 6, Nom

Next
End Sub

Private Sub L_LlenarTotalesEZEIZA()
Dim Nom As String
Dim i As Integer

For i = 1 To 6
    SprEZE.MaxRows = SprEZE.MaxRows + 1
    
    Select Case i
        Case 1
            Nom = "Facturación"
        Case 2
            Nom = "Tickets"
        Case 3
            Nom = "Pr. x Tickets"
        Case 4
            Nom = "Pax Viajados"
        Case 5
            Nom = "Pr x Pax Viaj."
        Case 6
            Nom = "Captación"
    End Select
    
    SprEZE.SetText 1, i, Nom
    
        L_RellenarDatos_EZE i, 2, Nom
        L_RellenarDatos_EZE i, 3, Nom
        L_RellenarDatos_EZE i, 6, Nom

Next
End Sub
Private Sub L_LlenarTotalesCIA_IF()
Dim Nom As String
Dim i As Integer

For i = 1 To 6
    SprCIA_IF.MaxRows = SprCIA_IF.MaxRows + 1
    
    Select Case i
        Case 1
            Nom = "Facturación"
        Case 2
            Nom = "Tickets"
        Case 3
            Nom = "Pr. x Tickets"
        Case 4
            Nom = "Pax Viajados"
        Case 5
            Nom = "Pr x Pax Viaj."
        Case 6
            Nom = "Captación"
    End Select
    
    SprCIA_IF.SetText 1, i, Nom
    
        L_RellenarDatos_CIA_IF i, 2, Nom
        L_RellenarDatos_CIA_IF i, 3, Nom
        L_RellenarDatos_CIA_IF i, 6, Nom

Next
End Sub

Private Sub L_ocultar(spr As control, CD As Integer, CH As Integer, FD As Integer, FH As Integer, valor As Boolean)
Dim i


For i = CD To CH
    spr.col = i
    spr.ColHidden = valor
    If Not valor Then
       If i = CH Then
        spr.ColWidth(i) = 4
       Else
        spr.ColWidth(i) = 12
       End If
    End If

Next
    
End Sub

Public Sub L_PonerItems()

SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Facturación"
SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Tickets"
SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Pr. x Tickets"
SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Pax Viajados"
SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Pr x Pax Viaj."
SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Captación"

SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Facturación"
SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Tickets"
SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Pr. x Tickets"
SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Pax Viajados"
SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Pr x Pax Viaj."
SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Captación"

SprEzeAS.MaxRows = SprEzeAS.MaxRows + 1
SprEzeAS.SetText 1, SprEzeAS.MaxRows, "Facturación"
SprEzeAS.MaxRows = SprEzeAS.MaxRows + 1
SprEzeAS.SetText 1, SprEzeAS.MaxRows, "Tickets"
SprEzeAS.MaxRows = SprEzeAS.MaxRows + 1
SprEzeAS.SetText 1, SprEzeAS.MaxRows, "Pr. x Tickets"
SprEzeAS.MaxRows = SprEzeAS.MaxRows + 1
SprEzeAS.SetText 1, SprEzeAS.MaxRows, "Pax Viajados"
SprEzeAS.MaxRows = SprEzeAS.MaxRows + 1
SprEzeAS.SetText 1, SprEzeAS.MaxRows, "Pr x Pax Viaj."
SprEzeAS.MaxRows = SprEzeAS.MaxRows + 1
SprEzeAS.SetText 1, SprEzeAS.MaxRows, "Captación"

SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Facturación"
SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Tickets"
SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Pr. x Tickets"
SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Pax Viajados"
SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Pr x Pax Viaj."
SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Captación"

SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Facturación"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Tickets"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Pr. x Tickets"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Pax Viajados"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Pr x Pax Viaj."
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Captación"

SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, SprFlight.MaxRows, "Facturación"
SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, SprFlight.MaxRows, "Tickets"
SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, SprFlight.MaxRows, "Pr. x Tickets"
SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, SprFlight.MaxRows, "Pax Viajados"
SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, SprFlight.MaxRows, "Pr x Pax Viaj."
SprFlight.MaxRows = SprFlight.MaxRows + 1
SprFlight.SetText 1, SprFlight.MaxRows, "Captación"

End Sub

Private Sub L_Refrescar()
Dim SQL As String
Dim sqlX As String
Dim rs As Recordset

On Error GoTo ErrInd:

frmIndic_t.caption = Aplicacion.SeteoProceso(frmIndic_t.caption)


SQL = " SELECT "
SQL = SQL & " s.cod_depn, "
SQL = SQL & " decode(s.cod_depn,'INT','INT','IFL','IFL',cod_Ssdep) cod_sdep, "
SQL = SQL & " nvl(sum(cant_tickets),0) cant_t, "
SQL = SQL & " nvl(sum(importe),0) imp "
SQL = SQL & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)))
SQL = SQL & " S ,VENTAS.APERTURA_SDEP A "
SQL = SQL & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)))
SQL = SQL & "group by s.cod_depn,decode(s.cod_depn,'INT','INT','IFL','IFL',cod_Ssdep) "
SQL = SQL & "  "
SQL = SQL & " union "
SQL = SQL & " SELECT 'IFL','IFL',0,nvl(sum(importe_total),0) imp "
SQL = SQL & " FROM VENTAS.FL_VENTAS "
SQL = SQL & L_ArmarcondicionIFL(Year(CDate(mskFDesde.FormattedText)))


sqlX = " SELECT "
sqlX = sqlX & " s.cod_depn, "
sqlX = sqlX & " decode(s.cod_depn,'INT','INT','IFL','IFL',cod_ssdep) cod_sdep, "
sqlX = sqlX & " sum(cant_tickets) cant_t, "
sqlX = sqlX & " sum(importe) imp "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)) - 1)
sqlX = sqlX & " S ,VENTAS.APERTURA_SDEP A "
sqlX = sqlX & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)) - 1)
sqlX = sqlX & "group by s.cod_depn,decode(s.cod_depn,'INT','INT','IFL','IFL',cod_ssdep) "
sqlX = sqlX & " order by s.cod_depn,decode(s.cod_depn,'INT','INT','IFL','IFL',cod_ssdep)"


If Aplicacion.ObtenerRsDAO(SQL, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        
        L_VentaNeta
              
        L_PonerItems
        L_DecoEspigon
        If Aplicacion.ObtenerRsDAO(sqlX, RsDataAnt) Then
            L_DecoEspigonAnt
        End If
        L_TratarVolados
        L_TratarVoladosHIST
        L_TratarEstimado
        L_LlenarTotales
        L_LlenarTotalesBsAs
        L_LlenarTotalesEZEIZA
        L_LlenarTotalesCIA_IF
        L_Resaltar
        
        Aplicacion.CerrarDAO RsDataAnt
    End If
    
    Aplicacion.CerrarDAO RsData

End If

ErrInd:
    frmIndic_t.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub
Private Sub L_LlenarInFlight()
Dim SqlIF As String
Dim SqlIFLX As String
Dim rsIFL As Recordset
Dim rsIFLX As Recordset

On Error GoTo ErrInd:

SqlIF = " SELECT "
SqlIF = SqlIF & " sum(importe_total) imp "
SqlIF = SqlIF & " FROM VENTAS.FL_VENTAS " '''& funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)))
SqlIF = SqlIF & L_ArmarcondicionIFL(Year(CDate(mskFDesde.FormattedText)))
'SqlIF = SqlIF & "group by s.cod_depn,decode(s.cod_depn,'INT','INT','IFL','IFL',cod_Ssdep) "
'SqlIF = SqlIF & " order by s.cod_depn,decode(s.cod_depn,'INT','INT','IFL','IFL',cod_Ssdep) "


SqlIFLX = " SELECT "
SqlIFLX = SqlIFLX & " sum(importe_total) imp "
SqlIFLX = SqlIFLX & " FROM VENTAS.FL_VENTAS "
SqlIFLX = SqlIFLX & L_ArmarcondicionIFL(Year(CDate(mskFDesde.FormattedText)) - 1)


If Aplicacion.ObtenerRsDAO(SqlIF, rsIFL) Then
    
    If Aplicacion.CantReg(rsIFL) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        
        'L_VentaNeta
              
        ' L_PonerItemsIFL
        'L_DecoEspigon
        If Aplicacion.ObtenerRsDAO(SqlIFLX, rsIFLX) Then
         '   L_DecoEspigonAnt
        End If
        L_TratarVolados
        L_TratarVoladosHIST
        'L_TratarEstimado
        'L_LlenarInFlight
        'L_LlenarTotales
        'L_LlenarTotalesBsAs
        'L_LlenarTotalesEZEIZA
        L_Resaltar
        
        Aplicacion.CerrarDAO RsDataAnt
    End If
    
    Aplicacion.CerrarDAO RsData

End If

ErrInd:
    frmIndic_t.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub
Private Sub L_RellenarDatos_CIA_IF(i As Integer, col As Integer, Nom As String)
Dim valor As Variant, divsor As Variant
Dim Impr As Double
    
    If i = 1 Or i = 2 Then
        SprEzeA.GetText col, i, valor
        Impr = Impr + valor
        SprEzeAS.GetText col, i, valor
        Impr = Impr + valor
        SprEzeB.GetText col, i, valor
        Impr = Impr + valor
        SprAep.GetText col, i, valor
        Impr = Impr + valor
        SprInt.GetText col, i, valor
        Impr = Impr + valor
        SprFlight.GetText col, i, valor
        Impr = Impr + valor
    ElseIf i = 4 Then
        SprEzeA.GetText col, i, valor
        Impr = Impr + valor
        SprEzeAS.GetText col, i, valor
        Impr = Impr + valor
        SprEzeB.GetText col, i, valor
        Impr = Impr + valor
        SprAep.GetText col, i, valor
        Impr = Impr + valor
        SprInt.GetText col, i, valor
        Impr = Impr + valor
    ElseIf i = 3 Then
        SprCIA_IF.GetText col, 1, valor
        SprCIA_IF.GetText col, 2, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 5 Then
        SprCIA_IF.GetText col, 1, valor
        SprCIA_IF.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 6 Then
        SprCIA_IF.GetText col, 2, valor
        SprCIA_IF.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor * 100
        End If
    
    End If
    
    SprCIA_IF.SetText col, i, Format$((Impr), "#.00")

End Sub
Private Sub L_RellenarDatos(i As Integer, col As Integer, Nom As String)
Dim valor As Variant, divsor As Variant
Dim Impr As Double
    
    If i = 1 Or i = 2 Then
        SprEzeA.GetText col, i, valor
        Impr = Impr + valor
        SprEzeAS.GetText col, i, valor
        Impr = Impr + valor
        SprEzeB.GetText col, i, valor
        Impr = Impr + valor
        SprAep.GetText col, i, valor
        Impr = Impr + valor
        SprInt.GetText col, i, valor
        Impr = Impr + valor
    '    SprFlight.GetText col, i, valor
    '    Impr = Impr + valor
    ElseIf i = 4 Then
        SprEzeA.GetText col, i, valor
        Impr = Impr + valor
        SprEzeAS.GetText col, i, valor
        Impr = Impr + valor
        SprEzeB.GetText col, i, valor
        Impr = Impr + valor
        SprAep.GetText col, i, valor
        Impr = Impr + valor
        SprInt.GetText col, i, valor
        Impr = Impr + valor
    ElseIf i = 3 Then
        SprTotal.GetText col, 1, valor
        SprTotal.GetText col, 2, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 5 Then
        SprTotal.GetText col, 1, valor
        SprTotal.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 6 Then
        SprTotal.GetText col, 2, valor
        SprTotal.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor * 100
        End If
    
    End If
    
    SprTotal.SetText col, i, Format$((Impr), "#.00")

End Sub


Private Sub L_RellenarDatos_BsAs(i As Integer, col As Integer, Nom As String)
Dim valor As Variant, divsor As Variant
Dim Impr As Double
    
    If i = 1 Or i = 2 Then
        SprEzeA.GetText col, i, valor
        Impr = Impr + valor
        SprEzeAS.GetText col, i, valor
        Impr = Impr + valor
        SprEzeB.GetText col, i, valor
        Impr = Impr + valor
        SprAep.GetText col, i, valor
        Impr = Impr + valor
    ElseIf i = 4 Then
        SprEzeA.GetText col, i, valor
        Impr = Impr + valor
        SprEzeAS.GetText col, i, valor
        Impr = Impr + valor
        SprEzeB.GetText col, i, valor
        Impr = Impr + valor
        SprAep.GetText col, i, valor
        Impr = Impr + valor
    ElseIf i = 3 Then
        SprTotalBsAs.GetText col, 1, valor
        SprTotalBsAs.GetText col, 2, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 5 Then
        SprTotalBsAs.GetText col, 1, valor
        SprTotalBsAs.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 6 Then
        SprTotalBsAs.GetText col, 2, valor
        SprTotalBsAs.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor * 100
        End If
    
    End If
    
    SprTotalBsAs.SetText col, i, Format$((Impr), "#.00")

End Sub
Private Sub L_RellenarDatos_EZE(i As Integer, col As Integer, Nom As String)
Dim valor As Variant, divsor As Variant
Dim Impr As Double
    
    If i = 1 Or i = 2 Or i = 4 Then
        SprEzeA.GetText col, i, valor
        Impr = Impr + valor
        SprEzeAS.GetText col, i, valor
        Impr = Impr + valor
        SprEzeB.GetText col, i, valor
        Impr = Impr + valor
    ElseIf i = 3 Then
        SprEZE.GetText col, 1, valor
        SprEZE.GetText col, 2, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 5 Then
        SprEZE.GetText col, 1, valor
        SprEZE.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 6 Then
        SprEZE.GetText col, 2, valor
        SprEZE.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor * 100
        End If
    
    End If
    
    SprEZE.SetText col, i, Format$((Impr), "#.00")

End Sub

Private Sub L_Resaltar()
Dim fila As Integer

For fila = 1 To SprTotal.MaxRows
    Spread.spread_ResaltarCelda SprTotal, 4, fila
    Spread.spread_ResaltarCelda SprTotal, 5, fila
    Spread.spread_ResaltarCelda SprTotal, 7, fila
    Spread.spread_ResaltarCelda SprTotal, 8, fila
    
    Spread.spread_ResaltarCelda SprEzeA, 4, fila
    Spread.spread_ResaltarCelda SprEzeA, 5, fila
    Spread.spread_ResaltarCelda SprEzeA, 7, fila
    Spread.spread_ResaltarCelda SprEzeA, 8, fila

    Spread.spread_ResaltarCelda SprEzeAS, 4, fila
    Spread.spread_ResaltarCelda SprEzeAS, 5, fila
    Spread.spread_ResaltarCelda SprEzeAS, 7, fila
    Spread.spread_ResaltarCelda SprEzeAS, 8, fila

    Spread.spread_ResaltarCelda SprEzeB, 4, fila
    Spread.spread_ResaltarCelda SprEzeB, 5, fila
    Spread.spread_ResaltarCelda SprEzeB, 7, fila
    Spread.spread_ResaltarCelda SprEzeB, 8, fila

    Spread.spread_ResaltarCelda SprAep, 4, fila
    Spread.spread_ResaltarCelda SprAep, 5, fila
    Spread.spread_ResaltarCelda SprAep, 7, fila
    Spread.spread_ResaltarCelda SprAep, 8, fila

    Spread.spread_ResaltarCelda SprInt, 4, fila
    Spread.spread_ResaltarCelda SprInt, 5, fila
    Spread.spread_ResaltarCelda SprInt, 7, fila
    Spread.spread_ResaltarCelda SprInt, 8, fila

    Spread.spread_ResaltarCelda SprTotalBsAs, 4, fila
    Spread.spread_ResaltarCelda SprTotalBsAs, 5, fila
    Spread.spread_ResaltarCelda SprTotalBsAs, 7, fila
    Spread.spread_ResaltarCelda SprTotalBsAs, 8, fila
    
    Spread.spread_ResaltarCelda SprFlight, 4, fila
    Spread.spread_ResaltarCelda SprFlight, 5, fila
    Spread.spread_ResaltarCelda SprFlight, 7, fila
    Spread.spread_ResaltarCelda SprFlight, 8, fila
    
    Spread.spread_ResaltarCelda SprEZE, 4, fila
    Spread.spread_ResaltarCelda SprEZE, 5, fila
    Spread.spread_ResaltarCelda SprEZE, 7, fila
    Spread.spread_ResaltarCelda SprEZE, 8, fila
    
    Spread.spread_ResaltarCelda SprCIA_IF, 4, fila
    Spread.spread_ResaltarCelda SprCIA_IF, 5, fila
    Spread.spread_ResaltarCelda SprCIA_IF, 7, fila
    Spread.spread_ResaltarCelda SprCIA_IF, 8, fila

Next

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
'        Case FLIGHT
'           UsoPorFlight
        Case total
            UsoPorTotal = True
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
'        Case FLIGHT
'            UsoPorFlight = False
        Case INTE
            UsoPorInte = True
        Case total
            UsoPorTotal = False
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorAero Or UsoPorTotal Or UsoPorInte Or UsoPorFlight) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub


Private Sub L_TratarEstimado()
Dim SQL As String
Dim Tipo As String
Dim anio As Integer, Mes As Integer
Dim FD As String, FH As String

Aplicacion.ComienzoTrans

If optFechas(0).Value Then
    Tipo = "A"
    anio = Year(mskFDesde.FormattedText)
    Mes = Month(mskFDesde.FormattedText)
    FD = "01-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText)
    FH = mskFDesde.FormattedText
Else
    Tipo = "M"
    anio = Year(mskFDesde.FormattedText)
    Mes = Month(mskFDesde.FormattedText)
    FD = mskFDesde.FormattedText
    FH = mskFHasta.FormattedText
End If

SQL = "BEGIN estadis.Estimado_acum ( "
SQL = SQL & "'" & Tipo & "',"
SQL = SQL & anio & ","
SQL = SQL & Mes & ","
SQL = SQL & "" & Func.func_ToDate(FD) & ","
SQL = SQL & "" & Func.func_ToDate(FH) & " ); "
SQL = SQL & " End;"

If Aplicacion.EjecutarDAO(SQL) Then
    SQL = "SELECT cod_depn,"
    SQL = SQL & " decode(cod_depn,'INT','INT',cod_sdep) cod_sdep,"
    SQL = SQL & " tipo_porc,"
    SQL = SQL & " sum(valor) valor "
    SQL = SQL & " FROM estadis.resultados_estim  "
    SQL = SQL & " group by cod_depn,decode(cod_depn,'INT','INT',cod_sdep),tipo_porc"
    
    If Aplicacion.ObtenerRsDAO(SQL, RsDataEstim) Then
        L_DecoEstim
    End If
    Aplicacion.TerminarConExitoTrans
    Aplicacion.CerrarDAO RsDataEstim
Else
    Aplicacion.TerminarConErrorTrans
End If

End Sub

Private Sub L_TratarVolados()

Dim SQL As String
Dim SqlIFL As String
Dim FD As String, FH As String
Dim rs As Recordset

If optFechas(0).Value Then
    FD = "01-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText)
    FH = mskFDesde.FormattedText
Else
    FD = mskFDesde.FormattedText
    FH = mskFHasta.FormattedText
End If

L_ControlFecha


    If Control_BsAs Then
    SQL = " SELECT "
    SQL = SQL & "P.cod_depn,"
    SQL = SQL & "decode(P.cod_depn,'INT','INT',cod_ssdep) cod_sdep,"
'    SQL = SQL & " SUM(pax_AT_VS_vol.PAX_VOL) CANT_V "
    SQL = SQL & " SUM(cantidad) CANT_V "
    SQL = SQL & " From estadis.pax_volados P , ventas.apertura_sdep A"
    SQL = SQL & " Where P.local = A.cod_local and p.cod_depn = a.cod_depn and p.cod_sdep = a.cod_sdep  And "
    SQL = SQL & " fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
    SQL = SQL & " GROUP BY p.COD_DEPN,decode(P.cod_depn,'INT','INT',cod_ssdep)"
    
'    SqlIFL = " SELECT "
'    SqlIFL = SqlIFL & "P.cod_depn,"
'    SqlIFL = SqlIFL & "decode(P.cod_depn,'IFL','IFL',cod_ssdep) cod_sdep,"
'    SqlIFL = SqlIFL & " SUM(cantidad) CANT_V "
'    SqlIFL = SqlIFL & " From estadis.pax_volados P , ventas.apertura_sdep A"
'    SqlIFL = SqlIFL & " Where P.local = A.cod_local and p.cod_depn = a.cod_depn and p.cod_sdep = a.cod_sdep  And "
'    SqlIFL = SqlIFL & " fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
'    SqlIFL = SqlIFL & " and P.COD_CIA_AEREA = 4  AND P.COD_DEPN = 'IFL' "
'    SqlIFL = SqlIFL & " GROUP BY p.COD_DEPN,decode(P.cod_depn,'IFL','IFL',cod_ssdep)"
        
        SqlIFL = " SELECT SUM(CANTIDAD) CANT_V "
        SqlIFL = SqlIFL & " FROM estadis.pax_volados P "
        SqlIFL = SqlIFL & " WHERE "
        SqlIFL = SqlIFL & " (P.FCH_VUELO ,P.COD_VUELO) IN (SELECT DISTINCT FCH_VUELO,NRO_VUELO FROM VENTAS.FL_TRIPULACION) "
        SqlIFL = SqlIFL & " AND P.COD_CIA_AEREA = 4 "
        SqlIFL = SqlIFL & " AND P.COD_DEPN = 'EZE' "
        SqlIFL = SqlIFL & " AND fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
    
        If Aplicacion.ObtenerRsDAO(SqlIFL, rs) Then
                If Aplicacion.CantReg(rs) > 0 Then
                      Vol_Flight = rs!cant_v
                   Else
                      Vol_Flight = 0
                End If
           Else
              Vol_Flight = 0
        End If
    
      If Aplicacion.ObtenerRsDAO(SQL, RsDataVol) Then
          L_DecoVolados 2
          Aplicacion.CerrarDAO RsDataVol
      End If
        
    Else
    
        labInfo.caption = "Datos de pasajeros transitados Actualizados al día " & Fch_BsAs
        Timer1.Enabled = True
    End If

End Sub

Private Sub L_TratarVoladosHIST()
Dim SQL As String
Dim SqlIFL As String
Dim FD As String, FH As String
Dim rs As Recordset

If optFechas(0).Value Then
    FD = "01-01-" & Year(mskFDesde.FormattedText) - 1
    FH = Day(mskFDesde.FormattedText) & "-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText) - 1
Else
    FD = mskFDesdeAnt.FormattedText
    FH = mskFHastaAnt.FormattedText
End If

        SQL = " SELECT "
        SQL = SQL & "P.cod_depn,"
        SQL = SQL & "decode(P.cod_depn,'INT','INT',cod_ssdep) cod_sdep,"
        SQL = SQL & " SUM(p.cantidad) CANT_V "
        SQL = SQL & " From estadis.pax_volados P, ventas.apertura_sdep A "
        SQL = SQL & " Where P.local = A.cod_local and p.cod_depn = a.cod_depn and p.cod_sdep = a.cod_sdep  And"
        SQL = SQL & " fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
        SQL = SQL & " GROUP BY P.COD_DEPN,decode(P.cod_depn,'INT','INT',cod_ssdep) "
        
        SqlIFL = " SELECT NVL(SUM(CANTIDAD),0) CANT_V "
        SqlIFL = SqlIFL & " FROM estadis.pax_volados P "
        SqlIFL = SqlIFL & " WHERE "
        SqlIFL = SqlIFL & " (P.FCH_VUELO ,P.COD_VUELO) IN (SELECT DISTINCT FCH_VUELO,NRO_VUELO FROM VENTAS.FL_TRIPULACION) "
        SqlIFL = SqlIFL & " AND P.COD_CIA_AEREA = 4 "
        SqlIFL = SqlIFL & " AND P.COD_DEPN = 'EZE' "
        SqlIFL = SqlIFL & " AND fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
        
        'If Aplicacion.ObtenerRsDAO(SqlIFL, rs) Then
        '        If Aplicacion.CantReg(rs) > 0 Then
        '              Vol_Flight_Ant = rs!cant_v
        '           Else
                      Vol_Flight = 0
        '        End If
        '    Else
        '       Vol_Flight_Ant = 0
        'End If
            
        
        If Aplicacion.ObtenerRsDAO(SQL, RsDataVol) Then
            L_DecoVolados 3
            Aplicacion.CerrarDAO RsDataVol
        End If

End Sub


Private Sub L_VentaNeta()
Dim SQL As String


SQL = " select Vta.cod_depn,Vta.cod_sdep,impDC,impVta "
SQL = SQL & " From "
SQL = SQL & " (select P.cod_depn,cod_ssdep cod_sdep ,sum(P.importe) impDC "
SQL = SQL & " from " & funcLocal_Vista("pago_lpt", Year(CDate(mskFDesde.FormattedText)))
SQL = SQL & " P , ventas.apertura_sdep A "
SQL = SQL & " Where "
SQL = SQL & " p.fch_ticket Between " & Func.func_ToDate(mskFDesde.FormattedText) & " and " & Func.func_ToDate(mskFHasta.FormattedText)
SQL = SQL & " and p.cod_depn=A.cod_depn and P.cod_sdep=A.cod_sdep and P.cod_local=A.cod_local"
SQL = SQL & " and tipo_pago = 14"
SQL = SQL & " group by P.cod_depn,cod_ssdep ) Pag,"
SQL = SQL & " (select V.cod_depn,cod_ssdep cod_sdep ,sum(V.importe) impVta"
SQL = SQL & " from " & funcLocal_Vista("venta_lgi", Year(CDate(mskFDesde.FormattedText)))
SQL = SQL & " V , ventas.apertura_sdep A"
SQL = SQL & " Where"
SQL = SQL & " V.fch_ticket Between " & Func.func_ToDate(mskFDesde.FormattedText) & " And " & Func.func_ToDate(mskFHasta.FormattedText)
SQL = SQL & " and V.cod_depn=A.cod_depn and V.cod_sdep=A.cod_sdep and V.cod_local=A.cod_local"
SQL = SQL & " and comitente = 'T'"
SQL = SQL & " group by V.cod_depn,cod_ssdep ) Vta "
SQL = SQL & " Where"
SQL = SQL & " Pag.cod_depn = Vta.cod_depn And Pag.cod_sdep = Vta.cod_sdep "
        
If Not Aplicacion.ObtenerRsDAO(SQL, RsActual) Then
   MsgBox "No se pudo cargar datos de Duty Check ", vbOKOnly + vbInformation, "Atencion"
End If
        
SQL = " select Vta.cod_depn,Vta.cod_sdep,impDC,impVta "
SQL = SQL & " From "
SQL = SQL & " (select P.cod_depn,cod_ssdep cod_sdep ,sum(P.importe) impDC "
SQL = SQL & " from " & funcLocal_Vista("pago_lpt", Year(CDate(mskFDesdeAnt.FormattedText)))
SQL = SQL & " P , ventas.apertura_sdep A "
SQL = SQL & " Where "
SQL = SQL & " p.fch_ticket Between " & Func.func_ToDate(mskFDesdeAnt.FormattedText) & " and " & Func.func_ToDate(mskFHastaAnt.FormattedText)
SQL = SQL & " and p.cod_depn=A.cod_depn and P.cod_sdep=A.cod_sdep and P.cod_local=A.cod_local"
SQL = SQL & " and tipo_pago = 14 "
SQL = SQL & " group by P.cod_depn,cod_ssdep ) Pag,"
SQL = SQL & " (select V.cod_depn,cod_ssdep cod_sdep ,sum(V.importe) impVta"
SQL = SQL & " from " & funcLocal_Vista("venta_lgi", Year(CDate(mskFDesdeAnt.FormattedText)))
SQL = SQL & " V , ventas.apertura_sdep A"
SQL = SQL & " Where"
SQL = SQL & " V.fch_ticket Between " & Func.func_ToDate(mskFDesdeAnt.FormattedText) & " And " & Func.func_ToDate(mskFHastaAnt.FormattedText)
SQL = SQL & " and V.cod_depn=A.cod_depn and V.cod_sdep=A.cod_sdep and V.cod_local=A.cod_local"
SQL = SQL & " and comitente = 'T'"
SQL = SQL & " group by V.cod_depn,cod_ssdep ) Vta "
SQL = SQL & " Where"
SQL = SQL & " Pag.cod_depn = Vta.cod_depn And Pag.cod_sdep = Vta.cod_sdep "
        
If Not Aplicacion.ObtenerRsDAO(SQL, RsAnterior) Then
   MsgBox "No se pudo cargar datos de Duty Check ", vbOKOnly + vbInformation, "Atencion"
End If
        
End Sub

Private Sub borGrBar_Click(Index As Integer)
Dim col_datos As Collection

    Set col_datos = New Collection
    Select Case Index
        Case 0
            L_LLenarColeccion col_datos, SprTotalBsAs
            FrmGrafbar.CargarGrafico total, "", col_datos
        Case 1
            L_LLenarColeccion col_datos, SprEzeA
            FrmGrafbar.CargarGrafico EZEA, "", col_datos
        Case 2
            L_LLenarColeccion col_datos, SprEzeB
            FrmGrafbar.CargarGrafico EZEB, "", col_datos
        Case 3
            L_LLenarColeccion col_datos, SprAep
            FrmGrafbar.CargarGrafico AERO, "", col_datos
        Case 4
            L_LLenarColeccion col_datos, SprInt
            FrmGrafbar.CargarGrafico INTE, "", col_datos
        Case 5
            L_LLenarColeccion col_datos, SprTotal
            FrmGrafbar.CargarGrafico total, "", col_datos
        Case 6
            'L_LLenarColeccion col_datos, SprFlight
            'FrmGrafbar.CargarGrafico FLIGHT, "", col_datos
    End Select
        
    
End Sub

Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        L_Refrescar
    Case 1
        If chkNeto.Value <> 0 Then
            chkNeto.Value = 0
        End If
        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        tabEspigon.Enabled = False
        L_LimpiarGrillas
        
    Case 2
        Unload Me
End Select
End Sub



Private Sub botExcel_Click(Index As Integer)
    If optFechas(0).Value Then
        L_TratarExcel SprTotal, "INDICADORES ACUMULADOS ANUALES (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
        "del 1 de Ene al " & Format$(mskFDesde.FormattedText, "dd-mmm"), "", Exl_Blanco
    Else
        L_TratarExcel SprTotal, "INDICADORES ACUMULADOS MES CORRIENTE (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
        " del " & Format$(mskFDesde.FormattedText, "dd-mmm") & " al " & Format$(mskFHasta.FormattedText, "dd-mmm"), "", Exl_Blanco 'Exl_Ros '
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

Private Sub chkNeto_Click()
Dim ocultar As Boolean

AjustarVenta chkNeto.Value
L_Resaltar
If chkNeto.Value = 0 Then
   ocultar = False
Else
   ocultar = True
End If
L_ocultar SprTotalBsAs, 6, 9, 1, SprTotalBsAs.MaxRows, ocultar
L_ocultar SprTotal, 6, 9, 1, SprTotal.MaxRows, ocultar
L_ocultar SprInt, 6, 9, 1, SprInt.MaxRows, ocultar
L_ocultar SprEzeA, 6, 9, 1, SprEzeA.MaxRows, ocultar
L_ocultar SprEzeB, 6, 9, 1, SprEzeB.MaxRows, ocultar
L_ocultar SprEzeAS, 6, 9, 1, SprEzeAS.MaxRows, ocultar
L_ocultar SprAep, 6, 9, 1, SprAep.MaxRows, ocultar
L_ocultar SprFlight, 6, 9, 1, SprFlight.MaxRows, ocultar
L_ocultar SprEZE, 6, 9, 1, SprEZE.MaxRows, ocultar
L_ocultar SprCIA_IF, 6, 9, 1, SprCIA_IF.MaxRows, ocultar

End Sub



Private Sub Form_Activate()
'tabEspigon.TabEnabled(7) = False
End Sub

Private Sub L_LLenarColeccion(ByRef col As Collection, spr As control)
Dim cl_dato As CLlgi
Dim valor As Variant
Dim item As Variant

spr.Row = spr.ActiveRow
spr.GetText 1, spr.Row, item

Set cl_dato = New CLlgi
    
spr.GetText 2, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
col.Add cl_dato
    
Set cl_dato = New CLlgi
spr.GetText 3, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
col.Add cl_dato

Set cl_dato = New CLlgi
spr.GetText 6, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
col.Add cl_dato

End Sub

Private Sub L_TratarExcel(spr As control, Titulo As String, subTit As String, Info As String, color As Integer)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim fila As Integer
Dim i
Dim tit As Variant
Dim nombre As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

frmIndic_t.caption = Aplicacion.SeteoProceso(frmIndic_t.caption)

If nombre <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.Application.Visible = True
    
    ReDim titCol(spr.MaxCols)
    col = 1
    fila = 3
    
    For i = 1 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 1, 1, Titulo
    rango = Exl_rangos(1, 1, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, color
    
    Exl_PonerValor AppExcel, fila, col, subTit
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    fila = fila + 1
    
    Exl_PonerValor AppExcel, fila, col, "Total Companía"
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    AppExcel.Application.Range(rango).Merge
        
    fila = fila + 1
    
    Exl_BajarGrillaExel spr, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + spr.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '-------------------------
    fila = fila + 1
    
    Exl_PonerValor AppExcel, fila, col, "Total Buenos Aires"
    rango = Exl_rangos(fila, fila, 1, SprTotalBsAs.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    AppExcel.Application.Range(rango).Merge
        
    fila = fila + 1
    
    Exl_BajarGrillaExel SprTotalBsAs, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprTotalBsAs.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "INTERNACIONAL A Llegadas"
    rango = Exl_rangos(fila, fila, 1, SprEzeA.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprEzeA, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprEzeA.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "INTERNACIONAL A Salidas "
    rango = Exl_rangos(fila, fila, 1, SprEzeAS.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprEzeAS, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprEzeAS.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
        
    '-----------------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "INTERNACIONAL B "
    rango = Exl_rangos(fila, fila, 1, SprEzeB.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba
    
    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprEzeB, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprEzeB.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '--------------------------------
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "AEROPARQUE "
    rango = Exl_rangos(fila, fila, 1, SprAep.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba
    
    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprAep, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprAep.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
        
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris

    '------------------------
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "INTERIOR "
    rango = Exl_rangos(fila, fila, 1, SprInt.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba
    
    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprInt, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprInt.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
        
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris

    '------------------------
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "IN FLIGHT "
    rango = Exl_rangos(fila, fila, 1, SprFlight.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprFlight, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprFlight.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '-----------------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, " TOTAL EZEIZA "
    rango = Exl_rangos(fila, fila, 1, SprEZE.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprEZE, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprEZE.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '-----------------------------------

    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.Zoom = False
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    AppExcel.SaveAs nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmIndic_t.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6300
Me.Width = 10300

If Day(Date) = 1 Then
    mskFDesde.Text = Format$(Func.func_Dia1SegunMes_Anio(Month(Date - 1), Year(Date)), FTOFECHA)
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
    
    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
    mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio, FTOFECHA)
    
Else
    mskFDesde.Text = "01-" & Format$(Month(Date), "0#") & "-" & Format$(Year(Date), "####")
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
    
    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
    mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio, FTOFECHA)
    
End If


L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmindic_t"

End Sub


Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmindic_t"
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
            SprTotal.SetFocus
        Case 1
            SprEzeA.SetFocus
        Case 2
            SprEzeAS.SetFocus
        Case 3
            SprEzeB.SetFocus
        Case 4
            SprAep.SetFocus
    End Select
    
    
ErrT:
    Exit Sub



End Sub

Private Sub Timer1_Timer()
    
If Not Control_BsAs Then
    labInfo.Visible = Not labInfo.Visible
 ElseIf Not Control_Int Then
     If tabEspigon.Tab > 4 Then
        labInfo.Visible = Not labInfo.Visible
     Else
        labInfo.Visible = False
     End If
End If


End Sub


