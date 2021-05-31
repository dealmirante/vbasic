VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIndic 
   Caption         =   "INDICADORES"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9195
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
      TabIndex        =   40
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
      Left            =   75
      TabIndex        =   2
      Top             =   1545
      Width           =   8950
      _ExtentX        =   15796
      _ExtentY        =   6720
      _Version        =   327680
      Tabs            =   7
      TabsPerRow      =   7
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
      TabCaption(0)   =   "TOT BS.AS."
      TabPicture(0)   =   "frmIndic.frx":0000
      Tab(0).ControlCount=   3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "borGrBar(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "botExcel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "sprTotalBsAs"
      Tab(0).Control(2).Enabled=   0   'False
      TabCaption(1)   =   "INTA-L"
      TabPicture(1)   =   "frmIndic.frx":001C
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sprEzeA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "borGrBar(1)"
      Tab(1).Control(1).Enabled=   -1  'True
      TabCaption(2)   =   "INTA-S"
      TabPicture(2)   =   "frmIndic.frx":0038
      Tab(2).ControlCount=   2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "sprEzeAS"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "borGrBar(6)"
      Tab(2).Control(1).Enabled=   -1  'True
      TabCaption(3)   =   "INTER B"
      TabPicture(3)   =   "frmIndic.frx":0054
      Tab(3).ControlCount=   2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "sprEzeB"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "borGrBar(2)"
      Tab(3).Control(1).Enabled=   -1  'True
      TabCaption(4)   =   "AEP"
      TabPicture(4)   =   "frmIndic.frx":0070
      Tab(4).ControlCount=   2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "borGrBar(3)"
      Tab(4).Control(0).Enabled=   -1  'True
      Tab(4).Control(1)=   "sprAep"
      Tab(4).Control(1).Enabled=   0   'False
      TabCaption(5)   =   "TOT CIA"
      TabPicture(5)   =   "frmIndic.frx":008C
      Tab(5).ControlCount=   3
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "sprTotal"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "borGrBar(5)"
      Tab(5).Control(1).Enabled=   -1  'True
      Tab(5).Control(2)=   "Command1"
      Tab(5).Control(2).Enabled=   -1  'True
      TabCaption(6)   =   "INTERIOR"
      TabPicture(6)   =   "frmIndic.frx":00A8
      Tab(6).ControlCount=   2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "borGrBar(4)"
      Tab(6).Control(0).Enabled=   -1  'True
      Tab(6).Control(1)=   "SprInt"
      Tab(6).Control(1).Enabled=   0   'False
      Begin FPSpread.vaSpread sprEzeB 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic.frx":00C4
         TabIndex        =   36
         Top             =   450
         Width           =   7695
      End
      Begin FPSpread.vaSpread sprAep 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic.frx":0805
         TabIndex        =   34
         Top             =   450
         Width           =   7695
      End
      Begin FPSpread.vaSpread sprTotal 
         Height          =   3105
         Left            =   -74775
         OleObjectBlob   =   "frmIndic.frx":0F58
         TabIndex        =   31
         Top             =   450
         Width           =   7695
      End
      Begin FPSpread.vaSpread SprInt 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic.frx":1699
         TabIndex        =   29
         Top             =   450
         Width           =   7695
      End
      Begin FPSpread.vaSpread sprTotalBsAs 
         Height          =   3105
         Left            =   195
         OleObjectBlob   =   "frmIndic.frx":1DEC
         TabIndex        =   26
         Top             =   450
         Width           =   7695
      End
      Begin FPSpread.vaSpread sprEzeA 
         Height          =   3100
         Left            =   -74800
         OleObjectBlob   =   "frmIndic.frx":252D
         TabIndex        =   14
         Top             =   450
         Width           =   7700
      End
      Begin FPSpread.vaSpread sprEzeAS 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic.frx":2C6E
         TabIndex        =   38
         Top             =   450
         Width           =   7695
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   6
         Left            =   -66975
         Picture         =   "frmIndic.frx":33AF
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3030
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   2
         Left            =   -66990
         Picture         =   "frmIndic.frx":37F1
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3030
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   3
         Left            =   -66990
         Picture         =   "frmIndic.frx":3C33
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3030
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   5
         Left            =   -66960
         Picture         =   "frmIndic.frx":4075
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2370
         Width           =   780
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -66960
         Picture         =   "frmIndic.frx":44B7
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3060
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   4
         Left            =   -66990
         Picture         =   "frmIndic.frx":4A49
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3045
         Width           =   780
      End
      Begin VB.CommandButton botExcel 
         Caption         =   "Excel"
         Height          =   510
         Left            =   8040
         Picture         =   "frmIndic.frx":4E8B
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3045
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   0
         Left            =   8040
         Picture         =   "frmIndic.frx":541D
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2385
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   1
         Left            =   -66975
         Picture         =   "frmIndic.frx":585F
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
         Picture         =   "frmIndic.frx":5CA1
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
         Picture         =   "frmIndic.frx":5DA3
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
         Picture         =   "frmIndic.frx":5EA5
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
         Caption         =   "Fechas de  Consultas para Comparaci�n"
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
            Caption         =   "Per�odo Anterior"
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
               Picture         =   "frmIndic.frx":66C7
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmIndic.frx":6839
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
               _Version        =   393216
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
            Caption         =   "Per�odo "
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
               Picture         =   "frmIndic.frx":69AB
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2550
               Picture         =   "frmIndic.frx":6B1D
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
      TabIndex        =   25
      Top             =   5400
      Width           =   8730
   End
End
Attribute VB_Name = "frmIndic"
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

Dim Control_BsAs As Boolean
Dim Control_Int As Boolean

Dim Fch_BsAs As String
Dim Fch_Int As String

Private Sub AjustarVenta(Tipo As Integer)
Dim valorTicket As Variant, valorPax As Variant

On Error GoTo err1:

RsActual.MoveFirst
Do While Not RsActual.EOF
    Select Case RsActual!Cod_Sdep
        Case "AEP"
            sprAep.GetText 2, 2, valorTicket
            sprAep.GetText 2, 4, valorPax
            If Tipo = 0 Then
             sprAep.SetText 2, 1, str(RsActual!impvta)
             sprAep.SetText 2, 3, str(RsActual!impvta / valorTicket)
             If valorPax > 0 Then
                sprAep.SetText 2, 5, str(RsActual!impvta / valorPax)
             End If
            Else
             sprAep.SetText 2, 1, str(RsActual!impvta - RsActual!impdc)
             sprAep.SetText 2, 3, str((RsActual!impvta - RsActual!impdc) / valorTicket)
             If valorPax > 0 Then
                sprAep.SetText 2, 5, str((RsActual!impvta - RsActual!impdc) / valorPax)
             End If
            End If
            
        Case "INTAL"
            sprEzeA.GetText 2, 2, valorTicket
            sprEzeA.GetText 2, 4, valorPax
            If Tipo = 0 Then
             sprEzeA.SetText 2, 1, str(RsActual!impvta)
             sprEzeA.SetText 2, 3, str(RsActual!impvta / valorTicket)
             If valorPax > 0 Then
                sprEzeA.SetText 2, 5, str(RsActual!impvta / valorPax)
             End If
            Else
             sprEzeA.SetText 2, 1, str(RsActual!impvta - RsActual!impdc)
             sprEzeA.SetText 2, 3, str((RsActual!impvta - RsActual!impdc) / valorTicket)
             If valorPax > 0 Then
                sprEzeA.SetText 2, 5, str((RsActual!impvta - RsActual!impdc) / valorPax)
             End If
            End If
        Case "INTAS"
            sprEzeAS.GetText 2, 2, valorTicket
            sprEzeAS.GetText 2, 4, valorPax
            If Tipo = 0 Then
             sprEzeAS.SetText 2, 1, str(RsActual!impvta)
             sprEzeAS.SetText 2, 3, str(RsActual!impvta / valorTicket)
             If valorPax > 0 Then
                sprEzeAS.SetText 2, 5, str(RsActual!impvta / valorPax)
             End If
            Else
             sprEzeAS.SetText 2, 1, str(RsActual!impvta - RsActual!impdc)
             sprEzeAS.SetText 2, 3, str((RsActual!impvta - RsActual!impdc) / valorTicket)
             If valorPax > 0 Then
                sprEzeAS.SetText 2, 5, str((RsActual!impvta - RsActual!impdc) / valorPax)
             End If
            End If
        Case "INTB"
            sprEzeB.GetText 2, 2, valorTicket
            sprEzeB.GetText 2, 4, valorPax
            If Tipo = 0 Then
             sprEzeB.SetText 2, 1, str(RsActual!impvta)
             sprEzeB.SetText 2, 3, str(RsActual!impvta / valorTicket)
             If valorPax > 0 Then
                sprEzeB.SetText 2, 5, str(RsActual!impvta / valorPax)
             End If
            Else
             sprEzeB.SetText 2, 1, str(RsActual!impvta - RsActual!impdc)
             sprEzeB.SetText 2, 3, str((RsActual!impvta - RsActual!impdc) / valorTicket)
             If valorPax > 0 Then
                sprEzeB.SetText 2, 5, str((RsActual!impvta - RsActual!impdc) / valorPax)
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
L_RellenarDatos_BsAs 1, 2, ""
L_RellenarDatos 3, 2, ""
L_RellenarDatos 5, 2, ""
L_RellenarDatos_BsAs 3, 2, ""
L_RellenarDatos_BsAs 5, 2, ""

err1:
  Resume err12
  
err12:
On Error GoTo err2:

RsAnterior.MoveFirst
Do While Not RsAnterior.EOF
    Select Case RsAnterior!Cod_Sdep
        Case "AEP"
            sprAep.GetText 3, 2, valorTicket
            sprAep.GetText 3, 4, valorPax
            If Tipo = 0 Then
             sprAep.SetText 3, 1, str(RsAnterior!impvta)
             sprAep.SetText 3, 3, str(RsAnterior!impvta / valorTicket)
             If valorPax > 0 Then
                sprAep.SetText 3, 5, str(RsAnterior!impvta / valorPax)
             End If
            Else
             sprAep.SetText 3, 1, str(RsAnterior!impvta - RsAnterior!impdc)
             sprAep.SetText 3, 3, str((RsAnterior!impvta - RsAnterior!impdc) / valorTicket)
             If valorPax > 0 Then
                sprAep.SetText 3, 5, str((RsAnterior!impvta - RsAnterior!impdc) / valorPax)
             End If
            End If
        Case "INTAL"
            sprEzeA.GetText 3, 2, valorTicket
            sprEzeA.GetText 3, 4, valorPax
            If Tipo = 0 Then
             sprEzeA.SetText 3, 1, str(RsAnterior!impvta)
             sprEzeA.SetText 3, 3, str(RsAnterior!impvta / valorTicket)
             If valorPax > 0 Then
                sprEzeA.SetText 3, 5, str(RsAnterior!impvta / valorPax)
             End If
            Else
             sprEzeA.SetText 3, 1, str(RsAnterior!impvta - RsAnterior!impdc)
             sprEzeA.SetText 3, 3, str((RsAnterior!impvta - RsAnterior!impdc) / valorTicket)
             If valorPax > 0 Then
                sprEzeA.SetText 3, 5, str((RsAnterior!impvta - RsAnterior!impdc) / valorPax)
             End If
            End If
        Case "INTAS"
            sprEzeAS.GetText 3, 2, valorTicket
            sprEzeAS.GetText 3, 4, valorPax
            If Tipo = 0 Then
             sprEzeAS.SetText 3, 1, str(RsAnterior!impvta)
             sprEzeAS.SetText 3, 3, str(RsAnterior!impvta / valorTicket)
             If valorPax > 0 Then
                sprEzeAS.SetText 3, 5, str(RsAnterior!impvta / valorPax)
             End If
            Else
             sprEzeAS.SetText 3, 1, str(RsAnterior!impvta - RsAnterior!impdc)
             sprEzeAS.SetText 3, 3, str((RsAnterior!impvta - RsAnterior!impdc) / valorTicket)
             If valorPax > 0 Then
                sprEzeAS.SetText 3, 5, str((RsAnterior!impvta - RsAnterior!impdc) / valorPax)
             End If
            End If
        Case "INTB"
            sprEzeB.GetText 3, 2, valorTicket
            sprEzeB.GetText 3, 4, valorPax
            If Tipo = 0 Then
             sprEzeB.SetText 3, 1, str(RsAnterior!impvta)
             sprEzeB.SetText 3, 3, str(RsAnterior!impvta / valorTicket)
             If valorPax > 0 Then
                sprEzeB.SetText 3, 5, str(RsAnterior!impvta / valorPax)
             End If
            Else
             sprEzeB.SetText 3, 1, str(RsAnterior!impvta - RsAnterior!impdc)
             sprEzeB.SetText 3, 3, str((RsAnterior!impvta - RsAnterior!impdc) / valorTicket)
             If valorPax > 0 Then
                sprEzeB.SetText 3, 5, str((RsAnterior!impvta - RsAnterior!impdc) / valorPax)
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
L_RellenarDatos_BsAs 1, 3, ""
L_RellenarDatos 3, 3, ""
L_RellenarDatos 5, 3, ""
L_RellenarDatos_BsAs 3, 3, ""
L_RellenarDatos_BsAs 5, 3, ""

err2:
   Exit Sub
End Sub

Private Sub L_ControlFecha()
Dim FH As String
'Dim Fch_BsAs As String
'Dim Fch_Int As String
Dim rs_pax As Recordset
Dim sql As String

FH = mskFHasta.FormattedText

sql = "select min(fch_actualizado) Fch from estadis.control_carga Where Cod_sdep <> 'INT' "

If Aplicacion.ObtenerRsDAO(sql, rs_pax) Then
   Fch_BsAs = CDate(rs_pax!fch)
End If


sql = "select min(fch_actualizado) Fch from estadis.control_carga Where Cod_sdep = 'INT' "

If Aplicacion.ObtenerRsDAO(sql, rs_pax) Then
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

Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant

Do While Not RsData.EOF
        Select Case RsData!cod_depn
            Case DSLoc(1).Dep

                sprAep.SetText 2, 1, str(RsData!imp)
                
                sprAep.SetText 2, 2, str(RsData!cant_t)
                
                sprAep.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                                   
            
            Case DSLoc(2).Dep
                Select Case RsData!Cod_Sdep
                    Case DSLoc(2).Sdep

                        sprEzeA.SetText 2, 1, str(RsData!imp)
                        
                        sprEzeA.SetText 2, 2, str(RsData!cant_t)
                        
                        sprEzeA.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                        

                    Case DSLoc(3).Sdep
                        sprEzeB.SetText 2, 1, str(RsData!imp)
                        
                        sprEzeB.SetText 2, 2, str(RsData!cant_t)
                        
                        sprEzeB.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                    
                    Case DSLoc(9).Sdep

                        sprEzeA.SetText 2, 1, str(RsData!imp)
                        
                        sprEzeA.SetText 2, 2, str(RsData!cant_t)
                        
                        sprEzeA.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                        

                    Case DSLoc(10).Sdep
                        sprEzeAS.SetText 2, 1, str(RsData!imp)
                        
                        sprEzeAS.SetText 2, 2, str(RsData!cant_t)
                        
                        sprEzeAS.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                        
                
                End Select
                
            Case DSLoc(4).Dep

                        SprInt.SetText 2, 1, str(RsData!imp)
                        
                        SprInt.SetText 2, 2, str(RsData!cant_t)
                        
                        SprInt.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                        
                

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
                sprAep.SetText 3, 1, str(RsDataAnt!imp)
                sprAep.SetText 3, 2, str(RsDataAnt!cant_t)
                sprAep.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
            Case DSLoc(2).Dep
                Select Case RsDataAnt!Cod_Sdep
                    Case DSLoc(2).Sdep
                        sprEzeA.SetText 3, 1, str(RsDataAnt!imp)
                        sprEzeA.SetText 3, 2, str(RsDataAnt!cant_t)
                        sprEzeA.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                    Case DSLoc(3).Sdep
                        sprEzeB.SetText 3, 1, str(RsDataAnt!imp)
                        sprEzeB.SetText 3, 2, str(RsDataAnt!cant_t)
                        sprEzeB.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                    Case DSLoc(9).Sdep
                        sprEzeA.SetText 3, 1, str(RsDataAnt!imp)
                        sprEzeA.SetText 3, 2, str(RsDataAnt!cant_t)
                        sprEzeA.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                    Case DSLoc(10).Sdep
                        sprEzeAS.SetText 3, 1, str(RsDataAnt!imp)
                        sprEzeAS.SetText 3, 2, str(RsDataAnt!cant_t)
                        sprEzeAS.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                End Select
            Case DSLoc(4).Dep
                SprInt.SetText 3, 1, str(RsDataAnt!imp)
                SprInt.SetText 3, 2, str(RsDataAnt!cant_t)
                SprInt.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
             
        End Select
                
        RsDataAnt.MoveNext

Loop

End Sub
Private Sub L_DecoEstim()
Dim ImpAep As Double, ImpA As Double, ImpB As Double
Dim valor As Variant

Do While Not RsDataEstim.EOF
        Select Case RsDataEstim!cod_depn
            Case DSLoc(1).Dep
                Select Case RsDataEstim!tipo_porc
                    Case "I"
                        sprAep.SetText 6, 1, str(RsDataEstim!valor)
                        ImpAep = RsDataEstim!valor
                    Case "T"
                        sprAep.SetText 6, 2, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            sprAep.SetText 6, 3, str(ImpAep / RsDataEstim!valor)
                        End If
                        sprAep.GetText 6, 4, valor
                        If Val(valor) > 0 Then
                            sprAep.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                        End If
                    
                    Case "P"
                        sprAep.SetText 6, 4, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            sprAep.SetText 6, 5, str(ImpAep / RsDataEstim!valor)
                        End If
                
                End Select

            Case DSLoc(2).Dep
                Select Case RsDataEstim!Cod_Sdep
                    Case DSLoc(2).Sdep
                        Select Case RsDataEstim!tipo_porc
                            Case "I"
                                sprEzeA.SetText 6, 1, str(RsDataEstim!valor)
                                ImpA = RsDataEstim!valor
                            Case "T"
                                sprEzeA.SetText 6, 2, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    sprEzeA.SetText 6, 3, str(ImpA / RsDataEstim!valor)
                                End If
                                sprEzeA.GetText 6, 4, valor
                                If Val(valor) > 0 Then
                                    sprEzeA.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                                End If
                            
                            Case "P"
                                sprEzeA.SetText 6, 4, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    sprEzeA.SetText 6, 5, str(ImpA / RsDataEstim!valor)
                                End If
                            
                        End Select
                        
                    Case DSLoc(3).Sdep
                            Select Case RsDataEstim!tipo_porc
                                Case "I"
                                    sprEzeB.SetText 6, 1, str(RsDataEstim!valor)
                                    ImpB = RsDataEstim!valor
                                Case "T"
                                    sprEzeB.SetText 6, 2, str(RsDataEstim!valor)
                                    If RsDataEstim!valor > 0 Then
                                        sprEzeB.SetText 6, 3, str(ImpB / RsDataEstim!valor)
                                    End If
                                    sprEzeB.GetText 6, 4, valor
                                    If Val(valor) > 0 Then
                                        sprEzeB.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                                    End If
                                    
                                Case "P"
                                    sprEzeB.SetText 6, 4, str(RsDataEstim!valor)
                                    If RsDataEstim!valor > 0 Then
                                        sprEzeB.SetText 6, 5, str(ImpB / RsDataEstim!valor)
                                    End If
                            End Select
                    Case DSLoc(9).Sdep
                        Select Case RsDataEstim!tipo_porc
                            Case "I"
                                sprEzeA.SetText 6, 1, str(RsDataEstim!valor)
                                ImpA = RsDataEstim!valor
                            Case "T"
                                sprEzeA.SetText 6, 2, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    sprEzeA.SetText 6, 3, str(ImpA / RsDataEstim!valor)
                                End If
                                sprEzeA.GetText 6, 4, valor
                                If Val(valor) > 0 Then
                                    sprEzeA.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                                End If
                            
                            Case "P"
                                sprEzeA.SetText 6, 4, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    sprEzeA.SetText 6, 5, str(ImpA / RsDataEstim!valor)
                                End If
                            
                        End Select
                    Case DSLoc(10).Sdep
                        Select Case RsDataEstim!tipo_porc
                            Case "I"
                                sprEzeAS.SetText 6, 1, str(RsDataEstim!valor)
                                ImpA = RsDataEstim!valor
                            Case "T"
                                sprEzeAS.SetText 6, 2, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    sprEzeAS.SetText 6, 3, str(ImpA / RsDataEstim!valor)
                                End If
                                sprEzeAS.GetText 6, 4, valor
                                If Val(valor) > 0 Then
                                    sprEzeAS.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                                End If
                            
                            Case "P"
                                sprEzeAS.SetText 6, 4, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    sprEzeAS.SetText 6, 5, str(ImpA / RsDataEstim!valor)
                                End If
                            
                        End Select
                    
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


Private Sub L_DecoVolados(Col As Integer)
Dim valor As Variant

Do While Not RsDataVol.EOF
        Select Case RsDataVol!cod_depn
            Case DSLoc(1).Dep
                   sprAep.SetText Col, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      sprAep.GetText Col, 1, valor
                      sprAep.SetText Col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                      sprAep.GetText Col, 2, valor
                      sprAep.SetText Col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
            Case DSLoc(2).Dep
                Select Case RsDataVol!Cod_Sdep
                    Case DSLoc(2).Sdep
                        sprEzeA.SetText Col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            sprEzeA.GetText Col, 1, valor
                            sprEzeA.SetText Col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            sprEzeA.GetText Col, 2, valor
                            sprEzeA.SetText Col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(3).Sdep
                        sprEzeB.SetText Col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            sprEzeB.GetText Col, 1, valor
                            sprEzeB.SetText Col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            sprEzeB.GetText Col, 2, valor
                            sprEzeB.SetText Col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(9).Sdep
                        sprEzeA.SetText Col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            sprEzeA.GetText Col, 1, valor
                            sprEzeA.SetText Col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            sprEzeA.GetText Col, 2, valor
                            sprEzeA.SetText Col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(10).Sdep
                        sprEzeAS.SetText Col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            sprEzeAS.GetText Col, 1, valor
                            sprEzeAS.SetText Col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            sprEzeAS.GetText Col, 2, valor
                            sprEzeAS.SetText Col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                
                End Select
        
            Case DSLoc(4).Dep
               If (Control_Int And Col = 2) Or Col = 3 Then
                   SprInt.SetText Col, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprInt.GetText Col, 1, valor
                      SprInt.SetText Col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                      SprInt.GetText Col, 2, valor
                      SprInt.SetText Col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
                 Else
                   labInfo.caption = "Datos de pasajeros transitados del Interior al d�a " & Fch_Int
                   Timer1.Enabled = True
                   Exit Do
               End If
        End Select
                
        RsDataVol.MoveNext
Loop
End Sub

Private Sub L_DecoVoladosHIST()
Dim valor As Variant

Do While Not RsDataVol.EOF
        Select Case RsDataVol!cod_depn
            Case DSLoc(1).Dep
                   sprAep.SetText 3, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      sprAep.GetText 3, 1, valor
                      sprAep.SetText 3, 5, str(valor / RsDataVol!cant_v)
                      sprAep.GetText 3, 2, valor
                      sprAep.SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
            Case DSLoc(2).Dep
                Select Case RsDataVol!Cod_Sdep
                    Case DSLoc(2).Sdep
                        sprEzeA.SetText 3, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            sprEzeA.GetText 3, 1, valor
                            sprEzeA.SetText 3, 5, str(valor / RsDataVol!cant_v)
                            sprEzeA.GetText 3, 2, valor
                            sprEzeA.SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(3).Sdep
                        sprEzeB.SetText 2, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            sprEzeB.GetText 2, 1, valor
                            sprEzeB.SetText 2, 5, str(valor / RsDataVol!cant_v)
                            sprEzeB.GetText 2, 2, valor
                            sprEzeB.SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
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
End Sub


Private Sub L_LimpiarGrillas()
Dim i
    
    sprEzeA.MaxRows = 0
    sprEzeAS.MaxRows = 0
    sprEzeB.MaxRows = 0
    sprAep.MaxRows = 0
    SprInt.MaxRows = 0
    sprTotal.MaxRows = 0
    sprTotalBsAs.MaxRows = 0
    
    labInfo.caption = ""
    Timer1.Enabled = False
End Sub

Private Sub L_LlenarTotales()
Dim nom As String
Dim i As Integer

For i = 1 To 6
    sprTotal.MaxRows = sprTotal.MaxRows + 1
    
    Select Case i
        Case 1
            nom = "Facturaci�n"
        Case 2
            nom = "Tickets"
        Case 3
            nom = "Pr. x Tickets"
        Case 4
            nom = "Pax Viajados"
       Case 5
            nom = "Pr x Pax Viaj."
        Case 6
            nom = "Captaci�n"
    End Select
    
    sprTotal.SetText 1, i, nom
    
     If i < 4 Or Control_Int Then
        L_RellenarDatos i, 2, nom
     End If
        L_RellenarDatos i, 3, nom
        L_RellenarDatos i, 6, nom

Next
End Sub
Private Sub L_LlenarTotalesBsAs()
Dim nom As String
Dim i As Integer

For i = 1 To 6
    sprTotalBsAs.MaxRows = sprTotalBsAs.MaxRows + 1
    
    Select Case i
        Case 1
            nom = "Facturaci�n"
        Case 2
            nom = "Tickets"
        Case 3
            nom = "Pr. x Tickets"
        Case 4
            nom = "Pax Viajados"
        Case 5
            nom = "Pr x Pax Viaj."
        Case 6
            nom = "Captaci�n"
    End Select
    
    sprTotalBsAs.SetText 1, i, nom
    
        L_RellenarDatos_BsAs i, 2, nom
        L_RellenarDatos_BsAs i, 3, nom
        L_RellenarDatos_BsAs i, 6, nom

Next
End Sub

Private Sub L_ocultar(spr As Control, CD As Integer, CH As Integer, FD As Integer, FH As Integer, valor As Boolean)
Dim i


For i = CD To CH
    spr.Col = i
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
sprAep.MaxRows = sprAep.MaxRows + 1
sprAep.SetText 1, sprAep.MaxRows, "Facturaci�n"
sprAep.MaxRows = sprAep.MaxRows + 1
sprAep.SetText 1, sprAep.MaxRows, "Tickets"
sprAep.MaxRows = sprAep.MaxRows + 1
sprAep.SetText 1, sprAep.MaxRows, "Pr. x Tickets"
sprAep.MaxRows = sprAep.MaxRows + 1
sprAep.SetText 1, sprAep.MaxRows, "Pax Viajados"
sprAep.MaxRows = sprAep.MaxRows + 1
sprAep.SetText 1, sprAep.MaxRows, "Pr x Pax Viaj."
sprAep.MaxRows = sprAep.MaxRows + 1
sprAep.SetText 1, sprAep.MaxRows, "Captaci�n"

sprEzeA.MaxRows = sprEzeA.MaxRows + 1
sprEzeA.SetText 1, sprEzeA.MaxRows, "Facturaci�n"
sprEzeA.MaxRows = sprEzeA.MaxRows + 1
sprEzeA.SetText 1, sprEzeA.MaxRows, "Tickets"
sprEzeA.MaxRows = sprEzeA.MaxRows + 1
sprEzeA.SetText 1, sprEzeA.MaxRows, "Pr. x Tickets"
sprEzeA.MaxRows = sprEzeA.MaxRows + 1
sprEzeA.SetText 1, sprEzeA.MaxRows, "Pax Viajados"
sprEzeA.MaxRows = sprEzeA.MaxRows + 1
sprEzeA.SetText 1, sprEzeA.MaxRows, "Pr x Pax Viaj."
sprEzeA.MaxRows = sprEzeA.MaxRows + 1
sprEzeA.SetText 1, sprEzeA.MaxRows, "Captaci�n"

sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Facturaci�n"
sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Tickets"
sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Pr. x Tickets"
sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Pax Viajados"
sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Pr x Pax Viaj."
sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Captaci�n"

sprEzeB.MaxRows = sprEzeB.MaxRows + 1
sprEzeB.SetText 1, sprEzeB.MaxRows, "Facturaci�n"
sprEzeB.MaxRows = sprEzeB.MaxRows + 1
sprEzeB.SetText 1, sprEzeB.MaxRows, "Tickets"
sprEzeB.MaxRows = sprEzeB.MaxRows + 1
sprEzeB.SetText 1, sprEzeB.MaxRows, "Pr. x Tickets"
sprEzeB.MaxRows = sprEzeB.MaxRows + 1
sprEzeB.SetText 1, sprEzeB.MaxRows, "Pax Viajados"
sprEzeB.MaxRows = sprEzeB.MaxRows + 1
sprEzeB.SetText 1, sprEzeB.MaxRows, "Pr x Pax Viaj."
sprEzeB.MaxRows = sprEzeB.MaxRows + 1
sprEzeB.SetText 1, sprEzeB.MaxRows, "Captaci�n"

SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Facturaci�n"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Tickets"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Pr. x Tickets"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Pax Viajados"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Pr x Pax Viaj."
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Captaci�n"


End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim RS As Recordset

On Error GoTo ErrInd:

frmIndic.caption = Aplicacion.SeteoProceso(frmIndic.caption)


sql = " SELECT "
sql = sql & " s.cod_depn, "
sql = sql & " decode(s.cod_depn,'INT','INT',cod_Ssdep) cod_sdep, "
sql = sql & " sum(cant_tickets) cant_t, "
sql = sql & " sum(importe) imp "
sql = sql & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)))
sql = sql & " S ,VENTAS.APERTURA_SDEP A "
sql = sql & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)))
sql = sql & "group by s.cod_depn,decode(s.cod_depn,'INT','INT',cod_Ssdep) "
sql = sql & " order by s.cod_depn,decode(s.cod_depn,'INT','INT',cod_Ssdep) "

sqlX = " SELECT "
sqlX = sqlX & " s.cod_depn, "
sqlX = sqlX & " decode(s.cod_depn,'INT','INT',cod_ssdep) cod_sdep, "
sqlX = sqlX & " sum(cant_tickets) cant_t, "
sqlX = sqlX & " sum(importe) imp "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)) - 1)
sqlX = sqlX & " S ,VENTAS.APERTURA_SDEP A "
sqlX = sqlX & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)) - 1)
sqlX = sqlX & "group by s.cod_depn,decode(s.cod_depn,'INT','INT',cod_ssdep) "
sqlX = sqlX & " order by s.cod_depn,decode(s.cod_depn,'INT','INT',cod_ssdep)"


If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
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
        
        L_Resaltar
        
        Aplicacion.CerrarDAO RsDataAnt
    End If
    
    Aplicacion.CerrarDAO RsData

End If

ErrInd:
    frmIndic.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub


Private Sub L_RellenarDatos(i As Integer, Col As Integer, nom As String)
Dim valor As Variant, divsor As Variant
Dim Impr As Double
    
    If i = 1 Or i = 2 Or i = 4 Then
        sprEzeA.GetText Col, i, valor
        Impr = Impr + valor
        sprEzeAS.GetText Col, i, valor
        Impr = Impr + valor
        sprEzeB.GetText Col, i, valor
        Impr = Impr + valor
        sprAep.GetText Col, i, valor
        Impr = Impr + valor
        SprInt.GetText Col, i, valor
        Impr = Impr + valor
    ElseIf i = 3 Then
        sprTotal.GetText Col, 1, valor
        sprTotal.GetText Col, 2, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 5 Then
        sprTotal.GetText Col, 1, valor
        sprTotal.GetText Col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 6 Then
        sprTotal.GetText Col, 2, valor
        sprTotal.GetText Col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor * 100
        End If
    
    End If
    
    sprTotal.SetText Col, i, Format$((Impr), "#.00")

End Sub
Private Sub L_RellenarDatos_BsAs(i As Integer, Col As Integer, nom As String)
Dim valor As Variant, divsor As Variant
Dim Impr As Double
    
    If i = 1 Or i = 2 Or i = 4 Then
        sprEzeA.GetText Col, i, valor
        Impr = Impr + valor
        sprEzeAS.GetText Col, i, valor
        Impr = Impr + valor
        sprEzeB.GetText Col, i, valor
        Impr = Impr + valor
        sprAep.GetText Col, i, valor
        Impr = Impr + valor
        'SprInt.GetText col, i, valor
        'Impr = Impr + valor
    ElseIf i = 3 Then
        sprTotalBsAs.GetText Col, 1, valor
        sprTotalBsAs.GetText Col, 2, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 5 Then
        sprTotalBsAs.GetText Col, 1, valor
        sprTotalBsAs.GetText Col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 6 Then
        sprTotalBsAs.GetText Col, 2, valor
        sprTotalBsAs.GetText Col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor * 100
        End If
    
    End If
    
    sprTotalBsAs.SetText Col, i, Format$((Impr), "#.00")

End Sub

Private Sub L_Resaltar()
Dim fila As Integer

For fila = 1 To sprTotal.MaxRows
    Spread.spread_ResaltarCelda sprTotal, 4, fila
    Spread.spread_ResaltarCelda sprTotal, 5, fila
    Spread.spread_ResaltarCelda sprTotal, 7, fila
    Spread.spread_ResaltarCelda sprTotal, 8, fila
    
    Spread.spread_ResaltarCelda sprEzeA, 4, fila
    Spread.spread_ResaltarCelda sprEzeA, 5, fila
    Spread.spread_ResaltarCelda sprEzeA, 7, fila
    Spread.spread_ResaltarCelda sprEzeA, 8, fila

    Spread.spread_ResaltarCelda sprEzeAS, 4, fila
    Spread.spread_ResaltarCelda sprEzeAS, 5, fila
    Spread.spread_ResaltarCelda sprEzeAS, 7, fila
    Spread.spread_ResaltarCelda sprEzeAS, 8, fila

    Spread.spread_ResaltarCelda sprEzeB, 4, fila
    Spread.spread_ResaltarCelda sprEzeB, 5, fila
    Spread.spread_ResaltarCelda sprEzeB, 7, fila
    Spread.spread_ResaltarCelda sprEzeB, 8, fila

    Spread.spread_ResaltarCelda sprAep, 4, fila
    Spread.spread_ResaltarCelda sprAep, 5, fila
    Spread.spread_ResaltarCelda sprAep, 7, fila
    Spread.spread_ResaltarCelda sprAep, 8, fila

    Spread.spread_ResaltarCelda SprInt, 4, fila
    Spread.spread_ResaltarCelda SprInt, 5, fila
    Spread.spread_ResaltarCelda SprInt, 7, fila
    Spread.spread_ResaltarCelda SprInt, 8, fila

    Spread.spread_ResaltarCelda sprTotalBsAs, 4, fila
    Spread.spread_ResaltarCelda sprTotalBsAs, 5, fila
    Spread.spread_ResaltarCelda sprTotalBsAs, 7, fila
    Spread.spread_ResaltarCelda sprTotalBsAs, 8, fila

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
        Case INTE
            UsoPorInte = True
        Case total
            UsoPorTotal = False
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorAero Or UsoPorTotal Or UsoPorInte) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub


Private Sub L_TratarEstimado()
Dim sql As String
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

sql = "BEGIN estadis.Estimado_acum ( "
 sql = sql & "'" & Tipo & "',"
sql = sql & anio & ","
sql = sql & Mes & ","
sql = sql & "" & Func.func_ToDate(FD) & ","
sql = sql & "" & Func.func_ToDate(FH) & " ); "
sql = sql & " End;"

If Aplicacion.EjecutarDAO(sql) Then
    sql = "SELECT cod_depn,"
    sql = sql & " decode(cod_depn,'INT','INT',cod_sdep) cod_sdep,"
    sql = sql & " tipo_porc,"
    sql = sql & " sum(valor) valor "
    sql = sql & " FROM estadis.resultados_estim "
    sql = sql & " group by cod_depn,decode(cod_depn,'INT','INT',cod_sdep),tipo_porc"
    
    If Aplicacion.ObtenerRsDAO(sql, RsDataEstim) Then
        L_DecoEstim
    End If
    Aplicacion.TerminarConExitoTrans
    Aplicacion.CerrarDAO RsDataEstim
Else
    Aplicacion.TerminarConErrorTrans
End If

End Sub

Private Sub L_TratarVolados()

Dim sql As String
Dim FD As String, FH As String

If optFechas(0).Value Then
    FD = "01-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText)
    FH = mskFDesde.FormattedText
Else
    FD = mskFDesde.FormattedText
    FH = mskFHasta.FormattedText
End If

L_ControlFecha


    If Control_BsAs Then
    sql = " SELECT "
    sql = sql & "P.cod_depn,"
    sql = sql & "decode(P.cod_depn,'INT','INT',cod_ssdep) cod_sdep,"
'    SQL = SQL & " SUM(pax_AT_VS_vol.PAX_VOL) CANT_V "
    sql = sql & " SUM(cantidad) CANT_V "
    sql = sql & " From estadis.pax_volados P , ventas.apertura_sdep A"
    sql = sql & " Where P.local = A.cod_local and p.cod_depn = a.cod_depn and p.cod_sdep = a.cod_sdep  And "
    sql = sql & " fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
    sql = sql & " GROUP BY p.COD_DEPN,decode(P.cod_depn,'INT','INT',cod_ssdep)"
    
    If Aplicacion.ObtenerRsDAO(sql, RsDataVol) Then
        L_DecoVolados 2
        Aplicacion.CerrarDAO RsDataVol
    End If
        
    Else
        labInfo.caption = "Datos de pasajeros transitados Actualizados al d�a " & Fch_BsAs
        Timer1.Enabled = True
    End If

End Sub

Private Sub L_TratarVoladosHIST()
Dim sql As String
Dim FD As String, FH As String

If optFechas(0).Value Then
    FD = "01-01-" & Year(mskFDesde.FormattedText) - 1
    FH = Day(mskFDesde.FormattedText) & "-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText) - 1
Else
    FD = mskFDesdeAnt.FormattedText
    FH = mskFHastaAnt.FormattedText
End If

        sql = " SELECT "
        sql = sql & "P.cod_depn,"
        sql = sql & "decode(P.cod_depn,'INT','INT',cod_ssdep) cod_sdep,"
        sql = sql & " SUM(p.cantidad) CANT_V "
        sql = sql & " From estadis.pax_volados P, ventas.apertura_sdep A "
        sql = sql & " Where P.local = A.cod_local and p.cod_depn = a.cod_depn and p.cod_sdep = a.cod_sdep  And"
        sql = sql & " fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
        sql = sql & " GROUP BY P.COD_DEPN,decode(P.cod_depn,'INT','INT',cod_ssdep) "
        
        If Aplicacion.ObtenerRsDAO(sql, RsDataVol) Then
            L_DecoVolados 3
            Aplicacion.CerrarDAO RsDataVol
        End If

End Sub


Private Sub L_VentaNeta()
Dim sql As String


sql = " select Vta.cod_depn,Vta.cod_sdep,impDC,impVta "
sql = sql & " From "
sql = sql & " (select P.cod_depn,cod_ssdep cod_sdep ,sum(P.importe) impDC "
sql = sql & " from " & funcLocal_Vista("pago_lpt", Year(CDate(mskFDesde.FormattedText)))
sql = sql & " P , ventas.apertura_sdep A "
sql = sql & " Where "
sql = sql & " p.fch_ticket Between " & Func.func_ToDate(mskFDesde.FormattedText) & " and " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " and p.cod_depn=A.cod_depn and P.cod_sdep=A.cod_sdep and P.cod_local=A.cod_local"
sql = sql & " and tipo_pago = 14"
sql = sql & " group by P.cod_depn,cod_ssdep ) Pag,"
sql = sql & " (select V.cod_depn,cod_ssdep cod_sdep ,sum(V.importe) impVta"
sql = sql & " from " & funcLocal_Vista("venta_lgi", Year(CDate(mskFDesde.FormattedText)))
sql = sql & " V , ventas.apertura_sdep A"
sql = sql & " Where"
sql = sql & " V.fch_ticket Between " & Func.func_ToDate(mskFDesde.FormattedText) & " And " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " and V.cod_depn=A.cod_depn and V.cod_sdep=A.cod_sdep and V.cod_local=A.cod_local"
sql = sql & " and comitente = 'T'"
sql = sql & " group by V.cod_depn,cod_ssdep ) Vta "
sql = sql & " Where"
sql = sql & " Pag.cod_depn = Vta.cod_depn And Pag.cod_sdep = Vta.cod_sdep "
        
If Not Aplicacion.ObtenerRsDAO(sql, RsActual) Then
   MsgBox "No se pudo cargar datos de Duty Check ", vbOKOnly + vbInformation, "Atencion"
End If
        
sql = " select Vta.cod_depn,Vta.cod_sdep,impDC,impVta "
sql = sql & " From "
sql = sql & " (select P.cod_depn,cod_ssdep cod_sdep ,sum(P.importe) impDC "
sql = sql & " from " & funcLocal_Vista("pago_lpt", Year(CDate(mskFDesdeAnt.FormattedText)))
sql = sql & " P , ventas.apertura_sdep A "
sql = sql & " Where "
sql = sql & " p.fch_ticket Between " & Func.func_ToDate(mskFDesdeAnt.FormattedText) & " and " & Func.func_ToDate(mskFHastaAnt.FormattedText)
sql = sql & " and p.cod_depn=A.cod_depn and P.cod_sdep=A.cod_sdep and P.cod_local=A.cod_local"
sql = sql & " and tipo_pago = 14 "
sql = sql & " group by P.cod_depn,cod_ssdep ) Pag,"
sql = sql & " (select V.cod_depn,cod_ssdep cod_sdep ,sum(V.importe) impVta"
sql = sql & " from " & funcLocal_Vista("venta_lgi", Year(CDate(mskFDesdeAnt.FormattedText)))
sql = sql & " V , ventas.apertura_sdep A"
sql = sql & " Where"
sql = sql & " V.fch_ticket Between " & Func.func_ToDate(mskFDesdeAnt.FormattedText) & " And " & Func.func_ToDate(mskFHastaAnt.FormattedText)
sql = sql & " and V.cod_depn=A.cod_depn and V.cod_sdep=A.cod_sdep and V.cod_local=A.cod_local"
sql = sql & " and comitente = 'T'"
sql = sql & " group by V.cod_depn,cod_ssdep ) Vta "
sql = sql & " Where"
sql = sql & " Pag.cod_depn = Vta.cod_depn And Pag.cod_sdep = Vta.cod_sdep "
        
If Not Aplicacion.ObtenerRsDAO(sql, RsAnterior) Then
   MsgBox "No se pudo cargar datos de Duty Check ", vbOKOnly + vbInformation, "Atencion"
End If
        
End Sub

Private Sub borGrBar_Click(Index As Integer)
Dim col_datos As Collection

    Set col_datos = New Collection
    Select Case Index
        Case 0
            L_LLenarColeccion col_datos, sprTotalBsAs
            FrmGrafbar.CargarGrafico total, "", col_datos
        Case 1
            L_LLenarColeccion col_datos, sprEzeA
            FrmGrafbar.CargarGrafico EZEA, "", col_datos
        Case 2
            L_LLenarColeccion col_datos, sprEzeB
            FrmGrafbar.CargarGrafico EZEB, "", col_datos
        Case 3
            L_LLenarColeccion col_datos, sprAep
            FrmGrafbar.CargarGrafico AERO, "", col_datos
        Case 4
            L_LLenarColeccion col_datos, SprInt
            FrmGrafbar.CargarGrafico INTE, "", col_datos
        Case 5
            L_LLenarColeccion col_datos, sprTotal
            FrmGrafbar.CargarGrafico total, "", col_datos
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



Private Sub botExcel_Click()
    If optFechas(0).Value Then
        L_TratarExcel sprTotal, "INDICADORES ACUMULADOS ANUALES (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
        "del 1 de Ene al " & Format$(mskFDesde.FormattedText, "dd-mmm"), "", Exl_Blanco
    Else
        L_TratarExcel sprTotal, "INDICADORES ACUMULADOS MES CORRIENTE (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
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
L_ocultar sprTotalBsAs, 6, 9, 1, sprTotalBsAs.MaxRows, ocultar
L_ocultar sprTotal, 6, 9, 1, sprTotal.MaxRows, ocultar
L_ocultar SprInt, 6, 9, 1, SprInt.MaxRows, ocultar
L_ocultar sprEzeA, 6, 9, 1, sprEzeA.MaxRows, ocultar
L_ocultar sprEzeB, 6, 9, 1, sprEzeB.MaxRows, ocultar
L_ocultar sprEzeAS, 6, 9, 1, sprEzeAS.MaxRows, ocultar
L_ocultar sprAep, 6, 9, 1, sprAep.MaxRows, ocultar

End Sub

Private Sub Command1_Click()
    If optFechas(0).Value Then
        L_TratarExcel sprTotal, "INDICADORES ACUMULADOS ANUALES (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
        "del 1 de Ene al " & Format$(mskFDesde.FormattedText, "dd-mmm"), "", Exl_Blanco
    Else
        L_TratarExcel sprTotal, "INDICADORES ACUMULADOS MES CORRIENTE (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
        " del " & Format$(mskFDesde.FormattedText, "dd-mmm") & " al " & Format$(mskFHasta.FormattedText, "dd-mmm"), "", Exl_Blanco 'Exl_Ros '
    End If
End Sub

Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabEspigon

End Sub

Private Sub L_LLenarColeccion(ByRef Col As Collection, spr As Control)
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

Private Sub L_TratarExcel(spr As Control, titulo As String, subTit As String, Info As String, color As Integer)
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

frmIndic.caption = Aplicacion.SeteoProceso(frmIndic.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.Application.Visible = True
    
    ReDim titCol(spr.MaxCols)
    Col = 1
    fila = 3
    
    For i = 1 To spr.MaxCols
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
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    fila = fila + 1
    
    Exl_PonerValor AppExcel, fila, Col, "Total Compan�a"
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    AppExcel.Application.Range(rango).Merge
        
    fila = fila + 1
    
    Exl_BajarGrillaExel spr, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + spr.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '-------------------------
    fila = fila + 1
    
    Exl_PonerValor AppExcel, fila, Col, "Total Buenos Aires"
    rango = Exl_rangos(fila, fila, 1, sprTotalBsAs.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    AppExcel.Application.Range(rango).Merge
        
    fila = fila + 1
    
    Exl_BajarGrillaExel sprTotalBsAs, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprTotalBsAs.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, "INTERNACIONAL A Llegadas"
    rango = Exl_rangos(fila, fila, 1, sprEzeA.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel sprEzeA, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeA.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, "INTERNACIONAL A Salidas "
    rango = Exl_rangos(fila, fila, 1, sprEzeAS.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel sprEzeAS, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeAS.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
        
    '-----------------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, "INTERNACIONAL B "
    rango = Exl_rangos(fila, fila, 1, sprEzeB.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba
    
    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel sprEzeB, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeB.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '--------------------------------
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, "AEROPARQUE "
    rango = Exl_rangos(fila, fila, 1, sprAep.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba
    
    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel sprAep, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + sprAep.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
        
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris

    '------------------------
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, "INTERIOR "
    rango = Exl_rangos(fila, fila, 1, SprInt.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprInt, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + SprInt.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, Col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '-----------------------------------

    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.Zoom = False
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmIndic.caption = Aplicacion.SeteoFin
    Exit Sub
    
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

frmPrincipal.lstForms.AddItem "frmindic"

End Sub


Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmindic"
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
            sprTotal.SetFocus
        Case 1
            sprEzeA.SetFocus
        Case 2
            sprEzeAS.SetFocus
        Case 3
            sprEzeB.SetFocus
        Case 4
            sprAep.SetFocus
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


