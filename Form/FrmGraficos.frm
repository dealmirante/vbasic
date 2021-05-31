VERSION 5.00
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form FrmGraficos 
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3900
      Left            =   15
      TabIndex        =   0
      Top             =   -75
      Width           =   7650
      Begin VB.OptionButton Option1 
         Caption         =   "Vista Nominal"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   195
         TabIndex        =   3
         Top             =   3375
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Vista Porcentual"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   195
         TabIndex        =   2
         Top             =   3030
         Width           =   1695
      End
      Begin GraphLib.Graph GRTot 
         Height          =   3285
         Left            =   2040
         TabIndex        =   1
         Top             =   555
         Width           =   5520
         _Version        =   65536
         _ExtentX        =   9737
         _ExtentY        =   5794
         _StockProps     =   96
         BorderStyle     =   1
         AutoInc         =   0
         Background      =   9
         Foreground      =   15
         GraphStyle      =   2
         GraphType       =   2
         GridStyle       =   3
         LegendStyle     =   1
         NumPoints       =   7
         Palette         =   1
         PrintStyle      =   2
         RandomData      =   1
         ColorData       =   7
         ColorData[0]    =   11
         ColorData[2]    =   14
         ColorData[3]    =   10
         ColorData[4]    =   7
         ColorData[5]    =   4
         ColorData[6]    =   15
         ExtraData       =   0
         ExtraData[]     =   0
         FontFamily      =   4
         FontSize        =   4
         FontSize[0]     =   150
         FontSize[1]     =   150
         FontSize[2]     =   100
         FontSize[3]     =   100
         FontStyle       =   4
         FontStyle[0]    =   2
         GraphData       =   1
         GraphData[]     =   5
         GraphData[0,0]  =   0
         GraphData[0,1]  =   0
         GraphData[0,2]  =   0
         GraphData[0,3]  =   0
         GraphData[0,4]  =   0
         LabelText       =   0
         LegendText      =   1
         LegendText[0]   =   "Uno"
         PatternData     =   0
         SymbolData      =   0
         XPosData        =   0
         XPosData[]      =   0
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   165
         TabIndex        =   8
         Top             =   2040
         Width           =   1785
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   7
         Top             =   600
         Width           =   1920
      End
      Begin VB.Label Label3 
         Caption         =   "Informaci�n sobre :"
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
         Left            =   450
         TabIndex        =   6
         Top             =   1740
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grafico de Participaci�n de ventas, por Local, para una fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   5
         Top             =   180
         Width           =   7455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   330
         TabIndex        =   4
         Top             =   975
         Width           =   1680
      End
   End
End
Attribute VB_Name = "FrmGraficos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim grafico As String
Dim info_local As String
Dim col_local As Collection



Public Sub CargarGrafico(graf As String, Info As String, Col As Collection)

grafico = graf
info_local = Info
Set col_local = Col

Me.Show 1

End Sub

Private Sub Form_Load()

Me.Left = 600
Me.Top = 1500

Label1.caption = col_local.Item(1).fch
If Label1.caption = "" Then
    Label1.caption = "TOTALES"
End If
Label5.caption = info_local
Select Case grafico
    Case TOTAL
        Label4.caption = "COMP�IA"
    Case EZEA
        Label4.caption = "EZEIZA (INT A)"
    Case EZEB
        Label4.caption = "EZEIZA (INT B)"
    Case AERO
        Label4.caption = "AEROPARQUE"
End Select
Func_CargarGraficos GRTot, col_local

End Sub


Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
GRTot.GraphStyle = 6
Else
GRTot.GraphStyle = 2
End If
GRTot.DrawMode = 2
End Sub


