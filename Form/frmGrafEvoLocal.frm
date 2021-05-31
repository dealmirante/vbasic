VERSION 5.00
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form frmGrafEvoLocal 
   Caption         =   "Gráfico de evolución"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   60
      TabIndex        =   6
      Top             =   4065
      Width           =   8145
      Begin VB.CheckBox chkLocal 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   7110
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   6000
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   4
         Left            =   4860
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   3
         Left            =   3690
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   2
         Left            =   2505
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   1365
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4125
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   8160
      Begin GraphLib.Graph Gr 
         Height          =   3225
         Left            =   75
         TabIndex        =   1
         Top             =   840
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   5689
         _StockProps     =   96
         BorderStyle     =   1
         Background      =   9
         Foreground      =   15
         GraphStyle      =   5
         GraphType       =   6
         GridStyle       =   1
         IndexStyle      =   1
         LegendStyle     =   1
         NumPoints       =   30
         NumSets         =   7
         Palette         =   1
         PatternedLines  =   1
         PrintStyle      =   1
         RandomData      =   1
         ThickLines      =   0
         YAxisStyle      =   2
         ColorData       =   7
         ColorData[0]    =   11
         ColorData[1]    =   14
         ColorData[3]    =   7
         ColorData[4]    =   13
         ColorData[5]    =   15
         ColorData[6]    =   1
         ExtraData       =   0
         ExtraData[]     =   0
         FontFamily      =   4
         FontSize        =   4
         FontSize[0]     =   200
         FontSize[1]     =   150
         FontSize[2]     =   100
         FontSize[3]     =   100
         FontStyle       =   4
         GraphData       =   7
         GraphData[]     =   30
         GraphData[0,0]  =   0
         GraphData[0,1]  =   0
         GraphData[0,2]  =   0
         GraphData[0,3]  =   0
         GraphData[0,4]  =   0
         GraphData[0,5]  =   0
         GraphData[0,6]  =   0
         GraphData[0,7]  =   0
         GraphData[0,8]  =   0
         GraphData[0,9]  =   0
         GraphData[0,10] =   0
         GraphData[0,11] =   0
         GraphData[0,12] =   0
         GraphData[0,13] =   0
         GraphData[0,14] =   0
         GraphData[0,15] =   0
         GraphData[0,16] =   0
         GraphData[0,17] =   0
         GraphData[0,18] =   0
         GraphData[0,19] =   0
         GraphData[0,20] =   0
         GraphData[0,21] =   0
         GraphData[0,22] =   0
         GraphData[0,23] =   0
         GraphData[0,24] =   0
         GraphData[0,25] =   0
         GraphData[0,26] =   0
         GraphData[0,27] =   0
         GraphData[0,28] =   0
         GraphData[0,29] =   0
         GraphData[1,0]  =   0
         GraphData[1,1]  =   0
         GraphData[1,2]  =   0
         GraphData[1,3]  =   0
         GraphData[1,4]  =   0
         GraphData[1,5]  =   0
         GraphData[1,6]  =   0
         GraphData[1,7]  =   0
         GraphData[1,8]  =   0
         GraphData[1,9]  =   0
         GraphData[1,10] =   0
         GraphData[1,11] =   0
         GraphData[1,12] =   0
         GraphData[1,13] =   0
         GraphData[1,14] =   0
         GraphData[1,15] =   0
         GraphData[1,16] =   0
         GraphData[1,17] =   0
         GraphData[1,18] =   0
         GraphData[1,19] =   0
         GraphData[1,20] =   0
         GraphData[1,21] =   0
         GraphData[1,22] =   0
         GraphData[1,23] =   0
         GraphData[1,24] =   0
         GraphData[1,25] =   0
         GraphData[1,26] =   0
         GraphData[1,27] =   0
         GraphData[1,28] =   0
         GraphData[1,29] =   0
         GraphData[2,0]  =   0
         GraphData[2,1]  =   0
         GraphData[2,2]  =   0
         GraphData[2,3]  =   0
         GraphData[2,4]  =   0
         GraphData[2,5]  =   0
         GraphData[2,6]  =   0
         GraphData[2,7]  =   0
         GraphData[2,8]  =   0
         GraphData[2,9]  =   0
         GraphData[2,10] =   0
         GraphData[2,11] =   0
         GraphData[2,12] =   0
         GraphData[2,13] =   0
         GraphData[2,14] =   0
         GraphData[2,15] =   0
         GraphData[2,16] =   0
         GraphData[2,17] =   0
         GraphData[2,18] =   0
         GraphData[2,19] =   0
         GraphData[2,20] =   0
         GraphData[2,21] =   0
         GraphData[2,22] =   0
         GraphData[2,23] =   0
         GraphData[2,24] =   0
         GraphData[2,25] =   0
         GraphData[2,26] =   0
         GraphData[2,27] =   0
         GraphData[2,28] =   0
         GraphData[2,29] =   0
         GraphData[3,0]  =   0
         GraphData[3,1]  =   0
         GraphData[3,2]  =   0
         GraphData[3,3]  =   0
         GraphData[3,4]  =   0
         GraphData[3,5]  =   0
         GraphData[3,6]  =   0
         GraphData[3,7]  =   0
         GraphData[3,8]  =   0
         GraphData[3,9]  =   0
         GraphData[3,10] =   0
         GraphData[3,11] =   0
         GraphData[3,12] =   0
         GraphData[3,13] =   0
         GraphData[3,14] =   0
         GraphData[3,15] =   0
         GraphData[3,16] =   0
         GraphData[3,17] =   0
         GraphData[3,18] =   0
         GraphData[3,19] =   0
         GraphData[3,20] =   0
         GraphData[3,21] =   0
         GraphData[3,22] =   0
         GraphData[3,23] =   0
         GraphData[3,24] =   0
         GraphData[3,25] =   0
         GraphData[3,26] =   0
         GraphData[3,27] =   0
         GraphData[3,28] =   0
         GraphData[3,29] =   0
         GraphData[4,0]  =   0
         GraphData[4,1]  =   0
         GraphData[4,2]  =   0
         GraphData[4,3]  =   0
         GraphData[4,4]  =   0
         GraphData[4,5]  =   0
         GraphData[4,6]  =   0
         GraphData[4,7]  =   0
         GraphData[4,8]  =   0
         GraphData[4,9]  =   0
         GraphData[4,10] =   0
         GraphData[4,11] =   0
         GraphData[4,12] =   0
         GraphData[4,13] =   0
         GraphData[4,14] =   0
         GraphData[4,15] =   0
         GraphData[4,16] =   0
         GraphData[4,17] =   0
         GraphData[4,18] =   0
         GraphData[4,19] =   0
         GraphData[4,20] =   0
         GraphData[4,21] =   0
         GraphData[4,22] =   0
         GraphData[4,23] =   0
         GraphData[4,24] =   0
         GraphData[4,25] =   0
         GraphData[4,26] =   0
         GraphData[4,27] =   0
         GraphData[4,28] =   0
         GraphData[4,29] =   0
         GraphData[5,0]  =   0
         GraphData[5,1]  =   0
         GraphData[5,2]  =   0
         GraphData[5,3]  =   0
         GraphData[5,4]  =   0
         GraphData[5,5]  =   0
         GraphData[5,6]  =   0
         GraphData[5,7]  =   0
         GraphData[5,8]  =   0
         GraphData[5,9]  =   0
         GraphData[5,10] =   0
         GraphData[5,11] =   0
         GraphData[5,12] =   0
         GraphData[5,13] =   0
         GraphData[5,14] =   0
         GraphData[5,15] =   0
         GraphData[5,16] =   0
         GraphData[5,17] =   0
         GraphData[5,18] =   0
         GraphData[5,19] =   0
         GraphData[5,20] =   0
         GraphData[5,21] =   0
         GraphData[5,22] =   0
         GraphData[5,23] =   0
         GraphData[5,24] =   0
         GraphData[5,25] =   0
         GraphData[5,26] =   0
         GraphData[5,27] =   0
         GraphData[5,28] =   0
         GraphData[5,29] =   0
         GraphData[6,0]  =   0
         GraphData[6,1]  =   0
         GraphData[6,2]  =   0
         GraphData[6,3]  =   0
         GraphData[6,4]  =   0
         GraphData[6,5]  =   0
         GraphData[6,6]  =   0
         GraphData[6,7]  =   0
         GraphData[6,8]  =   0
         GraphData[6,9]  =   0
         GraphData[6,10] =   0
         GraphData[6,11] =   0
         GraphData[6,12] =   0
         GraphData[6,13] =   0
         GraphData[6,14] =   0
         GraphData[6,15] =   0
         GraphData[6,16] =   0
         GraphData[6,17] =   0
         GraphData[6,18] =   0
         GraphData[6,19] =   0
         GraphData[6,20] =   0
         GraphData[6,21] =   0
         GraphData[6,22] =   0
         GraphData[6,23] =   0
         GraphData[6,24] =   0
         GraphData[6,25] =   0
         GraphData[6,26] =   0
         GraphData[6,27] =   0
         GraphData[6,28] =   0
         GraphData[6,29] =   0
         LabelText       =   3
         LegendText      =   0
         PatternData     =   1
         SymbolData      =   7
         SymbolData[0]   =   7
         SymbolData[1]   =   5
         SymbolData[2]   =   8
         SymbolData[3]   =   3
         SymbolData[5]   =   9
         SymbolData[6]   =   6
         XPosData        =   0
         XPosData[]      =   3
      End
      Begin VB.Label Label3 
         Caption         =   "Información sobre :"
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
         Left            =   5175
         TabIndex        =   5
         Top             =   585
         Width           =   1515
      End
      Begin VB.Label Label5 
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
         Left            =   6765
         TabIndex        =   4
         Top             =   570
         Width           =   1260
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   3
         Top             =   570
         Width           =   1920
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grafico de Evolución de ventas, por Local, para "
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
         Left            =   75
         TabIndex        =   2
         Top             =   165
         Width           =   7995
      End
   End
End
Attribute VB_Name = "frmGrafEvoLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim grafico As String
Dim info_local As String
Dim spr_local As Control
Dim FD_local As String
Dim FH_local As String


Dim LocEleg(1 To 10) As T_Col

Private Function L_CantLocales() As Integer
Dim i, cont As Integer
Dim valor As Variant

L_LimpiarElecion
cont = 0
For i = 0 To 5
    If chkLocal(i).Value Then
        cont = cont + 1
        LocEleg(cont).iCol = i + 2
        spr_local.GetText i + 2, 0, valor
        LocEleg(cont).sLocal = valor
    End If
Next
L_CantLocales = cont
End Function

Private Sub L_LimpiarElecion()
Dim i

For i = 1 To 10
    LocEleg(i).iCol = -1
    LocEleg(i).sLocal = ""
Next

End Sub


Private Sub L_SetearCheck()
Dim i
Dim valor As Variant

For i = 1 To spr_local.MaxCols - 2
    chkLocal(i - 1).Visible = True
    spr_local.GetText i + 1, 0, valor
    chkLocal(i - 1).caption = valor
Next
chkLocal(0).Value = 1
End Sub

Public Sub L_TratarGrafico()
Dim i, j
Dim cant As Integer, Puntos As Integer
Dim valor As Variant

'Graph1.Visible = False

On Error GoTo ErrGr:



cant = L_CantLocales


If cant > 0 And spr_local.MaxRows > 1 Then
'Gr.NumPoints = Func_CantidadDias("01-" & Trim(str(Mes_local)) & "-" & Trim(str(Anio_local)))
Puntos = spr_local.MaxRows - 1
If Puntos > 31 Then
    Puntos = 31
End If
GR.NumPoints = Puntos
GR.AutoInc = 0
    

GR.YAxisMax = (L_MaximoValor(spr_local.MaxRows - 1, cant, spr_local) / 100) * 100

GR.YAxisTicks = 7

GR.NumSets = cant
i = 1
For j = 1 To GR.NumSets
    GR.ThisSet = j
    GR.LegendText = LocEleg(j).sLocal
    
    For i = 1 To spr_local.MaxRows - 1
        GR.ThisPoint = i
        spr_local.GetText 1, i, valor
        If valor <> "" Then
            GR.LabelText = Day(CDate(valor))
        End If
        spr_local.GetText LocEleg(j).iCol, i, valor
        GR.GraphData = valor
    Next
Next

Else
    GR.YAxisMax = 1
    GR.NumSets = 1
    GR.LegendText = "NADA"
    GR.NumPoints = 2
    GR.ThisPoint = 1
    GR.GraphData = 0
    GR.ThisPoint = 2
    GR.GraphData = 1

End If

GR.DrawMode = 2
ErrGr:
    Exit Sub
    
End Sub

Public Function L_MaximoValor(fila As Long, col As Integer, spr As Control) As Long
Dim i, j
Dim Max As Long
Dim valor As Variant

Max = 0

For i = 1 To fila
    For j = 1 To col
        spr.GetText LocEleg(j).iCol, i, valor
        If Val(valor) > Max Then
            Max = valor
        End If
    Next
Next

L_MaximoValor = Max
End Function


Private Sub chkLocal_Click(Index As Integer)
L_TratarGrafico
End Sub

Private Sub Form_Load()
Me.Left = 600
Me.Top = 1500

Label2.caption = Label2.caption & " " & FD_local & " al " & FH_local

Label5.caption = info_local
Select Case grafico
    Case TOTAL
        Label4.caption = "COMPAÑIA"
    Case EZEA
        Label4.caption = "EZEIZA (INT A)"
    Case EZEB
        Label4.caption = "EZEIZA (INT B)"
    Case AERO
        Label4.caption = "AEROPARQUE"
End Select
L_SetearCheck
L_TratarGrafico

End Sub
Public Sub CargarGrafico(graf As String, Info As String, spr As Control, FD As String, FH As String)

grafico = graf
info_local = Info
Set spr_local = spr
FD_local = FD
FH_local = FH

Me.Show 1

End Sub

