VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Begin VB.Form frmVentaProvMetro 
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   810
      Left            =   7620
      TabIndex        =   3
      Top             =   -75
      Width           =   1875
      Begin VB.OptionButton optSort 
         Caption         =   "Por Imp. (Desc)"
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
         Left            =   180
         TabIndex        =   5
         Top             =   495
         Width           =   1530
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Por Nombre"
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
         Left            =   165
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3690
      Left            =   15
      TabIndex        =   1
      Top             =   735
      Width           =   9405
      Begin FPSpread.vaSpread SprImpResumen 
         Height          =   3330
         Left            =   120
         OleObjectBlob   =   "frmVentaProvMetro.frx":0000
         TabIndex        =   2
         Top             =   240
         Width           =   9195
      End
   End
   Begin VB.Label lblMts 
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
      Height          =   330
      Left            =   615
      TabIndex        =   6
      Top             =   480
      Width           =   6570
   End
   Begin VB.Label lblTit 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   330
      TabIndex        =   0
      Top             =   105
      Width           =   6870
   End
End
Attribute VB_Name = "frmVentaProvMetro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tit As String
Dim tit2 As String
Dim sector As Integer
Dim vTipo As Integer

Dim TotalGral As Single
Dim TotalGralant As Single

Dim sql As String

Dim RS As Recordset


Private Sub L_DecoEspigon()
Dim Sdep As String
Dim rubro As Variant
Dim linea As Integer
 
SprImpResumen.MaxRows = 0
linea = 1
Do While Not RS.EOF
     
    SprImpResumen.MaxRows = SprImpResumen.MaxRows + 1
        
    SprImpResumen.SetText 1, SprImpResumen.MaxRows, Trim(RS!Descrip)
    If vTipo = 0 Then
      SprImpResumen.SetText 2, SprImpResumen.MaxRows, str(RS!actimp)
      SprImpResumen.SetText 3, SprImpResumen.MaxRows, str(RS!antimp)
    Else
      If RS!ActMtr > 0 Then
         SprImpResumen.SetText 2, SprImpResumen.MaxRows, str(RS!actimp / RS!ActMtr)
      End If
      If RS!AntMtr > 0 Then
         SprImpResumen.SetText 3, SprImpResumen.MaxRows, str(RS!antimp / RS!AntMtr)
      End If
    End If
    'SprImpResumen.SetText 5, linea, str(RS!e_p)
    'SprImpResumen.SetText 6, linea, str(RS!v_p)
    'SprImpResumen.SetText 8, linea, str(RS!stk)
    
    RS.MoveNext

Loop

    Spread.Spread_TotalesGrillas SprImpResumen, 2, 2
    L_Participacion
    
End Sub

Private Sub L_Participacion()
Dim i As Integer
Dim fila As Integer
Dim valor As Variant
Dim total As Variant
Dim ValorAnt As Variant
Dim TotalANT As Variant
Dim part As Single
   
   SprImpResumen.GetText 2, SprImpResumen.MaxRows, total
   SprImpResumen.GetText 3, SprImpResumen.MaxRows, TotalANT
   
   
   For fila = 1 To SprImpResumen.MaxRows - 1
       SprImpResumen.GetText 2, fila, valor
       SprImpResumen.GetText 3, fila, ValorAnt
       part = 0
       If total > 0 Then
           SprImpResumen.SetText 5, fila, str(valor * 100 / total)
       End If
       If TotalANT > 0 Then
           SprImpResumen.SetText 6, fila, str(ValorAnt * 100 / TotalANT)
       End If
       If part <> 0 Then
          SprImpResumen.SetText 10, fila, Format(part, "###0.00")
       End If
   Next fila

'   SprImpResumen.SetText 4, SprImpResumen.MaxRows, "."
'   SprImpResumen.SetText 7, SprImpResumen.MaxRows, "."
End Sub




Private Sub L_Refrescar()

frmVentaProvMetro.caption = Aplicacion.SeteoProceso(frmVentaProvMetro.caption)

lblTit.caption = tit
lblMts.caption = tit2

DoEvents
If Aplicacion.ObtenerRsDAO(sql, RS) Then
    If Aplicacion.CantReg(RS) > 0 Then
        L_DecoEspigon
    End If
End If

frmVentaProvMetro.caption = Aplicacion.SeteoFin

End Sub

Public Sub Mostrar(pSQL As String, pTit As String, pTit2 As String, pSec As String, pTipo As Integer)

sector = pSec
tit = pTit
tit2 = pTit2

vTipo = pTipo

'TotalGral = pTot
'TotalGralant = pTotAnt
sql = pSQL


Me.Show 1

End Sub


Private Sub Form_Load()

Top = 2500
Left = 500

DoEvents

L_Refrescar

End Sub

Private Sub optSort_Click(Index As Integer)

Select Case Index
    Case 0
        SprImpResumen.Row = 1
        SprImpResumen.Col = 1
        SprImpResumen.Row2 = SprImpResumen.MaxRows - 1
        SprImpResumen.Col2 = 10
        
        ' Set sort definition for key 1
        SprImpResumen.SortBy = SS_SORT_BY_ROW

        SprImpResumen.SortKey(1) = 1
        SprImpResumen.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        SprImpResumen.Action = SS_ACTION_SORT
    Case 1
        SprImpResumen.Row = 1
        SprImpResumen.Col = 1
        SprImpResumen.Row2 = SprImpResumen.MaxRows - 1
        SprImpResumen.Col2 = 10
        
        ' Set sort definition for key 1
        SprImpResumen.SortBy = SS_SORT_BY_ROW
        
        SprImpResumen.SortKey(1) = 6
        SprImpResumen.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        SprImpResumen.Action = SS_ACTION_SORT
End Select

End Sub


