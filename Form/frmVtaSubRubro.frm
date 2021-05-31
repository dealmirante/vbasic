VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmVtaSubRubro 
   Caption         =   "frmVtaSubRubro"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   810
      Left            =   8565
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
      Height          =   3435
      Left            =   75
      TabIndex        =   1
      Top             =   675
      Width           =   10380
      Begin FPSpread.vaSpread SprImpResumen 
         Height          =   3090
         Left            =   165
         OleObjectBlob   =   "frmVtaSubRubro.frx":0000
         TabIndex        =   2
         Top             =   240
         Width           =   9915
      End
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
      Height          =   480
      Left            =   330
      TabIndex        =   0
      Top             =   105
      Width           =   8415
   End
End
Attribute VB_Name = "frmVtaSubRubro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tit As String
Dim rub As String

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
 
        'Do While Sdep = RS!cod_sdep
        
    SprImpResumen.MaxRows = SprImpResumen.MaxRows + 1
        
    SprImpResumen.SetText 1, linea, Trim(RS!Descr)
    SprImpResumen.SetText 2, linea, str(RS!Importe_act)
    SprImpResumen.SetText 3, linea, str(RS!Importe_ant)
                                
    RS.MoveNext
         If RS.EOF Then
            Exit Do
         End If
        linea = linea + 1

      'Loop
    
    'End Select
Loop

    Spread.Spread_TotalesGrillas SprImpResumen, 2, 2
    L_Participacion
    
End Sub




Private Sub L_Refrescar()

frmVtaSubRubro.caption = Aplicacion.SeteoProceso(frmVtaSubRubro.caption)

lblTit.caption = Tit & " : " & "" & "-" & rub
DoEvents
If Aplicacion.ObtenerRsDAO(sql, RS) Then
    If Aplicacion.CantReg(RS) > 0 Then
        L_DecoEspigon
    End If
End If

frmVtaSubRubro.caption = Aplicacion.SeteoFin

End Sub

Public Sub Mostrar(pNac As String, pRub As String, pSQL As String, pTot As Single, pTotAnt As Single, pTit As String)

rub = pRub
Tit = pTit
TotalGral = pTot
TotalGralant = pTotAnt

sql = pSQL


Me.Show 1

End Sub


Private Sub Form_Activate()
    L_Refrescar
End Sub

Private Sub L_Participacion()
Dim i As Integer
Dim fila As Integer
Dim valor As Variant
Dim Total As Variant
Dim ValorAnt As Variant
Dim TotalANT As Variant
Dim part As Single
   
   SprImpResumen.GetText 2, SprImpResumen.MaxRows, Total
   SprImpResumen.GetText 3, SprImpResumen.MaxRows, TotalANT
   
   
   For fila = 1 To SprImpResumen.MaxRows - 1
       SprImpResumen.GetText 2, fila, valor
       SprImpResumen.GetText 3, fila, ValorAnt
       part = 0
       If Total > 0 Then
           SprImpResumen.SetText 5, fila, str(valor * 100 / Total)
       End If
       If TotalANT > 0 Then
           SprImpResumen.SetText 6, fila, str(ValorAnt * 100 / TotalANT)
       End If
       If TotalGral > 0 Then
           SprImpResumen.SetText 8, fila, Format((valor * 100 / TotalGral), "####0.00")
       End If
       If TotalGralant > 0 Then
           SprImpResumen.SetText 9, fila, str(Format((ValorAnt * 100 / TotalGralant), "####0.00"))
           
       End If
       If ValorAnt > 0 Then
          part = 100 * (((valor * 100 / TotalGral) / (ValorAnt * 100 / TotalGralant)) - 1)
       End If
       If part <> 0 Then
          SprImpResumen.SetText 10, fila, Format(part, "###0.00")
       End If
   Next fila

'   SprImpResumen.SetText 4, SprImpResumen.MaxRows, "."
'   SprImpResumen.SetText 7, SprImpResumen.MaxRows, "."
End Sub

Private Sub Form_Load()

Top = 2500
Left = 200

DoEvents

End Sub

Private Sub optSort_Click(Index As Integer)

Select Case Index
    Case 0
        SprImpResumen.Row = 1
        SprImpResumen.col = 1
        SprImpResumen.Row2 = SprImpResumen.MaxRows - 1
        SprImpResumen.Col2 = 10
        
        ' Set sort definition for key 1
        SprImpResumen.SortBy = SS_SORT_BY_ROW

        SprImpResumen.SortKey(1) = 1
        SprImpResumen.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        SprImpResumen.Action = SS_ACTION_SORT
    Case 1
        SprImpResumen.Row = 1
        SprImpResumen.col = 1
        SprImpResumen.Row2 = SprImpResumen.MaxRows - 1
        SprImpResumen.Col2 = 10
        
        ' Set sort definition for key 1
        SprImpResumen.SortBy = SS_SORT_BY_ROW
        
        SprImpResumen.SortKey(1) = 2
        SprImpResumen.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        SprImpResumen.Action = SS_ACTION_SORT
End Select

End Sub


