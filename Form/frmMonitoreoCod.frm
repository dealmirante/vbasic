VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmMonitoreoCod 
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
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
      Height          =   3450
      Left            =   15
      TabIndex        =   1
      Top             =   675
      Width           =   9405
      Begin FPSpread.vaSpread SprImpResumen 
         Height          =   3090
         Left            =   90
         OleObjectBlob   =   "frmMonitoreoCod.frx":0000
         TabIndex        =   2
         Top             =   240
         Width           =   9195
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
Attribute VB_Name = "frmMonitoreoCod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tit As String
Dim prov As String

Dim TotalGral As Single
Dim TotalGralant As Single

Dim sql As String

Dim aniomes As String

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
        
    SprImpResumen.SetText 1, linea, Trim(RS!Cod_prod)
    SprImpResumen.SetText 2, linea, str(RS!e_u)
    SprImpResumen.SetText 3, linea, str(RS!v_u)
    SprImpResumen.SetText 5, linea, str(RS!e_p)
    SprImpResumen.SetText 6, linea, str(RS!v_p)
    SprImpResumen.SetText 8, linea, str(RS!stk)
    
    RS.MoveNext
         If RS.EOF Then
            Exit Do
         End If
        linea = linea + 1

      'Loop
    
    'End Select
Loop

    'Spread.Spread_TotalesGrillas SprImpResumen, 2, 2
    'L_Participacion
    
End Sub




Private Sub L_Refrescar()

frmMonitoreoCod.caption = Aplicacion.SeteoProceso(frmMonitoreoCod.caption)

lblTit.caption = tit & " : " & "" & "-" & prov
DoEvents
If Aplicacion.ObtenerRsDAO(sql, RS) Then
    If Aplicacion.CantReg(RS) > 0 Then
        L_DecoEspigon
    End If
End If

frmMonitoreoCod.caption = Aplicacion.SeteoFin

End Sub

Public Sub Mostrar(pSQL As String, pTit As String, pProv As String)

prov = pProv
tit = pTit
'TotalGral = pTot
'TotalGralant = pTotAnt
aniomes = frmMonitoreo.mskAnio.Text & "01 AND " & frmMonitoreo.mskAnio.Text & Format(frmMonitoreo.mskMes.Text - 1, "00")
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


Private Sub SprImpResumen_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim sql As String
Dim prod As String, descProd As String, PO As Variant
Dim pos As Integer

SprImpResumen.GetText 1, Row, PO

pos = InStr(1, PO, "-")

prod = Left(Trim(PO), pos - 1)
descProd = Right(Trim(PO), Len(PO) - (pos))

    sql = "  SELECT ANIOMES"
    sql = sql & " ,sum(venta_real_U)venta_real_U"
    sql = sql & " ,sum(estimado_U)estimado_U"
    sql = sql & " ,sum(venta_real_$)venta_real_P"
    sql = sql & " ,sum(estimado_$)estimado_P"
    sql = sql & " From baires.estm_venta"
    sql = sql & " Where"
    sql = sql & " aniomes BETWEEN " & aniomes
    sql = sql & " and cod_rubr <> 'REG'"
    sql = sql & " and cod_prod = '" & prod & "' "
    sql = sql & " group by ANIOMES "
    
    frmMonitoreoMes.Mostrar sql, Trim(prod), Trim(descProd), "", "Producto"


End Sub


