VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMonedas 
   Caption         =   "Participación por monedas"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   8265
   Begin VB.Frame frdatos 
      Height          =   1245
      Left            =   90
      TabIndex        =   6
      Top             =   -15
      Width           =   4605
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2910
         Picture         =   "frmMonedas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   2895
         Picture         =   "frmMonedas.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   750
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   315
         Left            =   1710
         TabIndex        =   9
         Top             =   330
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFHasta 
         Height          =   300
         Left            =   1710
         TabIndex        =   10
         Top             =   765
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
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
         Index           =   0
         Left            =   510
         TabIndex        =   12
         Top             =   330
         Width           =   1140
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
         Height          =   300
         Index           =   1
         Left            =   510
         TabIndex        =   11
         Top             =   750
         Width           =   1140
      End
   End
   Begin VB.CommandButton botExcelAep 
      Caption         =   "Excel"
      Height          =   510
      Left            =   7350
      Picture         =   "frmMonedas.frx":02E4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4350
      Width           =   765
   End
   Begin VB.Frame Frame3 
      Height          =   1245
      Left            =   4770
      TabIndex        =   0
      Top             =   -15
      Width           =   2430
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
         Height          =   510
         Index           =   0
         Left            =   120
         Picture         =   "frmMonedas.frx":0876
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   375
         Width           =   630
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
         Height          =   510
         Index           =   1
         Left            =   870
         Picture         =   "frmMonedas.frx":0978
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   375
         Width           =   630
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
         Height          =   510
         Index           =   2
         Left            =   1620
         Picture         =   "frmMonedas.frx":0A7A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   375
         Width           =   630
      End
   End
   Begin TabDlg.SSTab tabTotal 
      Height          =   3555
      Left            =   90
      TabIndex        =   1
      Top             =   1305
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   6271
      _Version        =   327680
      TabOrientation  =   1
      TabHeight       =   441
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "TOTALES"
      TabPicture(0)   =   "frmMonedas.frx":129C
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "sprPago(2)"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "LLEGADAS"
      TabPicture(1)   =   "frmMonedas.frx":12B8
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sprPago(0)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "SALIDAS"
      TabPicture(2)   =   "frmMonedas.frx":12D4
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "sprPago(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Begin FPSpread.vaSpread sprPago 
         Height          =   2955
         Index           =   2
         Left            =   165
         OleObjectBlob   =   "frmMonedas.frx":12F0
         TabIndex        =   15
         Top             =   120
         Width           =   6765
      End
      Begin FPSpread.vaSpread sprPago 
         Height          =   2955
         Index           =   1
         Left            =   -74850
         OleObjectBlob   =   "frmMonedas.frx":1EF5
         TabIndex        =   14
         Top             =   120
         Width           =   6750
      End
      Begin FPSpread.vaSpread sprPago 
         Height          =   2955
         Index           =   0
         Left            =   -74835
         OleObjectBlob   =   "frmMonedas.frx":2AF9
         TabIndex        =   13
         Top             =   120
         Width           =   6750
      End
   End
End
Attribute VB_Name = "frmMonedas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsPagos  As Recordset


Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset

On Error GoTo ErrTot:
frmMonedas.caption = Aplicacion.SeteoProceso(frmMonedas.caption)



    sql = " Select tipo_pago,tipo_loc,nacionalidad nac "
    sql = sql & " , decode(tipo_pago,1,cod_moneda,decode(cod_moneda,1,1,16,1,23,1,2,2,18,2,25,2,3,3,19,3,26,3,4,4,17,4,24,4,98)) tipo_moneda"
    sql = sql & " ,sum(importe) importe "
    sql = sql & " From estadis.pago_moneda "
    sql = sql & " where fch_ticket between " & func_ToDate(Format(mskFDesde.FormattedText, FTOFECHA))
    sql = sql & " and " & func_ToDate(Format(mskFHasta.FormattedText, FTOFECHA))
    sql = sql & " group by tipo_pago,tipo_loc,nacionalidad,decode(tipo_pago,1,cod_moneda,decode(cod_moneda,1,1,16,1,23,1,2,2,18,2,25,2,3,3,19,3,26,3,4,4,17,4,24,4,98)) "
    sql = sql & " order by tipo_pago,decode(tipo_pago,1,cod_moneda,decode(cod_moneda,1,1,16,1,23,1,2,2,18,2,25,2,3,3,19,3,26,3,4,4,17,4,24,4,98)),tipo_loc, nacionalidad "
    
If Aplicacion.ObtenerRsDAO(sql, RsPagos) Then
    
    If Aplicacion.CantReg(RsPagos) > 0 Then
        L_LlenarGrillas
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabTotal.Enabled = True
    End If

    Aplicacion.CerrarDAO RsPagos
    
End If

ErrTot:
    frmMonedas.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub
Private Sub botEjecutar_Click(Index As Integer)

Select Case Index
    Case 0
        
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabTotal.Enabled = True
        L_Refrescar

    Case 1
        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        tabTotal.Enabled = False
        L_LimpiarGrillas
    Case 2
        Unload Me
End Select

End Sub
Private Sub L_SeteoTabPax(valor As Boolean)
    tabTotal.TabEnabled(2) = valor
    tabTotal.TabEnabled(4) = valor
    tabTotal.TabEnabled(1) = valor
    tabTotal.TabEnabled(3) = valor
End Sub



Private Sub L_LimpiarGrillas()
Dim i
    
For i = 0 To 2
    sprPago(i).MaxRows = 2
Next

End Sub

Private Sub L_DescripMonedas()
Dim sql As String
Dim rs As Recordset
Dim i

sql = "Select cod_moneda,descrip From ventas.monedas where cod_moneda not in (14,0) order by cod_moneda"
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        For i = 0 To 2
            sprPago(i).MaxRows = sprPago(i).MaxRows + 1
            
            sprPago(i).SetText 1, sprPago(i).MaxRows, str(rs!cod_moneda)
            sprPago(i).SetText 2, sprPago(i).MaxRows, Trim(rs!Descrip)
        Next
        rs.MoveNext
    Loop
    'Spread_TotalesGrillas sprPago, sprPago.MaxCols - 2, 2
    For i = 0 To 2
        sprPago(i).MaxRows = sprPago(i).MaxRows + 1
        Spread_TotalesLinea sprPago(i)
    
        sprPago(i).SetText 1, sprPago(i).MaxRows, "00"
        sprPago(i).SetText 2, sprPago(i).MaxRows, "Total Efectivo"
    Next
    Aplicacion.CerrarDAO rs
End If

sql = "Select cod_tarjeta,descrip From ventas.tarjetas Where cod_tarjeta in (1,2,3,4) ORDER BY cod_tarjeta"
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        For i = 0 To 2
        sprPago(i).MaxRows = sprPago(i).MaxRows + 1
        
        sprPago(i).SetText 1, sprPago(i).MaxRows, str(rs!cod_tarjeta)
        sprPago(i).SetText 2, sprPago(i).MaxRows, Trim(rs!Descrip)
        Next
        rs.MoveNext
    Loop
    For i = 0 To 2
        sprPago(i).MaxRows = sprPago(i).MaxRows + 1
          
        sprPago(i).SetText 1, sprPago(i).MaxRows, "98"
        sprPago(i).SetText 2, sprPago(i).MaxRows, "Otras"
        
        sprPago(i).MaxRows = sprPago(i).MaxRows + 1
        Spread_TotalesLinea sprPago(i)
    
        sprPago(i).SetText 1, sprPago(i).MaxRows, "99"
        sprPago(i).SetText 2, sprPago(i).MaxRows, "Total Tarjeta"
    Next
    Aplicacion.CerrarDAO rs
End If

End Sub
Private Sub L_LlenarGrillas()
Dim tipoPago As Integer
Dim tipoLocal As String, tipoMon As Integer, nacion As String
Dim moneda As Variant
Dim ind As Integer
Dim TotPagoSE As Single, TotPagoLE As Single, TotPagoSA As Single, TotPagoLA As Single
Dim PagoE As Single, PagoA As Single
    L_DescripMonedas
        
    ind = 2
    
    Do While Not RsPagos.EOF
    tipoPago = RsPagos!tipo_pago
        TotPagoSA = 0
        TotPagoLA = 0
        TotPagoSE = 0
        TotPagoLE = 0
        Do While tipoPago = RsPagos!tipo_pago And Val(moneda) <> 99
            tipoMon = RsPagos!tipo_moneda
            ind = ind + 1
            sprPago(0).GetText 1, ind, moneda
            PagoA = 0
            PagoE = 0
            Do While Val(moneda) <> 0 And Val(moneda) <> 99 And tipoPago = RsPagos!tipo_pago And tipoMon = RsPagos!tipo_moneda
                nacion = RsPagos!nac
                If Val(moneda) = tipoMon Then
                    Select Case RsPagos!tipo_loc
                        Case "L"
                            Select Case RsPagos!nac
                                Case "A"
                                    sprPago(0).SetText 3, ind, str(RsPagos!Importe)
                                    TotPagoLA = TotPagoLA + RsPagos!Importe
                                    PagoA = PagoA + RsPagos!Importe
                                Case "E"
                                    sprPago(0).SetText 6, ind, str(RsPagos!Importe)
                                    TotPagoLE = TotPagoLE + RsPagos!Importe
                                    PagoE = PagoE + RsPagos!Importe
                            End Select
                        Case "S"
                            Select Case RsPagos!nac
                                Case "A"
                                    sprPago(1).SetText 3, ind, str(RsPagos!Importe)
                                    TotPagoSA = TotPagoSA + RsPagos!Importe
                                    PagoA = PagoA + RsPagos!Importe
                                Case "E"
                                    sprPago(1).SetText 6, ind, str(RsPagos!Importe)
                                    TotPagoSE = TotPagoSE + RsPagos!Importe
                                    PagoE = PagoE + RsPagos!Importe
                            End Select
                    
                    End Select
                    
                    RsPagos.MoveNext
                    If RsPagos.EOF Then
                        Exit Do
                    End If
                Else
                    If RsPagos!tipo_moneda = 14 Then
                        RsPagos.MoveNext
                        tipoMon = RsPagos!tipo_moneda
                        If RsPagos.EOF Then
                            Exit Do
                        End If
                    Else
                        ind = ind + 1
                        sprPago(0).GetText 1, ind, moneda
                    End If
                End If
            Loop

            If Val(moneda) = 99 Or Val(moneda) = 0 Then
                If Not RsPagos.EOF Then
                    RsPagos.MoveNext
                End If
                Exit Do
            Else
                sprPago(2).SetText 3, ind, str(PagoA)
                sprPago(2).SetText 6, ind, str(PagoE)
                'ind = ind + 1
                'sprPago(0).GetText 1, ind, moneda
            End If
            If RsPagos.EOF Then
                Exit Do
            End If
        Loop
        Do While Val(moneda) <> 0 And Val(moneda) <> 99
            ind = ind + 1
            sprPago(0).GetText 1, ind, moneda
        Loop
        sprPago(0).SetText 3, ind, str(TotPagoLA)
        sprPago(0).SetText 6, ind, str(TotPagoLE)
        sprPago(1).SetText 3, ind, str(TotPagoSA)
        sprPago(1).SetText 6, ind, str(TotPagoSE)
            
        sprPago(2).SetText 3, ind, str(TotPagoLA + TotPagoSA)
        sprPago(2).SetText 6, ind, str(TotPagoLE + TotPagoSE)
        
        If Val(moneda) = 99 Then
            Exit Do
        End If
'        ind = ind + 1
    Loop
    
L_Participacion sprPago(2)
L_Participacion sprPago(0)
L_Participacion sprPago(1)

End Sub

Private Sub L_Participacion(spr As control)
Dim TotA00 As Variant, TotE00 As Variant, TotA99 As Variant, TotE99 As Variant
Dim Tot00 As Variant, Tot99 As Variant
Dim i, codigo As Variant, valor As Variant
Dim Tipo As String

For i = 3 To spr.MaxRows
    spr.GetText 1, i, codigo
    If codigo = "00" Then
        spr.GetText 3, i, TotA00
        spr.GetText 6, i, TotE00
        spr.GetText 9, i, Tot00
    ElseIf codigo = "99" Then
        spr.GetText 3, i, TotA99
        spr.GetText 6, i, TotE99
        spr.GetText 9, i, Tot99
    End If
Next

Tipo = "E"
For i = 3 To spr.MaxRows

    spr.GetText 1, i, codigo
    If Tipo = "E" Then
        spr.GetText 3, i, valor
        If Val(TotA00) > 0 Then
            spr.SetText 4, i, str(Val(valor) / Val(TotA00) * 100)
            spr.SetText 5, i, str(Val(valor) / Val(TotA00 + TotA99) * 100)
        End If
        spr.GetText 6, i, valor
        If Val(TotE00) > 0 Then
            spr.SetText 7, i, str(Val(valor) / Val(TotE00) * 100)
            spr.SetText 8, i, str(Val(valor) / Val(TotE00 + TotE99) * 100)
        End If
        spr.GetText 9, i, valor
        If Val(Tot00) > 0 Then
            spr.SetText 10, i, str(Val(valor) / Val(Tot00) * 100)
            spr.SetText 11, i, str(Val(valor) / Val(Tot00 + Tot99) * 100)
        End If
    
    Else
        spr.GetText 3, i, valor
        If Val(TotA99) > 0 Then
            spr.SetText 4, i, str(Val(valor) / Val(TotA99) * 100)
            spr.SetText 5, i, str(Val(valor) / Val(TotA99 + TotA00) * 100)
        End If
        spr.GetText 6, i, valor
        If Val(TotE99) > 0 Then
            spr.SetText 7, i, str(Val(valor) / Val(TotE99) * 100)
            spr.SetText 8, i, str(Val(valor) / Val(TotE99 + TotE00) * 100)
        End If
        spr.GetText 9, i, valor
        If Val(Tot99) > 0 Then
            spr.SetText 10, i, str(Val(valor) / Val(Tot99) * 100)
            spr.SetText 11, i, str(Val(valor) / Val(Tot99 + Tot00) * 100)
        End If
    
    End If
    If codigo = "00" Then
        Tipo = "T"
    End If
    
Next
End Sub


Private Sub botExcelAep_Click()
'    L_TratarExcel "TOTAL COMPAÑIA ", "Informe por Local (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")"
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


Private Sub Form_Load()
Me.Left = 50
Me.Top = 100
Me.Height = 5800
Me.Width = 9000

mskFDesde.Text = func_Dia1SegunMes_Anio(Month(Date), Year(Date))
If CDate(mskFDesde.FormattedText) = Date Then
    mskFHasta.Text = Format$(Date, FTOFECHA)
Else
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
End If


L_LimpiarGrillas
' frmPrincipal.lstForms.AddItem "frmTot"

End Sub


Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmTot"
End Sub


Private Sub mskFDesde_LostFocus()
    
    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    End If
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
    
    If mskFHasta.Text <> "" Then
        If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
            mskFHasta.Text = mskFDesde.Text
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

End Sub


