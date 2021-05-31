VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmInterfaceSAP 
   Caption         =   "Generación de interface SAP"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   5340
   Begin VB.CommandButton botCancelar 
      Caption         =   "Salir"
      Height          =   390
      Left            =   3390
      TabIndex        =   10
      Top             =   2295
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   2880
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   5250
      Begin VB.OptionButton opt 
         Caption         =   "Ventas Rubro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3450
         TabIndex        =   17
         Top             =   705
         Width           =   1560
      End
      Begin VB.OptionButton opt 
         Caption         =   "Ventas Gral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3450
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.CheckBox chkDepn 
         Caption         =   "BARI"
         Height          =   210
         Index           =   4
         Left            =   3705
         TabIndex        =   15
         Top             =   1245
         Value           =   1  'Checked
         Width           =   705
      End
      Begin VB.CheckBox chkDepn 
         Caption         =   "MEND"
         Height          =   210
         Index           =   3
         Left            =   2700
         TabIndex        =   14
         Top             =   1245
         Value           =   1  'Checked
         Width           =   810
      End
      Begin VB.CheckBox chkDepn 
         Caption         =   "CORD"
         Height          =   210
         Index           =   2
         Left            =   1800
         TabIndex        =   13
         Top             =   1245
         Value           =   1  'Checked
         Width           =   780
      End
      Begin VB.CheckBox chkDepn 
         Caption         =   "AEP"
         Height          =   210
         Index           =   1
         Left            =   1035
         TabIndex        =   12
         Top             =   1245
         Value           =   1  'Checked
         Width           =   630
      End
      Begin VB.CheckBox chkDepn 
         Caption         =   "EZE"
         Height          =   210
         Index           =   0
         Left            =   255
         TabIndex        =   11
         Top             =   1245
         Value           =   1  'Checked
         Width           =   630
      End
      Begin VB.CommandButton botAceptar 
         Caption         =   "Aceptar"
         Height          =   390
         Left            =   315
         TabIndex        =   9
         Top             =   2220
         Width           =   1380
      End
      Begin VB.TextBox txtArchivo 
         Height          =   300
         Left            =   1515
         TabIndex        =   8
         Top             =   1635
         Width           =   3510
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2670
         Picture         =   "frmInterfaceSAP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   2655
         Picture         =   "frmInterfaceSAP.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   675
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1515
         TabIndex        =   3
         Top             =   330
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
         Left            =   1515
         TabIndex        =   4
         Top             =   675
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
         Caption         =   "Archivo"
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
         Left            =   255
         TabIndex        =   7
         Top             =   1650
         Width           =   1200
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
         Left            =   270
         TabIndex        =   6
         Top             =   330
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
         Index           =   1
         Left            =   255
         TabIndex        =   5
         Top             =   675
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmInterfaceSAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function L_ControlVentas() As Boolean
Dim Result As Integer
Dim sqlSAP As String
Dim rsDat As Recordset
Dim fecha As String
Dim cDepn As String, cDif As Integer
Dim mensaje As String

Result = True

sqlSAP = " select fch_ticket, depn,sum(impsap)-sum(impvta)dif "
sqlSAP = sqlSAP & "from( "
sqlSAP = sqlSAP & " select fch_ticket,decode(depn,'INT',sdepn,depn) depn,sum(importe) Impsap,0 impvta "
sqlSAP = sqlSAP & " From estadis.interface_ventas_sap "
sqlSAP = sqlSAP & " where fch_ticket between " & Func.func_ToDate(mskFDesde.FormattedText)
sqlSAP = sqlSAP & " And " & Func.func_ToDate(mskFHasta.FormattedText)
sqlSAP = sqlSAP & " group by fch_ticket,decode(depn,'INT',sdepn,depn)"
sqlSAP = sqlSAP & " Union "
sqlSAP = sqlSAP & " select fch_ticket,decode (cod_depn,'INT',cod_sdep,cod_depn) ,0 impsap, sum(importe) impVta "
sqlSAP = sqlSAP & " From Ventas.ticket_h_mes "
sqlSAP = sqlSAP & " where fch_ticket Between " & Func.func_ToDate(mskFDesde.FormattedText)
sqlSAP = sqlSAP & " And " & Func.func_ToDate(mskFHasta.FormattedText)
sqlSAP = sqlSAP & " And nro_impresora <> 0 "
sqlSAP = sqlSAP & " group by fch_ticket,decode (cod_depn,'INT',cod_sdep,cod_depn)) "
sqlSAP = sqlSAP & " Group by fch_ticket,depn "

fecha = mskFDesde.FormattedText

mensaje = ""

If Aplicacion.ObtenerRsDAO(sqlSAP, rsDat) Then
    Do While CDate(fecha) <= CDate(mskFHasta.FormattedText)
        
       cDepn = ""
       If Not rsDat.EOF Then
       Do While CDate(fecha) = rsDat!Fch_Ticket
          cDepn = cDepn & rsDat!depn
          
          If rsDat!dif <> 0 Then
             mensaje = mensaje & "Diferencias : " & fecha & " " & rsDat!depn & Chr(10) & Chr(13)
          End If
          
          rsDat.MoveNext
          If rsDat.EOF Then
            Exit Do
          End If
       Loop
       End If
       If cDepn = "" Then
            mensaje = mensaje & "Faltante : completo " & fecha & Chr(10) & Chr(13)
       Else
          If InStr(1, cDepn, "EZE") = 0 Then
             mensaje = mensaje & "Faltante : " & fecha & " EZE " & Chr(10) & Chr(13)
          End If
          If InStr(1, cDepn, "AEP") = 0 Then
             mensaje = mensaje & "Faltante : " & fecha & " AEP " & Chr(10) & Chr(13)
          End If
          If InStr(1, cDepn, "CORD") = 0 Then
             mensaje = mensaje & "Faltante : " & fecha & " CORD " & Chr(10) & Chr(13)
          End If
          If InStr(1, cDepn, "MEND") = 0 Then
             mensaje = mensaje & "Faltante : " & fecha & " MEND " & Chr(10) & Chr(13)
          End If
       End If
       fecha = Format(CDate(fecha) + 1, FTOFECHA)
    Loop
    
End If

If mensaje <> "" Then
If MsgBox(mensaje, vbYesNo, "Resultados del control ") = vbYes Then
    L_ControlVentas = True
Else
    L_ControlVentas = False
End If
Else
    L_ControlVentas = True
End If
End Function

Private Function L_Dependencia()
Dim cadena As String

cadena = ""
For i = 0 To 4
    If chkDepn(i).Value = 1 Then
        cadena = cadena & chkDepn(i).caption
    End If
Next
L_Dependencia = cadena
End Function

Private Sub L_GenerarInterface()
Dim sql As String
Dim rs As Recordset

sql = " SELECT TO_CHAR(FCH_TICKET,'DDMMYYYY')||"
sql = sql & " LTRIM(TO_CHAR(NRO_IMPRESORA,'0009'))||LTRIM(DECODE(GREATEST(IMPORTE,0),IMPORTE,'-','B'))||"
sql = sql & " LTRIM(TO_CHAR(FF_DESDE,'00000099'))||"
sql = sql & " LTRIM(TO_CHAR(NRO_IMPRESORA,'0009'))||LTRIM(DECODE(GREATEST(IMPORTE,0),IMPORTE,'-','B'))||"
sql = sql & " LTRIM(TO_CHAR(FF_HASTA,'00000099'))||"
sql = sql & " LTRIM(TO_CHAR(NRO_IMPRESORA,'0009'))||"
'sql = sql & " decode(sdepn,'INTA','Z','INTB','Z','AEP','Y','CORD','X','MEND','W','BARI','T')||LTRIM(TO_CHAR(CAJERO,'009'))||"
sql = sql & " DECODE(importe,0,'',decode(sdepn,'INTA','Z','INTB','Z','AEP','Y','CORD','X','MEND','W','BARI','T'))||DECODE(importe,0,'9999',LTRIM(TO_CHAR(CAJERO,'009')))||"
'sql = sql & " TO_CHAR(DECODE(IMPORTE,0,0.01,DECODE(GREATEST(IMPORTE,0),IMPORTE,IMPORTE,IMPORTE*-1)),'999999999990.00')||"
sql = sql & " TO_CHAR(DECODE(IMPORTE,0,0.01,DECODE(GREATEST(IMPORTE,0),IMPORTE,IMPORTE+INTERES,(IMPORTE+INTERES)*-1)),'999999999990.00')||"
sql = sql & " TO_CHAR(CENTR_COSTO)||"
sql = sql & " DECODE(GREATEST(IMPORTE,0),IMPORTE,' ','*')||"
sql = sql & " DECODE(TIPO,'M','TF','  ')||"
sql = sql & " lpad(nvl(datos_pax,' '),59,' ')||"
sql = sql & " TO_CHAR(INTERES,'9999999999990.00')||"
'sql = sql & " TO_CHAR(DECODE(IMPORTE,0,0.01,DECODE(GREATEST(IMPORTE,0),IMPORTE,IMPORTE-INTERES,(IMPORTE-INTERES)*-1)),'99999990.00') campo "
sql = sql & " TO_CHAR(DECODE(IMPORTE,0,0.01,DECODE(GREATEST(IMPORTE,0),IMPORTE,IMPORTE,(IMPORTE)*-1)),'99999990.00') campo "
sql = sql & " From estadis.INTERFACE_VENTAS_SAP"
sql = sql & " where fch_ticket between " & Func.func_ToDate(mskFDesde.FormattedText)
sql = sql & " And " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " And instr('" & L_Dependencia & "',decode(depn,'INT',sdepn,depn),1) <> 0 "
sql = sql & " ORDER BY FCH_TICKET,SDEPN,NRO_IMPRESORA,CAJERO"

If txtArchivo.Text <> "" Then
If Aplicacion.ObtenerRsDAO(sql, rs) Then

Open txtArchivo.Text For Output Access Write Lock Read Write As #3    ' Abre el archivo.
    Do While Not rs.EOF
        Print #3, rs!campo
        rs.MoveNext
    Loop
Close #3
    MsgBox "Proceso Concluido", vbOKOnly, "Atención"
End If
Else
    MsgBox "Debe introducir el nombre de archivo de salida", vbOKOnly, "Atención"
End If

End Sub

Private Sub L_GenerarInterfaceRubro()
Dim sql As String
Dim rs As Recordset

sql = " Select '" & Format(mskFHasta.FormattedText, "DDMMYYYY") & "'||centro_costo||cod_rubr||to_char(sum(importe),'99999999.00')||to_char(sum(importe * ventas.obt_cotizacion_fecha(1,fch_ticket)),'99999999.00') campo"
sql = sql & " From ( "
sql = sql & " select fch_ticket,centro_costo,p.cod_rubr,sum(importe) importe "
sql = sql & " from  ventas.ticket_d_mes d , baires.local , baires.producto p , baires.comitente c, baires.prodcomi pd "
sql = sql & " where fch_ticket between " & Func.func_ToDate(mskFDesde.FormattedText)
sql = sql & " And " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " and cod_local = cod_loc and d.cod_prod = p.cod_prod "
sql = sql & " and pd.cod_prod = p.cod_prod and pd.cod_comi = c.cod_comi And modo = 'V' "
sql = sql & " And instr('" & L_Dependencia & "',decode(cod_depn,'INT',cod_sdep,cod_depn ),1) <> 0 "
sql = sql & " group by fch_ticket,centro_costo,p.cod_rubr "
sql = sql & " ) group by centro_costo,cod_rubr "

If txtArchivo.Text <> "" Then
If Aplicacion.ObtenerRsDAO(sql, rs) Then

Open txtArchivo.Text For Output Access Write Lock Read Write As #3    ' Abre el archivo.
    Do While Not rs.EOF
        Print #3, rs!campo
        rs.MoveNext
    Loop
Close #3
    MsgBox "Proceso Concluido", vbOKOnly, "Atención"
End If
Else
    MsgBox "Debe introducir el nombre de archivo de salida", vbOKOnly, "Atención"
End If

End Sub


Private Sub botAceptar_Click()

frmInterfaceSAP.caption = Aplicacion.SeteoProceso(frmInterfaceSAP.caption)
If opt(0).Value = True Then
    If L_ControlVentas Then
          L_GenerarInterface
    End If
Else
          L_GenerarInterfaceRubro
End If
frmInterfaceSAP.caption = Aplicacion.SeteoFin
End Sub



Private Sub botCancelar_Click()
Unload Me
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
Top = 1000
Left = 1000
Me.Height = 3500
Me.Width = 5500
mskFDesde.Text = Format(Now - 1, FTOFECHA)
mskFHasta.Text = Format(Now - 1, FTOFECHA)

txtArchivo.Text = "H:\SAP\INTERFAC.ES\FACTURAC\FACT_SAP"
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
ElseIf CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Or Year(mskFHasta.FormattedText) <> Year(mskFDesde.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
End If

mskFHasta.Text = Format$(mskFHasta.FormattedText, FTOFECHA)

End Sub


Private Sub txtArchivo_DblClick()
    txtArchivo.Text = frmDir.NombreArchivo()
End Sub


