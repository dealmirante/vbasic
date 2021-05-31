VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmControlTicket 
   Caption         =   "Control de ticket"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   9720
   Begin TabDlg.SSTab SSTab1 
      Height          =   3585
      Left            =   270
      TabIndex        =   18
      Top             =   1350
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6324
      _Version        =   327680
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tickets por CAJEROS"
      TabPicture(0)   =   "FrmControlTicket.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "sprTicket"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Venta por EMPLEADO"
      TabPicture(1)   =   "FrmControlTicket.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "sprLegajo"
      Tab(1).Control(0).Enabled=   0   'False
      Begin FPSpread.vaSpread sprLegajo 
         Height          =   2985
         Left            =   135
         OleObjectBlob   =   "FrmControlTicket.frx":0038
         TabIndex        =   20
         Top             =   435
         Width           =   8985
      End
      Begin FPSpread.vaSpread sprTicket 
         Height          =   2985
         Left            =   -74880
         OleObjectBlob   =   "FrmControlTicket.frx":03B1
         TabIndex        =   19
         Top             =   435
         Width           =   9030
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   360
      Left            =   8925
      Top             =   1380
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Left            =   8625
      TabIndex        =   15
      Top             =   60
      Width           =   885
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
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "FrmControlTicket.frx":0788
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   165
         Width           =   645
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
         Height          =   495
         Index           =   2
         Left            =   120
         Picture         =   "FrmControlTicket.frx":088A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   690
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   210
      TabIndex        =   0
      Top             =   60
      Width           =   8355
      Begin VB.TextBox txtTmpA 
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
         Height          =   285
         Left            =   5175
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   330
         Width           =   960
      End
      Begin VB.TextBox txtTmpC 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   5175
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   750
         Width           =   960
      End
      Begin MSMask.MaskEdBox mskMonto 
         Height          =   300
         Left            =   1365
         TabIndex        =   10
         Top             =   750
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton optDia 
         Caption         =   "Otros Días"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2865
         TabIndex        =   3
         Top             =   690
         Width           =   1245
      End
      Begin VB.OptionButton optDia 
         Caption         =   "ON LINE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2865
         TabIndex        =   1
         Top             =   390
         Width           =   1215
      End
      Begin ComctlLib.Slider sldTmp 
         Height          =   210
         Left            =   6315
         TabIndex        =   4
         Top             =   585
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   370
         _Version        =   327682
         OLEDropMode     =   1
         LargeChange     =   1
         Min             =   1
         SelStart        =   5
         Value           =   5
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   285
         Left            =   1365
         TabIndex        =   8
         Top             =   345
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "Hora Actual "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4275
         TabIndex        =   14
         Top             =   300
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "Tiempo Consulta"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4275
         TabIndex        =   13
         Top             =   705
         Width           =   930
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto tope"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         TabIndex        =   9
         Top             =   765
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "consulta cada..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6390
         TabIndex        =   7
         Top             =   375
         Width           =   1440
      End
      Begin VB.Label labMin 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "5"
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
         Height          =   270
         Left            =   6405
         TabIndex        =   6
         Top             =   795
         Width           =   330
      End
      Begin VB.Label Label5 
         Caption         =   "minutos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6930
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.Label LblFecha 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   2
         Top             =   345
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmControlTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vTabla As String
Dim vTablaD As String

Function l_buscarTicket(nro_ticket As Integer, cod_caja As Integer) As Variant
Dim tick As Variant, caj As Variant, items As Variant

l_buscarTicket = 1
For i = 1 To sprTicket.MaxRows
    sprTicket.GetText 1, i, tick
    sprTicket.GetText 3, i, caj
    
    If tick = nro_ticket And caj = cod_caja Then
       sprTicket.GetText 7, i, items
       l_buscarTicket = items
       Exit For
    End If
Next

End Function


Private Sub L_Refrescar()
Dim sql As String
Dim rsData As Recordset

If mskMonto.Text <> "" Then
    sql = " select nro_ticket,cod_local,cod_caja,c.cod_cajero"
    sql = sql & " ,c.legajo||'- '||Apellido,descrip,cant_items,to_char(importe,'999,999.99')"
    sql = sql & " from " & vTabla
    sql = sql & " ,ventas.cajeros c,"
    sql = sql & " personal.empleado e,"
    sql = sql & " Ventas.paises P"
    sql = sql & " Where"
    sql = sql & " h.cod_cajero  = c.cod_cajero (+)"
    sql = sql & " and c.legajo  = e.legajo (+)"
    sql = sql & " and cod_pais  = cod_nac"
    sql = sql & " and fch_ticket= " & Func.func_ToDate(mskFecha.FormattedText)
    sql = sql & " and importe >= " & mskMonto.Text
    sql = sql & " Order by importe desc"

    Spread.Spead_VaciarGrilla sprTicket

    If Aplicacion.ObtenerRsDAO(sql, rsData) Then
        Spread.Spread_CargarGrilla rsData, sprTicket
        rsData.Close
    End If

    'sql = " Select vendedor,Apellido||', '||nombre,sum(cantidad),to_char(sum(importe),'999,999.00')"
    'sql = sql & " from  " & vTablaD
    'sql = sql & " ,personal.empleado"
    'sql = sql & " Where vendedor = Legajo"
    'sql = sql & " group by vendedor,Apellido||', '||nombre"
    'sql = sql & " Having Sum(Importe) > 500"
    'sql = sql & " order by sum(importe) desc"

    sql = " Select nro_ticket,cod_local,cod_caja,vendedor,Apellido||', '||nombre ape,sum(cantidad) cant,to_char(sum(importe),'999,999.00') imp "
    sql = sql & " from  " & vTablaD
    sql = sql & " ,personal.empleado"
    sql = sql & " Where vendedor = Legajo"
    sql = sql & " and fch_ticket= " & Func.func_ToDate(mskFecha.FormattedText)
    sql = sql & " group by nro_ticket,cod_local,cod_caja,vendedor,Apellido||', '||nombre"
    sql = sql & " Having Sum(Importe) > " & mskMonto.Text
    sql = sql & " order by sum(importe) desc"

    Spread.Spead_VaciarGrilla sprLegajo

    If Aplicacion.ObtenerRsDAO(sql, rsData) Then
        'Spread.Spread_CargarGrilla rsData, sprLegajo
        Do While Not rsData.EOF
        sprLegajo.MaxRows = sprLegajo.MaxRows + 1
        sprLegajo.SetText 1, sprLegajo.MaxRows, str(rsData!nro_ticket)
        sprLegajo.SetText 2, sprLegajo.MaxRows, Trim(rsData!cod_local)
        sprLegajo.SetText 3, sprLegajo.MaxRows, str(rsData!cod_caja)
        sprLegajo.SetText 4, sprLegajo.MaxRows, str(rsData!Vendedor)
        sprLegajo.SetText 5, sprLegajo.MaxRows, Trim(rsData!ape)
        sprLegajo.SetText 6, sprLegajo.MaxRows, str(rsData!cant) & "-" & Format((100 * rsData!cant / l_buscarTicket(rsData!nro_ticket, rsData!cod_caja)), "###") & "%"
        sprLegajo.SetText 7, sprLegajo.MaxRows, Trim(rsData!imp)
                
        rsData.MoveNext
        Loop
        rsData.Close
    End If

Else
   MsgBox "Debe ingresar monto y/o fecha", vbOKOnly + vbCritical, "Error"
End If

End Sub
Private Sub botEjecutar_Click(Index As Integer)


Select Case Index
    Case 0
        'L_LimpiarGrillas
        txtTmpC.Text = Format$(Now, "hh:mm:ss")
        L_Refrescar
        'txtTmpC.Text = Format$(Now, "hh:mm:ss")
    Case 2
        Unload Me
End Select

End Sub

Private Sub Form_Activate()
'Call botEjecutar_Click(0)
End Sub

Private Sub Form_Load()
Top = 50
Left = 100
Width = 10000

mskMonto.Text = 500
txtTmpC.Text = Format$(Now, "hh:mm:ss")

vTabla = "ticket_h_dia@link_nt_vta h"
vTablaD = "ticket_d_dia@link_nt_vta d"

mskFecha.Text = Format(Now, "dd/mm/yyyy")


End Sub

Private Sub mskFecha_Change()
sprTicket.MaxRows = 0
sprLegajo.MaxRows = 0
End Sub

Private Sub mskFecha_LostFocus()
If mskFecha.Text = "" Then
   mskFecha.Text = Format(Now - 1, "dd/mm/yyyy")
End If
End Sub


Private Sub mskMonto_Change()
sprTicket.MaxRows = 0
sprLegajo.MaxRows = 0
End Sub

Private Sub mskMonto_LostFocus()
If mskMonto.Text = "" Then
   mskMonto.Text = 500
End If
End Sub


Private Sub optDia_Click(Index As Integer)
sprTicket.MaxRows = 0
sprLegajo.MaxRows = 0
If Index = 0 Then
   Label2.Visible = True
   Label4.Visible = True
   Label3.Visible = True
   Label5.Visible = True
   sldTmp.Visible = True
   labMin.Visible = True
   mskFecha.Text = Format(Now, "dd-mm-yyyy")
   mskFecha.Enabled = False
   txtTmpA.Visible = True
   txtTmpC.Visible = True
   
   Timer1.Enabled = True
   vTabla = "tickdia.ticket_h_dia@link_nt_vta h"
   vTablaD = "tickdia.ticket_d_dia@link_nt_vta d"
Else
   Label2.Visible = False
   Label4.Visible = False
   Label3.Visible = False
   Label5.Visible = False
   sldTmp.Visible = False
   labMin.Visible = False
   mskFecha.Enabled = True
   txtTmpA.Visible = False
   txtTmpC.Visible = False
   
   Timer1.Enabled = False
   vTabla = "ventas.ticket_h h"
   vTablaD = "ventas.ticket_d d"
End If

End Sub


Private Sub sldTmp_Click()
labMin.caption = sldTmp.Value
End Sub


Private Sub sprLegajo_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim sql As String
Dim Vendedor As Variant



'sql = " Select nro_ticket,cod_local,to_char(p.cod_prod) cp,descrip,sum(cantidad),to_char(sum(importe),'999,999.00') "
'sql = sql & " From baires.producto p, " & vTablaD
'sql = sql & " Where p.cod_prod = d.cod_prod "
'sql = sql & " And fch_ticket = " & Func.func_ToDate(mskFecha.FormattedText)
'sql = sql & " And vendedor = " & Vendedor
'sql = sql & " Group by  nro_ticket,cod_local,to_char(p.cod_prod) ,descrip "

sprLegajo.GetText 4, Row, Vendedor
sprLegajo.GetText 1, Row, ticket
sprLegajo.GetText 3, Row, caja

sql = " Select to_char(p.cod_prod) cp,descrip,cantidad,to_char(importe,'999,999.00') importe,vendedor "
sql = sql & " From baires.producto p, " & vTablaD
sql = sql & " Where p.cod_prod = d.cod_prod "
sql = sql & " And fch_ticket = " & Func.func_ToDate(mskFecha.FormattedText)
sql = sql & " And nro_ticket = " & ticket
sql = sql & " And cod_caja   = " & caja
sql = sql & " UNION all "
sql = sql & " Select ' ' cp,'TOTAL',sum(cantidad),to_char(sum(importe),'999,999.00'), 0 "
sql = sql & " From baires.producto p, " & vTablaD
sql = sql & " Where p.cod_prod = d.cod_prod "
sql = sql & " And fch_ticket = " & Func.func_ToDate(mskFecha.FormattedText)
sql = sql & " And nro_ticket = " & ticket
sql = sql & " And cod_caja   = " & caja
sql = sql & " Order by cp desc "

Call frmGridTicketD.DatosGrilla(sql, Int(Vendedor))

End Sub


Private Sub sprTicket_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim sql As String
Dim ticket As Variant
Dim caja As Variant

sprTicket.GetText 1, Row, ticket
sprTicket.GetText 3, Row, caja

sql = " Select to_char(p.cod_prod) cp,descrip,cantidad,to_char(importe,'999,999.00')importe,vendedor "
sql = sql & " From baires.producto p, " & vTablaD
sql = sql & " Where p.cod_prod = d.cod_prod "
sql = sql & " And fch_ticket = " & Func.func_ToDate(mskFecha.FormattedText)
sql = sql & " And nro_ticket = " & ticket
sql = sql & " And cod_caja   = " & caja
sql = sql & " UNION all "
sql = sql & " Select ' ' cp,'TOTAL',sum(cantidad),to_char(sum(importe),'999,999.00'), 0 "
sql = sql & " From baires.producto p, " & vTablaD
sql = sql & " Where p.cod_prod = d.cod_prod "
sql = sql & " And fch_ticket = " & Func.func_ToDate(mskFecha.FormattedText)
sql = sql & " And nro_ticket = " & ticket
sql = sql & " And cod_caja   = " & caja
sql = sql & " Order by cp desc "

Call frmGridTicketD.DatosGrilla(sql, -1)

End Sub


Private Sub Timer1_Timer()
txtTmpA.Text = Format$(Now, "hh:mm:ss")

    If ((Format$(txtTmpA.Text, "hhmmss")) - (Format$(txtTmpC.Text, "hhmmss"))) > (sldTmp.Value * 100) Then
        Call botEjecutar_Click(0)
    End If

End Sub


