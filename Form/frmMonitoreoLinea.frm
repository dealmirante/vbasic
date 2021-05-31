VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMonitoreoLinea 
   Caption         =   "Estimados vs Ventas-Tickets-Pasajeros"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   10995
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
      Left            =   9345
      Picture         =   "frmMonitoreoLinea.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   645
      Width           =   615
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
      Height          =   480
      Index           =   1
      Left            =   9345
      Picture         =   "frmMonitoreoLinea.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   165
      Width           =   615
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
      Height          =   480
      Index           =   0
      Left            =   10020
      Picture         =   "frmMonitoreoLinea.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   165
      Width           =   615
   End
   Begin VB.CommandButton botExcel 
      Caption         =   "Excel"
      Height          =   525
      Left            =   10035
      Picture         =   "frmMonitoreoLinea.frx":0A26
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   645
      Width           =   615
   End
   Begin VB.Frame frCab 
      Caption         =   "Datos de Cabecera"
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
      Height          =   1395
      Left            =   45
      TabIndex        =   0
      Top             =   -30
      Width           =   10755
      Begin VB.OptionButton optCota 
         Caption         =   "Todos"
         Height          =   210
         Index           =   2
         Left            =   6480
         TabIndex        =   21
         Top             =   960
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.OptionButton optCota 
         Caption         =   "Por Arriba"
         Height          =   210
         Index           =   1
         Left            =   5460
         TabIndex        =   20
         Top             =   945
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.OptionButton optCota 
         Caption         =   "Por Debajo"
         Height          =   210
         Index           =   0
         Left            =   4320
         TabIndex        =   19
         Top             =   945
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.ListBox lstProv 
         Height          =   255
         Left            =   6615
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox cboProv 
         Height          =   315
         Left            =   6675
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   300
         Width           =   2415
      End
      Begin VB.ComboBox CboRubro 
         Height          =   315
         ItemData        =   "frmMonitoreoLinea.frx":0FB8
         Left            =   3030
         List            =   "frmMonitoreoLinea.frx":0FBA
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   2415
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskMes 
         Height          =   300
         Left            =   960
         TabIndex        =   3
         Top             =   750
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskCota 
         Height          =   300
         Left            =   3600
         TabIndex        =   14
         Top             =   885
         Visible         =   0   'False
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PROVEEDOR"
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
         Index           =   2
         Left            =   5685
         TabIndex        =   16
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COTA %"
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
         Index           =   1
         Left            =   2610
         TabIndex        =   15
         Top             =   885
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Año"
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
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Width           =   780
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mes"
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
         TabIndex        =   5
         Top             =   750
         Width           =   780
      End
      Begin VB.Label LblCodAeropuerto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUBRO"
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
         Index           =   0
         Left            =   1890
         TabIndex        =   4
         Top             =   315
         Width           =   1110
      End
   End
   Begin TabDlg.SSTab tabEst 
      Height          =   5115
      Left            =   15
      TabIndex        =   11
      Top             =   1395
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   9022
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   6
      TabHeight       =   441
      ForeColor       =   255
      TabCaption(0)   =   "TOTAL CIA"
      TabPicture(0)   =   "frmMonitoreoLinea.frx":0FBC
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tabGA"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tabTotal"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "EZEIZA"
      TabPicture(1)   =   "frmMonitoreoLinea.frx":0FD8
      Tab(1).ControlCount=   0
      Tab(1).ControlEnabled=   0   'False
      TabCaption(2)   =   "AEROPARQUE"
      TabPicture(2)   =   "frmMonitoreoLinea.frx":0FF4
      Tab(2).ControlCount=   0
      Tab(2).ControlEnabled=   0   'False
      TabCaption(3)   =   "INTERIOR"
      TabPicture(3)   =   "frmMonitoreoLinea.frx":1010
      Tab(3).ControlCount=   0
      Tab(3).ControlEnabled=   0   'False
      Begin TabDlg.SSTab tabTotal 
         Height          =   4590
         Left            =   105
         TabIndex        =   12
         Top             =   375
         Width           =   10740
         _ExtentX        =   18944
         _ExtentY        =   8096
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "REAL MES"
         TabPicture(0)   =   "frmMonitoreoLinea.frx":102C
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprReal(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "LANZAMIENTOS"
         TabPicture(1)   =   "frmMonitoreoLinea.frx":1048
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprReal(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprReal 
            Height          =   4020
            Index           =   0
            Left            =   90
            OleObjectBlob   =   "frmMonitoreoLinea.frx":1064
            TabIndex        =   13
            Top             =   90
            Width           =   10575
         End
         Begin FPSpread.vaSpread sprReal 
            Height          =   4020
            Index           =   1
            Left            =   -74910
            OleObjectBlob   =   "frmMonitoreoLinea.frx":17E2
            TabIndex        =   23
            Top             =   75
            Width           =   10575
         End
      End
      Begin TabDlg.SSTab tabGA 
         Height          =   4605
         Left            =   225
         TabIndex        =   22
         Top             =   390
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   8123
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "REAL MES"
         TabPicture(0)   =   "frmMonitoreoLinea.frx":1F3F
         Tab(0).ControlCount=   0
         Tab(0).ControlEnabled=   -1  'True
         TabCaption(1)   =   "PROYECTADO"
         TabPicture(1)   =   "frmMonitoreoLinea.frx":1F5B
         Tab(1).ControlCount=   0
         Tab(1).ControlEnabled=   0   'False
      End
   End
End
Attribute VB_Name = "frmMonitoreoLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsEstim As Recordset
Dim rsEstimT As Recordset
Dim rs As Recordset
Dim rsPax As Recordset

Dim siImp As Boolean, siTic As Boolean, siPax As Boolean

Private Sub L_LimpiarGrillas()
Dim i

For i = 0 To 1
 sprReal(i).MaxRows = 0
 
 'lblProd(i).caption = ""
 'lblSTKBOD(i).caption = ""
Next
 
End Sub




Private Function L_NombreRubro(r As String)

Select Case r
   Case "ACC"
   L_NombreRubro = "ACCESORIOS"
   Case "BEB"
   L_NombreRubro = "BEBIDAS"
   Case "COS"
   L_NombreRubro = "COSMETICOS"
   Case "CIG"
   L_NombreRubro = "CIGARRILLOS"
   Case "COM"
   L_NombreRubro = "COMESTIBLES"
   Case "PER"
   L_NombreRubro = "PERFUMES"
   Case "TAB"
   L_NombreRubro = "TABACOS"
   Case Else
   L_NombreRubro = r
End Select

End Function

Private Function L_PasarGrillaExel(spr As control) As String
Dim i, j
Dim valor As Variant
Dim desc As String, sql As String

'sprEXEL.MaxRows = 0

For i = 1 To spr.MaxRows
    'sprEXEL.MaxRows = sprEXEL.MaxRows + 1
    
    For j = 1 To 9
    
      If j = 2 Then
          spr.GetText j, i, valor
          sql = "Select descrip From baires.producto where cod_prod = " & valor
          Func.Func_ObtenerDesc sql, desc
          'sprEXEL.SetText j, i, Trim(valor) & " - " & desc
      ElseIf j > 5 Then
         spr.GetText j, i, valor
         'sprEXEL.SetText j + 1, i, Trim(valor)
      Else
         spr.GetText j, i, valor
         'sprEXEL.SetText j, i, Trim(valor)
      End If
    Next
Next

End Function



Private Sub botEjecutar_Click(Index As Integer)

Select Case Index
    Case 0
        L_Refrescar
        L_RefrescarLANZ
    Case 1
        frCab.Enabled = True
        botEjecutar(0).Enabled = True
        tabEst.Enabled = False
        L_LimpiarGrillas
    Case 2
        Unload Me
End Select


End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlP As String

On Error GoTo ErrGLC:

frmMonitoreoLinea.caption = Aplicacion.SeteoProceso(frmMonitoreoLinea.caption)
sql = "  SELECT l.aniomes"
sql = sql & " , l.COD_PROV, L.COD_RUBR "
sql = sql & " , pv.descrip"
sql = sql & " , sum(importe) VtaDia"
sql = sql & " , round( ((ESTIMADO_$ - sum(importe)) / (last_day(sysdate-1) - sysdate +1) ) / ( sum(importe)/sum(cantidad) ) ) UniDia"
sql = sql & " , round((ESTIMADO_$ - sum(importe)) / (last_day(sysdate-1) - sysdate +1) ) ImpDia"
sql = sql & " , sum(importe) / to_char(sysdate-1,'dd') *  to_char(last_day(sysdate-1),'dd') Proy"
sql = sql & " , ESTIMADO_$ Est"
sql = sql & " , round(sum(importe)/sum(cantidad),2) "
sql = sql & " , round(sum(cantidad)/to_char(sysdate-1,'dd'))  PromUnid"
sql = sql & " FROM ESTADIS.MONITOREO_LINEA l,venta_mes v, baires.producto p, baires.proveedor pv"
sql = sql & " where l.aniomes = " & mskAnio.Text & Format(mskMes.Text, "00")
sql = sql & " and p.cod_prod = v.cod_prod"
sql = sql & " and p.cod_prov = l.cod_prov"
sql = sql & " and p.cod_rubr = l.cod_rubr"
sql = sql & " and l.aniomes  = v.aniomes"
sql = sql & " and p.cod_prov = pv.cod_prov"
sql = sql & "   AND DECODE (L.CONDICION,NULL,1,INSTR(L.CONDICION,P.COD_FAMILIA)) > 0"
sql = sql & L_Armarcondicion
sql = sql & " group by l.ANIOMES,L.COD_RUBR, l.COD_PROV,  ESTIMADO_$, pv.descrip"
sql = sql & " Having Sum(cantidad) <> 0 "
sql = sql & " ORDER BY L.COD_RUBR,ESTIMADO_$ DESC"


'    If optCota(2).Value Then
'       sql = sql & " Having Abs((Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$))) >= " & mskCota.Text / 100
'    ElseIf optCota(1).Value Then
'       sql = sql & " Having (Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$)) >= " & mskCota.Text / 100
'    ElseIf optCota(0).Value Then
'        sql = sql & " Having  (Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$)) <= - " & mskCota.Text / 100
'    End If

If Aplicacion.ObtenerRsDAO(sql, rs) Then
      L_LlenarGrillasReal
      tabEst.Enabled = True
      botEjecutar(0).Enabled = False
      frCab.Enabled = False
End If


ErrGLC:
    frmMonitoreoLinea.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub
Private Sub L_RefrescarLANZ()
Dim sql As String
Dim sqlP As String
    

On Error GoTo ErrGLC:

frmMonitoreoLinea.caption = Aplicacion.SeteoProceso(frmMonitoreoLinea.caption)
sql = "  SELECT l.aniomes"
sql = sql & " , P.COD_PROV, p.cod_rubr "
sql = sql & " , pv.descrip"
sql = sql & " , sum(cantidad) VtaDia "
sql = sql & " , round( ( SUM(OBJETIVO) - sum(CANTIDAD)) / (last_day(sysdate-1) - sysdate +1)  )  UniDia"
sql = sql & " , round( sum(CANTIDAD) / to_char(sysdate-1,'dd') *  to_char(last_day(sysdate-1),'dd')) Proy"
sql = sql & " , SUM(OBJETIVO) Est"
sql = sql & " , round(sum(cantidad)/to_char(sysdate-1,'dd'))  PromUnid"
sql = sql & " FROM ESTADIS.CODIGOS_LANZAMIENTOS l,venta_mes v   , baires.producto p, baires.proveedor pv "
sql = sql & " where l.aniomes = " & mskAnio.Text & Format(mskMes.Text, "00")
sql = sql & " and p.cod_prod = v.cod_prod"
sql = sql & " and p.cod_proD = l.cod_proD"
sql = sql & " and l.aniomes  = v.aniomes"
sql = sql & " and p.cod_prov = pv.cod_prov"
sql = sql & L_Armarcondicion
sql = sql & " group by l.ANIOMES, P.COD_PROV, p.cod_rubr ,  pv.descrip"
sql = sql & " ORDER BY p.COD_RUBR,EST desc"

'    If optCota(2).Value Then
'       sql = sql & " Having Abs((Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$))) >= " & mskCota.Text / 100
'    ElseIf optCota(1).Value Then
'       sql = sql & " Having (Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$)) >= " & mskCota.Text / 100
'    ElseIf optCota(0).Value Then
'        sql = sql & " Having  (Sum(venta_$)-Sum(ESTIMADO_$) ) / decode(Sum(ESTIMADO_$), 0, 1, Sum(ESTIMADO_$)) <= - " & mskCota.Text / 100
'    End If

If Aplicacion.ObtenerRsDAO(sql, rs) Then
      L_LlenarGrillasLANZ
      tabEst.Enabled = True
      botEjecutar(0).Enabled = False
      frCab.Enabled = False
End If


ErrGLC:
    frmMonitoreoLinea.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub

Private Function L_Armarcondicion()
Dim Cond

If cboRubro.Text <> "" Then
   Cond = Cond & " And P.cod_rubr = '" & cboRubro.Text & "'"
End If

If cboProv.Text <> "" Then
   Cond = Cond & " And pv.cod_prov = '" & lstProv.List(cboProv.ListIndex) & "'"
End If

L_Armarcondicion = Cond

End Function



Private Sub L_LlenarGrillasReal()
Dim rubro As String

    If Not rs.EOF Then
        rubro = rs!cod_rubr
        sprReal(0).MaxRows = sprReal(0).MaxRows + 1
        sprReal(0).SetText 1, sprReal(0).MaxRows, Trim(rubro)
    End If
    
    Do While Not rs.EOF
            sprReal(0).MaxRows = sprReal(0).MaxRows + 1
            
            If rs!cod_rubr <> rubro Then
               sprReal(0).MaxRows = sprReal(0).MaxRows + 1
               rubro = rs!cod_rubr
               sprReal(0).SetText 1, sprReal(0).MaxRows, Trim(rubro)
               sprReal(0).MaxRows = sprReal(0).MaxRows + 1
               
            End If
            
            sprReal(0).SetText 1, sprReal(0).MaxRows, Trim(rs!Descrip)
            sprReal(0).SetText 2, sprReal(0).MaxRows, str(rs!unidia)
            sprReal(0).SetText 3, sprReal(0).MaxRows, str(rs!impdia)
            sprReal(0).SetText 4, sprReal(0).MaxRows, Trim(rs!PromUnid)
            sprReal(0).SetText 5, sprReal(0).MaxRows, str(((rs!proy / rs!est) - 1) * 100)
            sprReal(0).SetText 6, sprReal(0).MaxRows, str(rs!est)
            sprReal(0).SetText 7, sprReal(0).MaxRows, str(rs!VtaDia)
            sprReal(0).SetText 8, sprReal(0).MaxRows, str(rs!est - rs!VtaDia)
            sprReal(0).SetText 9, sprReal(0).MaxRows, str(rs!proy)
            sprReal(0).SetText 10, sprReal(0).MaxRows, str(rs!VtaDia / rs!est * 100)
        rs.MoveNext
        If rs.EOF Then
            Exit Do
        End If
    Loop

End Sub

Private Sub L_LlenarGrillasLANZ()
Dim rubro As String

    If Not rs.EOF Then
        rubro = rs!cod_rubr
        sprReal(1).MaxRows = sprReal(1).MaxRows + 1
        sprReal(1).SetText 1, sprReal(1).MaxRows, Trim(rubro)
    End If

Do While Not rs.EOF
            sprReal(1).MaxRows = sprReal(1).MaxRows + 1
            If rs!cod_rubr <> rubro Then
               sprReal(1).MaxRows = sprReal(1).MaxRows + 1
               rubro = rs!cod_rubr
               sprReal(1).SetText 1, sprReal(1).MaxRows, Trim(rubro)
               sprReal(1).MaxRows = sprReal(1).MaxRows + 1
            End If
            sprReal(1).SetText 1, sprReal(1).MaxRows, Trim(rs!Descrip)
            sprReal(1).SetText 2, sprReal(1).MaxRows, str(rs!unidia)
            sprReal(1).SetText 3, sprReal(1).MaxRows, ""
            sprReal(1).SetText 4, sprReal(1).MaxRows, Trim(rs!PromUnid)
            sprReal(1).SetText 5, sprReal(1).MaxRows, str(((rs!proy / rs!est) - 1) * 100)
            sprReal(1).SetText 6, sprReal(1).MaxRows, str(rs!est)
            sprReal(1).SetText 7, sprReal(1).MaxRows, str(rs!VtaDia)
            sprReal(1).SetText 8, sprReal(1).MaxRows, str(rs!est - rs!VtaDia)
            sprReal(1).SetText 9, sprReal(1).MaxRows, str(rs!proy)
            sprReal(1).SetText 10, sprReal(1).MaxRows, str(rs!VtaDia / rs!est * 100)
                
        rs.MoveNext
        If rs.EOF Then
            Exit Do
        End If
    Loop

End Sub


Private Sub botExcel_Click()

    'L_TratarExcelTot tabEst.Tab
     L_PlanillaExcel "", ""
End Sub



Private Sub L_TratarExcel(Ap As Object, spr As control, Titulos As String)
Dim rango As String
Dim Col As Integer, c As Integer
Dim fila As Integer, f As Integer
Dim i As Integer
Dim tit As Variant, valor As Variant
Dim titCol() As String, tipo As String
Dim Varianza As Boolean
Dim LinAbajo As Integer, LinArriba As Integer

 On Error GoTo ErrorExl:


    ReDim titCol(spr.MaxCols)
    Col = 1
    fila = 3

    For i = 1 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next

    Exl_PonerValor Ap, 1, 1, Titulos
    
    rango = Exl_rangos(1, 1, 1, spr.MaxCols)
    
    Exl_Letra Ap, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion Ap, rango, Exl_Centro, Exl_CentroVert, False
    'AppExcel.Application.Range(rango).Merge
    Exl_Lineas Ap, rango, Exl_Linsimple
    Exl_ColorInt Ap, rango, Exl_Blanco
    
    If cboRubro.Text <> "" Then
        Exl_PonerValor Ap, fila, Col, "RUBRO : " & cboRubro.Text
    Else
        Exl_PonerValor Ap, fila, Col, "RUBRO : Todos "
    End If
    
    rango = Exl_rangos(fila, fila, 1, sprReal(0).MaxCols)
    Exl_Letra Ap, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion Ap, rango, Exl_Izq, Exl_CentroVert, False

    'AppExcel.Application.Range(rango).Merge
    'Exl_Lineas AppExcel, rango, Exl_Linsimple
    '---------------------------------------
    fila = fila + 2
    LinAbajo = 0
    LinArriba = 0
    'Exl_BajarGrillaExel spr, Ap, fila, Col, titCol
    For f = 0 To spr.MaxRows
        Varianza = False
        For c = 1 To spr.MaxCols
            If f = 0 Then
              Ap.Application.Cells(fila + f, c).Value = titCol(c)
              tipo = "CABECERA"
              
            Else
              spr.GetText c, f, valor
              If InStr(1, "COSPERCOMBEBACCREGTABCIGELE", valor) > 0 Then
                 Exl_PonerValor Ap, fila + f, c, L_NombreRubro(Trim(valor))
                 If valor <> "" Then
                   rango = Exl_rangos(fila + f, fila + f, 1, spr.MaxCols)
                   Exl_ColorInt Ap, rango, Exl_Gris
                 End If
                 tipo = "*"
                 Exit For
              Else
                 Exl_PonerValor Ap, fila + f, c, valor
                 tipo = "LINEA"
                 If Val(valor) < -21 And c = 5 Then
                    Varianza = True
                 End If
              End If
            End If
        Next
        Select Case tipo
            Case "CABECERA"
              rango = Exl_rangos(fila + f, fila + f, 1, spr.MaxCols)
              Exl_ColorInt Ap, rango, Exl_Gris
              Exl_Letra Ap, rango, NEGRITA, 8, "ARIAL"
              Exl_Lineas Ap, rango, Exl_Linsimple
              Exl_Justificacion Ap, rango, Exl_CentroCol, Exl_CentroVert, True
            Case "LINEA"
                rango = Exl_rangos(fila + f, fila + f, 1, spr.MaxCols)
                Exl_Lineas Ap, rango, Exl_Linsimple
                If Varianza Then
                    Ap.Application.Range(rango).Font.Bold = True
                    LinAbajo = LinAbajo + 1
                Else
                    LinArriba = LinArriba + 1
                End If
            

        End Select
    Next
    
    rango = Exl_rangos(fila + 1, fila + spr.MaxRows, Col + 1, spr.MaxCols)
    Ap.Application.Range(rango).NumberFormat = "#,##0_ ;[Red]-#,##0 "
    
    Exl_Letra Ap, rango, 2, 8, "Ms Serif"

    Exl_PonerValor Ap, fila + 2 + f, 1, "Total Líneas en Monitoreo:"
    Exl_PonerValor Ap, fila + 2 + f, 2, str(LinAbajo + LinArriba)
    Exl_PonerValor Ap, fila + 3 + f, 1, "Líneas arriba            :"
    Exl_PonerValor Ap, fila + 3 + f, 2, str(LinArriba)
    Exl_PonerValor Ap, fila + 4 + f, 1, "Líneas abajo             :"
    Exl_PonerValor Ap, fila + 4 + f, 2, str(LinAbajo)
    Exl_PonerValor Ap, fila + 5 + f, 1, "Porcentaje de cumplimiento:"
    Exl_PonerValor Ap, fila + 5 + f, 2, Format(100 * LinArriba / (LinAbajo + LinArriba), "#####.00")
    Exl_PonerValor Ap, fila + 7 + f, 1, "Cantidad de líneas a cumplir:"
    Exl_PonerValor Ap, fila + 7 + f, 2, Format((LinAbajo + LinArriba) * 0.8, "#####.00")

    Exl.Exl_AnchoCol Ap, 1, 1, 20
    'Exl.Exl_AnchoCol AppExcel, 1, 1, 19
    'Exl.Exl_AnchoCol AppExcel, 5, 5, 8
    'Exl.Exl_AnchoCol AppExcel, 8, 8, 8

    Ap.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    Ap.Application.ActiveSheet.PageSetup.Orientation = 2
    Ap.Application.ActiveSheet.PageSetup.PrintTitleRows = "$1:$5"
'    AppExcel.Application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
'    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
'
'End If

ErrorExl:

    frmMonitoreoLinea.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub




Private Sub L_PlanillaExcel(titulo As String, tipo As String)
Dim AppExcel As Excel.Application
Dim libroExcel As Excel.Workbook
Dim hojaExcel As Excel.Worksheet
Dim i As Integer
Dim NOMBRE As String
Dim tit(1 To 2) As String

'On Error GoTo ErrorExl:

tit(1) = " x LINEA "
tit(2) = " LANZ. "

NOMBRE = frmDir.NombreArchivo()
DoEvents

If NOMBRE <> "" Then
    frmMonitoreoLinea.caption = Aplicacion.SeteoProceso(frmMonitoreoLinea.caption)
    
    Set AppExcel = CreateObject("Excel.application")
    
    Set libroExcel = AppExcel.Workbooks.Add
    
    AppExcel.Visible = True
        
            Set hojaExcel = libroExcel.Worksheets(1)
            hojaExcel.Activate
            hojaExcel.Name = tit(1)
            L_TratarExcel hojaExcel, sprReal(0), "Listado de líneas Monitoreadas correspondientes al día " & Format(Now, "dd/mm/yyyy")
            Set hojaExcel = libroExcel.Worksheets(2)
            hojaExcel.Activate
            hojaExcel.Name = tit(2)
            L_TratarExcel hojaExcel, sprReal(1), "Listado de líneas con códigos en Lanzamiento"
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
      libroExcel.PrintOut
    End If
    
    libroExcel.SaveAs (NOMBRE & ".xls")
    AppExcel.Quit
    Set AppExcel = Nothing
    
    frmMonitoreoLinea.caption = Aplicacion.SeteoFin
End If
Exit Sub

ErrorExl:
    frmResumenDiario.caption = Aplicacion.SeteoFin
    MsgBox "Planilla no generada", vbOKOnly + vbCritical, "ATENCION"
    Exit Sub
End Sub


Private Sub cboProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    cboProv.ListIndex = -1
End If
End Sub

Private Sub cboRubro_Click()
Dim sql As String

'sql = " select distinct pv.cod_prov , pv.descrip "
'sql = sql & " from  baires.proveedor pv "
'sql = sql & " where cod_prov in (select cod_prov from estadis.monitoreo_linea ) "
'sql = sql & " and cod_rubr = '" & CboRubro.Text & "' "

'FuncCbos_LlenarCboLst cboProv, lstProv, sql

End Sub


Private Sub cboRubro_KeyPress(KeyAscii As Integer)
Dim sql As String

If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    cboRubro.ListIndex = -1


End If

End Sub


Private Sub Form_Activate()
tabEst.TabVisible(1) = False
tabEst.TabVisible(2) = False
tabEst.TabVisible(3) = False
End Sub

Private Sub Form_Load()
Dim sql As String

Top = 30
Left = 250
Height = 7500
Width = 11000

sql = " SELECT cod_rubr,descrip FROM baires.rubro "
sql = sql & " ORDER BY cod_rubr"

FuncCbos_LlenarCbo cboRubro, sql

sql = " select distinct pv.cod_prov , pv.descrip "
sql = sql & " from  baires.proveedor pv "
sql = sql & " where cod_prov in (select cod_prov from estadis.monitoreo_linea ) "

FuncCbos_LlenarCboLst cboProv, lstProv, sql

mskAnio.Text = Year(Date)
mskMes.Text = Month(Date)


L_LimpiarGrillas
'frmPrincipal.lstForms.AddItem "frmVsDia"

mskCota.Text = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)
    FuncLocal_SacarForm "frmVsDia"
End Sub


Private Sub mskAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub mskAnio_LostFocus()
If Val(mskAnio.Text) < 1996 Or Val(mskAnio) > 2050 Then
    mskAnio.Text = Year(Date)
End If

End Sub


Private Sub mskMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub mskMes_LostFocus()
If Val(mskMes.Text) < 1 Or Val(mskMes.Text) > 12 Then
    mskMes.Text = Month(Date)
End If

End Sub


Private Sub tabEst_Click(PreviousTab As Integer)
On Error GoTo ErrT:

'Call sprReal_LeaveCell(tabEst.Tab, 1, 1, 1, sprReal(tabEst.Tab).ActiveRow, True)
'Call sprProy_LeaveCell(tabEst.Tab, 1, 1, 1, sprProy(tabEst.Tab).ActiveRow, True)

ErrT:
    Exit Sub

End Sub

Private Sub tabGA_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    'sprGA(tabGA.Tab).SetFocus
    
ErrT:
    Exit Sub

End Sub

Private Sub tabTotal_Click(PreviousTab As Integer)

On Error GoTo ErrT:

    'sprTotal(tabTotal.Tab).SetFocus
    
ErrT:
    Exit Sub
End Sub

