VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMonitoreo 
   Caption         =   "Estimados por productos vs Ventas"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   10875
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   5850
      TabIndex        =   31
      Top             =   585
      Width           =   3270
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
         Left            =   1635
         TabIndex        =   33
         Top             =   210
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
         TabIndex        =   32
         Top             =   210
         Value           =   -1  'True
         Width           =   1290
      End
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
      Left            =   9345
      Picture         =   "frmMonitoreo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   720
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
      Picture         =   "frmMonitoreo.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   12
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
      Left            =   10035
      Picture         =   "frmMonitoreo.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   165
      Width           =   615
   End
   Begin VB.CommandButton botExcel 
      Caption         =   "Excel"
      Height          =   525
      Left            =   10035
      Picture         =   "frmMonitoreo.frx":0A26
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
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
      Height          =   1380
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   10755
      Begin VB.ListBox lstFamilia 
         Height          =   255
         Left            =   9090
         TabIndex        =   23
         Top             =   330
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox cboFamilia 
         Height          =   315
         Left            =   6810
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   315
         Width           =   2265
      End
      Begin VB.ListBox LstEspigon 
         Height          =   255
         Left            =   5760
         TabIndex        =   3
         Top             =   765
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox CboRubro 
         Height          =   315
         ItemData        =   "frmMonitoreo.frx":0FB8
         Left            =   3330
         List            =   "frmMonitoreo.frx":0FBA
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2415
      End
      Begin VB.ComboBox CboSubRubro 
         Height          =   315
         Left            =   3315
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   750
         Width           =   2415
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   315
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
         TabIndex        =   5
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
         Left            =   3960
         TabIndex        =   30
         Top             =   1035
         Visible         =   0   'False
         Width           =   1110
         _ExtentX        =   1958
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
         Caption         =   "FAMILIA"
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
         Left            =   5820
         TabIndex        =   21
         Top             =   315
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         Left            =   2085
         TabIndex        =   7
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label LblEspigon 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SUB RUBRO"
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
         Left            =   2085
         TabIndex        =   6
         Top             =   750
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab tabEst 
      Height          =   4800
      Left            =   15
      TabIndex        =   14
      Top             =   1470
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   8467
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   6
      TabHeight       =   441
      ForeColor       =   255
      TabCaption(0)   =   "TOTAL CIA"
      TabPicture(0)   =   "frmMonitoreo.frx":0FBC
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tabTotal"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "EZEIZA"
      TabPicture(1)   =   "frmMonitoreo.frx":0FD8
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabGA"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "AEROPARQUE"
      TabPicture(2)   =   "frmMonitoreo.frx":0FF4
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab1"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "INTERIOR"
      TabPicture(3)   =   "frmMonitoreo.frx":1010
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTab2"
      Tab(3).Control(0).Enabled=   0   'False
      Begin TabDlg.SSTab tabTotal 
         Height          =   4245
         Left            =   165
         TabIndex        =   15
         Top             =   360
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   7488
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
         TabPicture(0)   =   "frmMonitoreo.frx":102C
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprReal(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "PROYECTADO"
         TabPicture(1)   =   "frmMonitoreo.frx":1048
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprProy(0)"
         Tab(1).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprProy 
            Height          =   3660
            Index           =   0
            Left            =   -74895
            OleObjectBlob   =   "frmMonitoreo.frx":1064
            TabIndex        =   26
            Top             =   120
            Width           =   10200
         End
         Begin FPSpread.vaSpread sprReal 
            Height          =   3660
            Index           =   0
            Left            =   105
            OleObjectBlob   =   "frmMonitoreo.frx":178A
            TabIndex        =   16
            Top             =   120
            Width           =   10200
         End
      End
      Begin TabDlg.SSTab tabGA 
         Height          =   4290
         Left            =   -74820
         TabIndex        =   17
         Top             =   345
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   7567
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
         TabPicture(0)   =   "frmMonitoreo.frx":1EB0
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprReal(1)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "PROYECTADO"
         TabPicture(1)   =   "frmMonitoreo.frx":1ECC
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprProy(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprProy 
            Height          =   3660
            Index           =   1
            Left            =   -74895
            OleObjectBlob   =   "frmMonitoreo.frx":1EE8
            TabIndex        =   27
            Top             =   120
            Width           =   10200
         End
         Begin FPSpread.vaSpread sprReal 
            Height          =   1095
            Index           =   1
            Left            =   105
            OleObjectBlob   =   "frmMonitoreo.frx":25FC
            TabIndex        =   20
            Top             =   150
            Width           =   4095
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4290
         Left            =   -74820
         TabIndex        =   18
         Top             =   360
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   7567
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
         TabPicture(0)   =   "frmMonitoreo.frx":2D10
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprReal(2)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "PROYECTADO"
         TabPicture(1)   =   "frmMonitoreo.frx":2D2C
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprProy(2)"
         Tab(1).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprProy 
            Height          =   3660
            Index           =   2
            Left            =   -74895
            OleObjectBlob   =   "frmMonitoreo.frx":2D48
            TabIndex        =   28
            Top             =   120
            Width           =   10200
         End
         Begin FPSpread.vaSpread sprReal 
            Height          =   1140
            Index           =   2
            Left            =   90
            OleObjectBlob   =   "frmMonitoreo.frx":345C
            TabIndex        =   24
            Top             =   120
            Width           =   5085
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4290
         Left            =   -74850
         TabIndex        =   19
         Top             =   345
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   7567
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
         TabPicture(0)   =   "frmMonitoreo.frx":3B70
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprReal(3)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "PROYECTADO"
         TabPicture(1)   =   "frmMonitoreo.frx":3B8C
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprProy(3)"
         Tab(1).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprProy 
            Height          =   3660
            Index           =   3
            Left            =   -74895
            OleObjectBlob   =   "frmMonitoreo.frx":3BA8
            TabIndex        =   29
            Top             =   120
            Width           =   10200
         End
         Begin FPSpread.vaSpread sprReal 
            Height          =   3660
            Index           =   3
            Left            =   90
            OleObjectBlob   =   "frmMonitoreo.frx":42BC
            TabIndex        =   25
            Top             =   120
            Width           =   10200
         End
      End
   End
End
Attribute VB_Name = "frmMonitoreo"
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

For i = 0 To 3
 sprReal(i).MaxRows = 0
 sprProy(i).MaxRows = 0
Next
 
End Sub




Private Function L_NombreDato(i As Integer) As String
Select Case i
    Case 1
        L_NombreDato = "Importes"
    Case 2
        L_NombreDato = "Tickets"
    Case 3
        L_NombreDato = "Pasajeros Transitados"
    Case 4
        L_NombreDato = "Promedios por Ticket"
    Case 5
        L_NombreDato = "Promedios por Pasajeros Transitados"
End Select
End Function


Private Sub L_TratarPromedios(sprS As control, sprD As control, sprR As control)
Dim fila As Integer
Dim fch As Variant
Dim i

If sprS.MaxRows = sprD.MaxRows Then
fila = 1
Do While fila <= sprS.MaxRows
    sprR.MaxRows = sprR.MaxRows + 1
    sprS.GetText 1, fila, fch
    sprR.SetText 1, sprR.MaxRows, Format$(fch, "dd-mm-yy")
    fila = fila + 1
Loop
Spread.Func_PromediosCol sprS, sprD, sprR, 3, 1, sprS.MaxRows
Spread.Func_PromediosCol sprS, sprD, sprR, 4, 1, sprS.MaxRows

Spread.Spread_TotalesLinea sprR

For fila = 1 To sprR.MaxRows
    spread_ResaltarCelda sprR, 5, fila
    spread_ResaltarCelda sprR, 6, fila
    'spread_ResaltarCelda sprR, 9, fila
    'spread_ResaltarCelda sprR, 10, fila
Next
End If

End Sub

Private Sub botEjecutar_Click(Index As Integer)

Select Case Index
    Case 0
        L_Refrescar
      
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


'On Error GoTo ErrGLC:

frmMonitoreo.caption = Aplicacion.SeteoProceso(frmMonitoreo.caption)

sql = "  SELECT pv.cod_prov,pv.descrip "
sql = sql & " ,sum(venta_real_u) venta_real_u,sum(venta_real_$) venta_real_p"
sql = sql & " ,sum(estimado_u)estimado_u,sum(estimado_$)estimado_p"
sql = sql & " ,round(sum(estimado_u/to_number(to_char(last_day(sysdate-1),'DD'))*to_number(to_char(sysdate-1,'DD')))) estim_dia_u"
sql = sql & " ,round(sum(estimado_$/to_number(to_char(last_day(sysdate-1),'DD'))*to_number(to_char(sysdate-1,'DD'))),2) estim_dia_p"
sql = sql & " ,round(sum(venta_real_u/to_number(to_char(sysdate-1,'DD'))*to_number(to_char(last_day(sysdate-1),'DD')))) proyectado_u"
sql = sql & " ,round(sum(venta_real_$/to_number(to_char(sysdate-1,'DD'))*to_number(to_char(last_day(sysdate-1),'DD'))),2) proyectado_p"
sql = sql & " From baires.proveedor pv,baires.estm_venta e"
sql = sql & " Where aniomes = " & mskAnio.Text & Format(mskMes.Text, "00")
sql = sql & " and cod_rubr <> 'REG' and e.cod_prov = pv.cod_prov "
sql = sql & " " & L_Armarcondicion
sql = sql & " group by pv.cod_prov,descrip "


If Aplicacion.ObtenerRsDAO(sql, rs) Then
      L_LlenarGrillasReal
      tabEst.Enabled = True
      botEjecutar(0).Enabled = False
      frCab.Enabled = False
End If


ErrGLC:
    frmMonitoreo.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub

Private Function L_Armarcondicion()
Dim Cond

If CboRubro.Text <> "" Then
   Cond = Cond & " And cod_rubr = '" & CboRubro.Text & "'"
End If

If CboSubRubro.Text <> "" Then
   Cond = Cond & " And cod_srub = '" & LstEspigon.List(CboSubRubro.ListIndex) & "'"
End If

If cboFamilia.Text <> "" Then
   Cond = Cond & " And cod_familia = '" & lstFamilia.List(cboFamilia.ListIndex) & "'"
End If

L_Armarcondicion = Cond

End Function



Private Sub L_LlenarGrillasReal()

Do While Not rs.EOF
        sprReal(0).MaxRows = sprReal(0).MaxRows + 1
        sprReal(0).SetText 1, sprReal(0).MaxRows, Trim(rs!Descrip)
        sprReal(0).SetText 8, sprReal(0).MaxRows, Trim(rs!cod_prov)
        sprReal(0).SetText 2, sprReal(0).MaxRows, str(rs!estim_dia_u)
        sprReal(0).SetText 3, sprReal(0).MaxRows, str(rs!venta_real_U)
        sprReal(0).SetText 5, sprReal(0).MaxRows, Trim(rs!estim_dia_p)
        sprReal(0).SetText 6, sprReal(0).MaxRows, str(rs!venta_real_P)
                
        sprProy(0).MaxRows = sprProy(0).MaxRows + 1
        sprProy(0).SetText 1, sprProy(0).MaxRows, Trim(rs!Descrip)
        sprProy(0).SetText 8, sprProy(0).MaxRows, Trim(rs!cod_prov)
        sprProy(0).SetText 2, sprProy(0).MaxRows, str(rs!estimado_U)
        sprProy(0).SetText 3, sprProy(0).MaxRows, str(rs!proyectado_u)
        sprProy(0).SetText 5, sprProy(0).MaxRows, str(rs!estimado_P)
        sprProy(0).SetText 6, sprProy(0).MaxRows, str(rs!proyectado_p)
                
                
        rs.MoveNext
        If rs.EOF Then
            Exit Do
        End If
    Loop

End Sub

Private Sub botExcel_Click()


Select Case tabTotal.Tab
    Case 0
        L_TratarExcelTot sprReal(0), "Período " & mskAnio.Text & " / " & mskMes.Text & " caculado a la fecha corriente"
    Case 1
        L_TratarExcelTot sprProy(0), "Período " & mskAnio.Text & " / " & mskMes.Text & "  proyectado total mes "
    Case 2
    
    Case 3
    
End Select

End Sub



Private Sub L_TratarExcelTot(spr As control, titulo As String)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim i As Integer
Dim tit As Variant

Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmMonitoreo.caption = Aplicacion.SeteoProceso(frmMonitoreo.caption)

If NOMBRE <> "" Then '

    Set AppExcel = CreateObject("excel.sheet")

    'AppExcel.Application.Visible = True

    ReDim titCol(sprReal(0).MaxCols)
    Col = 1
    fila = 3

    For i = 1 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next

    Exl_PonerValor AppExcel, 1, 1, "INFORME COMPARATIVO VENTA-ESTIMADO"
    rango = Exl_rangos(1, 1, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Blanco

    Exl_PonerValor AppExcel, fila, Col, titulo
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False

    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    '---------------------------------------
    
    'fila = fila + 2

    'Exl_PonerValor AppExcel, fila, Col, ""
    'rango = Exl_rangos(fila, fila, 1, sprReal(0).MaxCols)
    'Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    'Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    'Exl_ColorInt AppExcel, rango, Exl_Blanco
    'AppExcel.Application.Range(rango).Merge

    fila = fila + 1

    Exl_BajarGrillaExel spr, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + spr.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows
'    rango = Exl_rangos(fila, fila, Col, sprReal(0).MaxCols)
'    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 1
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 8, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris

    'Exl.Exl_AnchoCol AppExcel, sprReal(0).MaxCols, sprReal(0).MaxCols, 1
    Exl.Exl_AnchoCol AppExcel, 1, 1, 20
    Exl.Exl_AnchoCol AppExcel, 4, 4, 7
    Exl.Exl_AnchoCol AppExcel, 7, 7, 7

    'Elimina las columnas que sobran
    rango = Exl_rangos(1, 1000, 8, 9)
    AppExcel.Application.Range(rango).Delete

    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    'AppExcel.Application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If

    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmMonitoreo.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub




Private Sub cboFamilia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    cboFamilia.ListIndex = -1
End If


End Sub

Private Sub cboRubro_Click()
Dim sql As String

sql = " SELECT cod_srub,descr FROM baires.subrubro "
sql = sql & " WHERE cod_rubr = '" & CboRubro.Text & "'"
sql = sql & " ORDER BY descr"
 
FuncCbos_LlenarCboLst CboSubRubro, LstEspigon, sql
    
cboFamilia.Clear

End Sub


Private Sub cboRubro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    CboRubro.ListIndex = -1
    cboFamilia.ListIndex = -1
End If

End Sub


Private Sub CboSubRubro_Click()
Dim sql As String

sql = " SELECT cod_FAMILIA,descrip FROM baires.familia "
sql = sql & " WHERE cod_rubr = '" & CboRubro.Text & "'"
sql = sql & " And cod_srub = '" & LstEspigon.List(CboSubRubro.ListIndex) & "'"
sql = sql & " ORDER BY descrip "
 
FuncCbos_LlenarCboLst cboFamilia, lstFamilia, sql

End Sub

Private Sub cboSubRubro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    CboSubRubro.ListIndex = -1
    cboFamilia.ListIndex = -1
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

FuncCbos_LlenarCbo CboRubro, sql

mskAnio.Text = Year(Date)
mskMes.Text = Month(Date)


L_LimpiarGrillas
'frmPrincipal.lstForms.AddItem "frmVsDia"

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


Private Sub optSort_Click(Index As Integer)

Select Case Index
    Case 0
        sprReal(0).Row = 1
        sprReal(0).Col = 1
        sprReal(0).Row2 = sprReal(0).MaxRows - 1
        sprReal(0).Col2 = 10
        
        ' Set sort definition for key 1
        sprReal(0).SortBy = SS_SORT_BY_ROW

        sprReal(0).SortKey(1) = 1
        sprReal(0).SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprReal(0).Action = SS_ACTION_SORT
    
        sprProy(0).Row = 1
        sprProy(0).Col = 1
        sprProy(0).Row2 = sprProy(0).MaxRows - 1
        sprProy(0).Col2 = 10
        
        ' Set sort definition for key 1
        sprProy(0).SortBy = SS_SORT_BY_ROW

        sprProy(0).SortKey(1) = 1
        sprProy(0).SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        sprProy(0).Action = SS_ACTION_SORT
    
    Case 1
        sprReal(0).Row = 1
        sprReal(0).Col = 1
        sprReal(0).Row2 = sprReal(0).MaxRows - 1
        sprReal(0).Col2 = 10
        
        ' Set sort definition for key 1
        sprReal(0).SortBy = SS_SORT_BY_ROW
        
        sprReal(0).SortKey(1) = 6
        sprReal(0).SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprReal(0).Action = SS_ACTION_SORT

        sprProy(0).Row = 1
        sprProy(0).Col = 1
        sprProy(0).Row2 = sprProy(0).MaxRows - 1
        sprProy(0).Col2 = 10
        
        ' Set sort definition for key 1
        sprProy(0).SortBy = SS_SORT_BY_ROW
        
        sprProy(0).SortKey(1) = 6
        sprProy(0).SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
        sprProy(0).Action = SS_ACTION_SORT

End Select


End Sub

Private Sub sprProy_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim datos As String, Esp As String
Dim sql As String
Dim prov As Variant

datos = "CIA."
 
sprProy(Index).GetText 8, Row, prov

If Col = 8 Then
    
    sql = "  SELECT cod_prod || '-'||descrip_prod cod_prod "
    sql = sql & " ,sum(estimado_u) e_u,sum(estimado_$) e_p "
    sql = sql & " ,round(sum(venta_real_u/to_number(to_char(sysdate-1,'DD'))*to_number(to_char(last_day(sysdate-1),'DD'))))   v_u "
    sql = sql & " ,round(sum(venta_real_$/to_number(to_char(sysdate-1,'DD'))*to_number(to_char(last_day(sysdate-1),'DD'))),2) v_p "
    sql = sql & " ,baires.stock_cia(cod_prod) stk"
    sql = sql & " From baires.estm_venta e "
    sql = sql & " Where aniomes = " & mskAnio.Text & Format(mskMes.Text, "00")
    sql = sql & " and cod_rubr <> 'REG' And cod_prov = '" & prov & "' "
    sql = sql & " " & L_Armarcondicion
    sql = sql & " group by cod_prod || '-'||descrip_prod , baires.stock_cia(cod_prod) "
        
    frmMonitoreoCod.Mostrar sql, "Proyección - Detalle por producto :" & datos, Trim(prov)

End If

End Sub

Private Sub sprReal_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim datos As String, Esp As String
Dim sql As String
Dim prov As Variant

 
sprReal(Index).GetText 8, Row, prov

If Col = 8 Then
    sql = "  SELECT cod_prod || '-'||descrip_prod cod_prod "
    sql = sql & " ,sum(venta_real_u) v_u,sum(venta_real_$) v_p"
    sql = sql & " ,round(sum(estimado_u/to_number(to_char(last_day(sysdate-1),'DD'))*to_number(to_char(sysdate-1,'DD')))) e_u"
    sql = sql & " ,round(sum(estimado_$/to_number(to_char(last_day(sysdate-1),'DD'))*to_number(to_char(sysdate-1,'DD'))),2) e_p"
    sql = sql & " ,baires.stock_cia(cod_prod) stk"
    sql = sql & " From baires.estm_venta e"
    sql = sql & " Where aniomes = " & mskAnio.Text & Format(mskMes.Text, "00")
    sql = sql & " and cod_rubr <> 'REG' And cod_prov = '" & prov & "' "
    sql = sql & " " & L_Armarcondicion
    sql = sql & " group by cod_prod || '-'||descrip_prod , baires.stock_cia(cod_prod)"

    frmMonitoreoCod.Mostrar sql, "A la fecha - Detalle por producto : CIA ", Trim(prov)
End If



End Sub

Private Sub sprReal_DblClick(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim sql As String
Dim prov As Variant, descPv As Variant

sprReal(Index).GetText 8, Row, prov
sprReal(Index).GetText 1, Row, descPv

If mskMes.Text > 1 Then
    sql = "  SELECT ANIOMES"
    sql = sql & " ,sum(venta_real_U)venta_real_U"
    sql = sql & " ,sum(estimado_U)estimado_U"
    sql = sql & " ,sum(venta_real_$)venta_real_P"
    sql = sql & " ,sum(estimado_$)estimado_P"
    sql = sql & " From baires.estm_venta"
    sql = sql & " Where"
    sql = sql & " aniomes BETWEEN " & mskAnio.Text & "01 AND " & mskAnio.Text & Format(mskMes.Text - 1, "00")
    sql = sql & " and cod_rubr <> 'REG'"
    sql = sql & " and cod_prov = '" & prov & "' "
    sql = sql & " group by ANIOMES "
    
    frmMonitoreoMes.Mostrar sql, Trim(prov), Trim(descPv), CboRubro.Text, "Proveedor"

End If
End Sub


Private Sub tabEst_Click(PreviousTab As Integer)
On Error GoTo ErrT:

'    Select Case tabEst.Tab
'        Case 0
'            sprTotal(tabTotal.Tab).SetFocus
'        Case 1
'            sprGA(tabGA.Tab).SetFocus
 '       Case 2
'            sprGB(tabGB.Tab).SetFocus
'        Case 3
'            sprGC(tabGC.Tab).SetFocus
'    End Select
    
    
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

