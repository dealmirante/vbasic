VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVentaMetros 
   Caption         =   "Ventas por Metro cuadrado"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10875
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   6165
      TabIndex        =   7
      Top             =   585
      Width           =   3045
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
         Left            =   1470
         TabIndex        =   9
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
         TabIndex        =   8
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
      Picture         =   "frmVentaMetros.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Picture         =   "frmVentaMetros.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   5
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
      Picture         =   "frmVentaMetros.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   150
      Width           =   615
   End
   Begin VB.CommandButton botExcel 
      Caption         =   "Excel"
      Height          =   525
      Left            =   10035
      Picture         =   "frmVentaMetros.frx":0A26
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Frame frCab 
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
      TabIndex        =   1
      Top             =   -30
      Width           =   10755
      Begin VB.Frame Frame4 
         Caption         =   "Período Actual"
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
         Height          =   1095
         Left            =   180
         TabIndex        =   15
         Top             =   135
         Width           =   2790
         Begin MSMask.MaskEdBox mskFDesde 
            Height          =   285
            Left            =   1485
            TabIndex        =   16
            Top             =   255
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
            Left            =   1485
            TabIndex        =   17
            Top             =   600
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
            Left            =   240
            TabIndex        =   19
            Top             =   255
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
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   1185
         End
      End
      Begin VB.Frame frPerAnt 
         Caption         =   "Período Anterior"
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
         Height          =   1095
         Left            =   3120
         TabIndex        =   10
         Top             =   135
         Width           =   2895
         Begin MSMask.MaskEdBox mskFDesdeAnt 
            Height          =   285
            Left            =   1560
            TabIndex        =   11
            Top             =   255
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "##-##-####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFHastaAnt 
            Height          =   285
            Left            =   1560
            TabIndex        =   12
            Top             =   600
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
            Index           =   3
            Left            =   315
            TabIndex        =   14
            Top             =   600
            Width           =   1185
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
            Index           =   2
            Left            =   330
            TabIndex        =   13
            Top             =   255
            Width           =   1185
         End
      End
      Begin VB.ComboBox CboLocal 
         Height          =   315
         ItemData        =   "frmVentaMetros.frx":0FB8
         Left            =   7635
         List            =   "frmVentaMetros.frx":0FBA
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   195
         Width           =   1515
      End
      Begin VB.Label LblLocal 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
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
         Left            =   6150
         TabIndex        =   0
         Top             =   210
         Width           =   1290
      End
   End
   Begin TabDlg.SSTab tabTotal 
      Height          =   5310
      Left            =   30
      TabIndex        =   20
      Top             =   1395
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   9366
      _Version        =   327680
      TabOrientation  =   1
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "VENTA POR SECTOR"
      TabPicture(0)   =   "frmVentaMetros.frx":0FBC
      Tab(0).ControlCount=   4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "sprReal(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(0)"
      Tab(0).Control(3).Enabled=   0   'False
      TabCaption(1)   =   "PRODUCTIVIDAD POR SECTOR"
      TabPicture(1)   =   "frmVentaMetros.frx":0FD8
      Tab(1).ControlCount=   4
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2(4)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2(5)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "sprReal(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Begin FPSpread.vaSpread sprReal 
         Height          =   4470
         Index           =   1
         Left            =   225
         OleObjectBlob   =   "frmVentaMetros.frx":0FF4
         TabIndex        =   25
         Top             =   375
         Width           =   10230
      End
      Begin FPSpread.vaSpread sprReal 
         Height          =   4455
         Index           =   0
         Left            =   -74790
         OleObjectBlob   =   "frmVentaMetros.frx":172A
         TabIndex        =   21
         Top             =   390
         Width           =   10230
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   5
         Left            =   2340
         TabIndex        =   28
         Top             =   0
         Width           =   2205
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   4
         Left            =   4515
         TabIndex        =   27
         Top             =   0
         Width           =   2130
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diferencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   3
         Left            =   6630
         TabIndex        =   26
         Top             =   0
         Width           =   2985
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diferencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   2
         Left            =   -68310
         TabIndex        =   24
         Top             =   0
         Width           =   2985
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   1
         Left            =   -70470
         TabIndex        =   23
         Top             =   0
         Width           =   2190
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   0
         Left            =   -72675
         TabIndex        =   22
         Top             =   0
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frmVentaMetros"
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
 'sprProy(i).MaxRows = 0
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
        L_LimpiarGrillas
    Case 2
        Unload Me
End Select


End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlP As String


On Error GoTo ErrGLC:

frmVentaMetros.caption = Aplicacion.SeteoProceso(frmVentaMetros.caption)

sql = " select cod_sector,descrip, sum(actMtr) actMtr ,sum(actImp) actImp, sum(antMtr) antMtr ,sum(antImp) antImp "
sql = sql & " From (SELECT s.COD_SECTOR,descrip, round(avg(v1.METROS_TOTAL_SECTOR),2) actMtr,sum(v1.IMPORTE_PROV) actImp, 0 antMtr, 0 antImp "
sql = sql & " FROM ESTADIS.Z_VENTA_SECTOR v1 , estadis.z_sector s "
sql = sql & " where v1.fch_ticket between " & Func.func_ToDate(mskFDesde.FormattedText) & " And " & Func.func_ToDate(mskFHasta.FormattedText)
sql = sql & " and v1.cod_sector = s.cod_sector "
sql = sql & " " & L_Armarcondicion
sql = sql & "  group by s.COD_SECTOR,descrip "
sql = sql & " Union "
sql = sql & " SELECT s.COD_SECTOR,descrip, 0 actMtr,0 actImp, round(avg(v2.METROS_TOTAL_SECTOR),2) antMtr, sum(v2.IMPORTE_PROV) antImp "
sql = sql & " FROM ESTADIS.Z_VENTA_SECTOR v2 , estadis.z_sector s "
sql = sql & " where v2.fch_ticket between " & Func.func_ToDate(mskFDesdeAnt.FormattedText) & " And " & Func.func_ToDate(mskFHastaAnt.FormattedText)
sql = sql & "  and v2.cod_sector = s.cod_sector "
sql = sql & " " & L_Armarcondicion
sql = sql & " group by s.COD_SECTOR,descrip) "
sql = sql & " group by cod_sector,descrip "
sql = sql & " order by descrip "

If Aplicacion.ObtenerRsDAO(sql, rs) Then
      L_LlenarGrillasReal

      botEjecutar(0).Enabled = False
      frCab.Enabled = False
End If


ErrGLC:
    frmVentaMetros.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub

Private Function L_Armarcondicion()
Dim Cond

If CboLocal.Text <> "" Then
   Cond = Cond & " And cod_local = '" & CboLocal.Text & "'"
End If

L_Armarcondicion = Cond

End Function



Private Sub L_LlenarGrillasReal()

Do While Not rs.EOF
        sprReal(0).MaxRows = sprReal(0).MaxRows + 1
        sprReal(0).SetText 1, sprReal(0).MaxRows, Trim(rs!Descrip)
        sprReal(0).SetText 10, sprReal(0).MaxRows, Trim(rs!cod_sector)
        sprReal(0).SetText 2, sprReal(0).MaxRows, str(rs!ActMtr)
        sprReal(0).SetText 3, sprReal(0).MaxRows, str(rs!actimp)
        sprReal(0).SetText 4, sprReal(0).MaxRows, Trim(rs!AntMtr)
        sprReal(0).SetText 5, sprReal(0).MaxRows, str(rs!antimp)
                
        sprReal(1).MaxRows = sprReal(1).MaxRows + 1
        sprReal(1).SetText 1, sprReal(1).MaxRows, Trim(rs!Descrip)
        sprReal(1).SetText 10, sprReal(1).MaxRows, Trim(rs!cod_sector)
        sprReal(1).SetText 2, sprReal(1).MaxRows, str(rs!ActMtr)
        If rs!ActMtr > 0 Then
            sprReal(1).SetText 3, sprReal(1).MaxRows, str(rs!actimp / rs!ActMtr)
        End If
        sprReal(1).SetText 4, sprReal(1).MaxRows, Trim(rs!AntMtr)
        If rs!AntMtr > 0 Then
        sprReal(1).SetText 5, sprReal(1).MaxRows, str(rs!antimp / rs!AntMtr)
        End If
                
                
        rs.MoveNext
        If rs.EOF Then
            Exit Do
        End If
    Loop

End Sub

Private Sub botExcel_Click()
Dim AppExcel As Excel.Application
Dim libroExcel As Excel.Workbook
Dim hojaExcel As Excel.Worksheet
Dim NOMBRE As String

NOMBRE = frmDir.NombreArchivo()
DoEvents

If NOMBRE <> "" Then

    Set AppExcel = CreateObject("Excel.application")
    'AppExcel.Visible = True
    Set libroExcel = AppExcel.Workbooks.Add

      Set hojaExcel = libroExcel.Worksheets(1)
      hojaExcel.Activate
      hojaExcel.Name = "Ventas Netas"
      L_TratarExcel hojaExcel, sprReal(0), "Ventas Netas"

      Set hojaExcel = libroExcel.Worksheets(2)
      hojaExcel.Activate
      hojaExcel.Name = "Productividad"
      L_TratarExcel hojaExcel, sprReal(1), "Productividad"

    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        libroExcel.PrintOut
    End If

    libroExcel.SaveAs NOMBRE & ".xls"
    Set libroExcel = Nothing
    AppExcel.Quit
    Set AppExcel = Nothing

End If
End Sub



Private Sub L_TratarExcel(AppExcel As Object, spr As control, titulo As String)
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer
Dim i As Integer
Dim tit As Variant


On Error GoTo ErrorExl:

DoEvents

frmVentaMetros.caption = Aplicacion.SeteoProceso(frmVentaMetros.caption)


    'Set AppExcel = CreateObject("excel.sheet")

    'AppExcel.Application.Visible = True

    ReDim titCol(spr.MaxCols)
    Col = 1
    fila = 3

    For i = 1 To spr.MaxCols - 2
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next

    Exl_PonerValor AppExcel, 1, 1, "Informe de productividad por Metro - Local " & CboLocal.Text & " - "
    rango = Exl_rangos(1, 1, 1, spr.MaxCols - 2)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Blanco

    Exl_PonerValor AppExcel, fila, Col, titulo
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False

    fila = fila + 1

    Exl_PonerValor AppExcel, fila, Col, "Período de ventas Actual : " & mskFDesde.FormattedText & " - " & mskFHasta.FormattedText
    
    fila = fila + 1

    Exl_PonerValor AppExcel, fila, Col, "Período de ventas Anterior : " & mskFDesdeAnt.FormattedText & " - " & mskFHastaAnt.FormattedText
    
    rango = Exl_rangos(fila - 2, fila, 1, spr.MaxCols - 2)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    'Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, Exl_Blanco
    'AppExcel.Application.Range(rango).Merge

    fila = fila + 2
    Exl_PonerValor AppExcel, fila, 2, "      ACTUAL"
    Exl_PonerValor AppExcel, fila, 4, "     ANTERIOR"
    Exl_PonerValor AppExcel, fila, 6, "     DIFERENCIAS"

    fila = fila + 1
    Exl_BajarGrillaExel spr, AppExcel, fila, Col, titCol
    rango = Exl_rangos(fila + 1, fila + spr.MaxRows, Col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows
'    rango = Exl_rangos(fila, fila, Col, sprReal(0).MaxCols)
'    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    fila = fila + 1
    
    'rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 8, spr.MaxCols)
    'Exl_ColorInt AppExcel, rango, Exl_Gris

    'Exl.Exl_AnchoCol AppExcel, sprReal(0).MaxCols, sprReal(0).MaxCols, 1
    Exl.Exl_AnchoCol AppExcel, 1, 1, 20
    Exl.Exl_AnchoCol AppExcel, 2, 2, 7
    Exl.Exl_AnchoCol AppExcel, 4, 4, 7
    Exl.Exl_AnchoCol AppExcel, 6, 6, 7
    Exl.Exl_AnchoCol AppExcel, 8, 8, 7
    Exl.Exl_AnchoCol AppExcel, 3, 3, 12
    Exl.Exl_AnchoCol AppExcel, 5, 5, 12

    'Elimina las columnas que sobran
    rango = Exl_rangos(1, 1000, 9, 10)
    AppExcel.Application.Range(rango).Delete

    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.Zoom = False
    
    'AppExcel.Application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    'If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
    '    AppExcel.PrintOut
    'End If

    'AppExcel.SaveAs NOMBRE & ".xls"
    'Set AppExcel = Nothing

ErrorExl:

    frmVentaMetros.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub




Private Sub Form_Load()
Dim sql As String

Top = 30
Left = 250
Height = 7500
Width = 11000

'sql = " SELECT cod_rubr,descrip FROM baires.rubro "
'sql = sql & " ORDER BY cod_rubr"
'
'FuncCbos_LlenarCbo CboRubro, sql

mskFDesde.Text = Format$(Func.func_Dia1SegunMes_Anio(Month(Date - 1), Year(Date - 1)), FTOFECHA)
mskFHasta.Text = Format$(Now - 1, FTOFECHA)
    
mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio - 1, FTOFECHA)
mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio - 1, FTOFECHA)

CboLocal.AddItem "L22"
CboLocal.AddItem "L03"
CboLocal.ListIndex = 0

L_LimpiarGrillas
'frmPrincipal.lstForms.AddItem "frmVsDia"

End Sub


Private Sub Form_Unload(Cancel As Integer)
    FuncLocal_SacarForm "frmVsDia"
End Sub


Private Sub mskFDesde_LostFocus()
    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    End If
    
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
    If mskFHasta.Text <> "" Then
        If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
            mskFHasta.Text = mskFDesde.Text
        End If
    End If

End Sub


Private Sub mskFDesdeAnt_LostFocus()
    
    If Not IsDate(mskFDesdeAnt.FormattedText) Then
        mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
    End If
    
    mskFDesdeAnt.Text = Format$(mskFDesdeAnt.FormattedText, FTOFECHA)
    If mskFHastaAnt.Text <> "" Then
        If CDate(mskFHastaAnt.FormattedText) < CDate(mskFDesdeAnt.FormattedText) Or Year(mskFHastaAnt.FormattedText) <> Year(mskFDesdeAnt.FormattedText) Then
            mskFHastaAnt.Text = mskFDesdeAnt.Text
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
mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio, FTOFECHA)

End Sub


Private Sub mskFHastaAnt_LostFocus()

If Not IsDate(mskFHastaAnt.FormattedText) Then
        mskFHastaAnt.Text = mskFDesdeAnt.Text
ElseIf CDate(mskFHastaAnt.FormattedText) < CDate(mskFDesdeAnt.FormattedText) _
    Or Year(mskFHastaAnt.FormattedText) <> Year(mskFDesdeAnt.FormattedText) Then
        mskFHastaAnt.Text = mskFDesdeAnt.Text
End If

mskFHastaAnt.Text = Format$(mskFHastaAnt.FormattedText, FTOFECHA)

End Sub


Private Sub optSort_Click(Index As Integer)
Dim i As Integer

Select Case Index
    Case 0
        For i = 0 To 1
            sprReal(i).Row = 1
            sprReal(i).Col = 1
            sprReal(i).Row2 = sprReal(i).MaxRows - 1
            sprReal(i).Col2 = 10
        
            ' Set sort definition for key 1
            sprReal(i).SortBy = SS_SORT_BY_ROW

            sprReal(i).SortKey(1) = 1
            sprReal(i).SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
            sprReal(i).Action = SS_ACTION_SORT
        Next
    
    Case 1
        For i = 0 To 1
            sprReal(i).Row = 1
            sprReal(i).Col = 1
            sprReal(i).Row2 = sprReal(i).MaxRows - 1
            sprReal(i).Col2 = 10
        
            ' Set sort definition for key 1
            sprReal(i).SortBy = SS_SORT_BY_ROW
        
            sprReal(i).SortKey(1) = 3
            sprReal(i).SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
            sprReal(i).Action = SS_ACTION_SORT
        Next

End Select


End Sub

Private Sub sprReal_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
Dim tit As String, Esp As String
Dim sql As String
Dim sec As Variant
Dim secDescrip As Variant
Dim ActMtr As Variant, AntMtr As Variant
 
sprReal(Index).GetText 10, Row, sec
sprReal(Index).GetText 1, Row, secDescrip

sprReal(Index).GetText 2, Row, ActMtr
sprReal(Index).GetText 4, Row, AntMtr


If Col = 9 Then
     sql = " SELECT DESCRIP,SUM(actMtr) actMtr,SUM(actImp) actImp, SUM(antMtr) antMtr, SUM(antImp) antImp "
     sql = sql & " FROM( SELECT descrip||'-'||cod_rubro DESCRIP "
     sql = sql & "  ,METROS_TOTAL_SECTOR*metros_porc_prov/100 actMtr "
     sql = sql & "  ,sum(v.IMPORTE_PROV) actImp , 0 antMtr, 0 antImp "
     sql = sql & " FROM ESTADIS.Z_VENTA_SECTOR v, baires.proveedor p "
     sql = sql & " where fch_ticket between " & Func.func_ToDate(mskFDesde.FormattedText) & " And " & Func.func_ToDate(mskFHasta.FormattedText)
     sql = sql & " and v.cod_prov = p.cod_prov  and cod_sector = " & sec
     sql = sql & " " & L_Armarcondicion
     sql = sql & " group by  METROS_TOTAL_SECTOR,descrip||'-'||cod_rubro,metros_porc_prov "
     sql = sql & " Union "
     sql = sql & " SELECT descrip||'-'||cod_rubro ,0 actMtr, 0 actImp "
     sql = sql & "  ,METROS_TOTAL_SECTOR*metros_porc_prov/100 antMtr, sum(v.IMPORTE_PROV) antImp "
     sql = sql & "  FROM ESTADIS.Z_VENTA_SECTOR v, baires.proveedor p "
     sql = sql & " where fch_ticket between " & Func.func_ToDate(mskFDesdeAnt.FormattedText) & " And " & Func.func_ToDate(mskFHastaAnt.FormattedText)
     sql = sql & " and v.cod_prov = p.cod_prov  and cod_sector = " & sec
     sql = sql & " " & L_Armarcondicion
     sql = sql & "  group by  METROS_TOTAL_SECTOR,descrip||'-'||cod_rubro,metros_porc_prov) "
     sql = sql & " GROUP BY DESCRIP "
     sql = sql & " "
    
    If Index = 0 Then
     tit = "Detalle ventas netas "
     frmVentaProvMetro.Mostrar sql, tit & secDescrip, "Metros p.actual: " & ActMtr & "  Metros p.anterir: " & AntMtr, Trim(sec), Index
    Else
     tit = "Detalle productividad por Mtr.2 "
    End If

    
End If



End Sub

Private Sub tabTotal_Click(PreviousTab As Integer)

On Error GoTo ErrT:

    'sprTotal(tabTotal.Tab).SetFocus
    
ErrT:
    Exit Sub
End Sub

