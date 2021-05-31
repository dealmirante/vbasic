VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAcum 
   Caption         =   "Acumulado de Ventas Diarias"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   7575
   Begin VB.CommandButton botExcel 
      Caption         =   "Excel"
      Height          =   405
      Left            =   5700
      Picture         =   "frmAcum.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   105
      Width           =   525
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
      Height          =   405
      Index           =   2
      Left            =   6225
      Picture         =   "frmAcum.frx":0592
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   105
      Width           =   570
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
      Height          =   405
      Index           =   1
      Left            =   5145
      Picture         =   "frmAcum.frx":0DB4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   105
      Width           =   570
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
      Height          =   405
      Index           =   0
      Left            =   4575
      Picture         =   "frmAcum.frx":0EB6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   105
      Width           =   570
   End
   Begin VB.Frame frdatos 
      Caption         =   "Período "
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
      Height          =   1335
      Left            =   165
      TabIndex        =   0
      Top             =   -15
      Width           =   6825
      Begin VB.ComboBox CboCodAeropuerto 
         Height          =   315
         ItemData        =   "frmAcum.frx":0FB8
         Left            =   4740
         List            =   "frmAcum.frx":0FBA
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   585
         Width           =   1995
      End
      Begin VB.ComboBox CboEspigon 
         Height          =   315
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   930
         Width           =   2010
      End
      Begin VB.ListBox LstEspigon 
         Height          =   255
         Left            =   7050
         TabIndex        =   8
         Top             =   975
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   2535
         Picture         =   "frmAcum.frx":0FBC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2550
         Picture         =   "frmAcum.frx":112E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   300
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1395
         TabIndex        =   3
         Top             =   315
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFHasta 
         Height          =   285
         Left            =   1395
         TabIndex        =   4
         Top             =   720
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.Label LblCodAeropuerto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cod. Aeropuerto :"
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
         Left            =   3105
         TabIndex        =   12
         Top             =   585
         Width           =   1590
      End
      Begin VB.Label LblEspigon 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Espigón :"
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
         Left            =   3120
         TabIndex        =   11
         Top             =   945
         Width           =   1575
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
         Left            =   135
         TabIndex        =   6
         Top             =   720
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
         Index           =   4
         Left            =   150
         TabIndex        =   5
         Top             =   315
         Width           =   1185
      End
   End
   Begin TabDlg.SSTab tabEst 
      Height          =   3330
      Left            =   165
      TabIndex        =   7
      Top             =   1395
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   5874
      _Version        =   327680
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "TOT. Mes"
      TabPicture(0)   =   "frmAcum.frx":12A0
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "sprTot"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Grupo A"
      TabPicture(1)   =   "frmAcum.frx":12BC
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sprGr(0)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Grupo B"
      TabPicture(2)   =   "frmAcum.frx":12D8
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "sprGr(1)"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "Grupo C"
      TabPicture(3)   =   "frmAcum.frx":12F4
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "sprGr(2)"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "% Part. Loc"
      TabPicture(4)   =   "frmAcum.frx":1310
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "sprPart(0)"
      Tab(4).Control(0).Enabled=   0   'False
      TabCaption(5)   =   "% Part. Gr."
      TabPicture(5)   =   "frmAcum.frx":132C
      Tab(5).ControlCount=   1
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "sprPart(1)"
      Tab(5).Control(0).Enabled=   0   'False
      Begin FPSpread.vaSpread sprPart 
         Height          =   2415
         Index           =   0
         Left            =   -74310
         OleObjectBlob   =   "frmAcum.frx":1348
         TabIndex        =   16
         Top             =   615
         Width           =   4995
      End
      Begin FPSpread.vaSpread sprTot 
         Height          =   2400
         Left            =   630
         OleObjectBlob   =   "frmAcum.frx":17A8
         TabIndex        =   17
         Top             =   630
         Width           =   5475
      End
      Begin FPSpread.vaSpread sprGr 
         Height          =   2400
         Index           =   0
         Left            =   -74355
         OleObjectBlob   =   "frmAcum.frx":1C82
         TabIndex        =   18
         Top             =   630
         Width           =   5475
      End
      Begin FPSpread.vaSpread sprGr 
         Height          =   2400
         Index           =   1
         Left            =   -74355
         OleObjectBlob   =   "frmAcum.frx":2080
         TabIndex        =   19
         Top             =   615
         Width           =   5475
      End
      Begin FPSpread.vaSpread sprGr 
         Height          =   2400
         Index           =   2
         Left            =   -74400
         OleObjectBlob   =   "frmAcum.frx":247E
         TabIndex        =   20
         Top             =   600
         Width           =   5475
      End
      Begin FPSpread.vaSpread sprPart 
         Height          =   2415
         Index           =   1
         Left            =   -74310
         OleObjectBlob   =   "frmAcum.frx":287C
         TabIndex        =   21
         Top             =   615
         Width           =   4995
      End
   End
End
Attribute VB_Name = "frmAcum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsVta As Recordset
Dim rsUni As Recordset

Private Sub L_DecoGrupoVta()
Dim fila As Integer, i
Dim valor As Variant
Dim f_a As Integer, f_b As Integer, f_e As Integer

f_a = 1
f_b = 1
f_e = 1
Do While Not rsVta.EOF
    Select Case rsVta!Grupo_Venta
        Case "A"
'            sprGr(0).MaxRows = sprGr(0).MaxRows + 1
'            sprGr(0).SetText 1, sprGr(0).MaxRows, Trim(rsVta!cod_local) & str(rsVta!cod_sloc)
            sprGr(0).GetText 1, f_a, valor
            If valor = rsVta!Cod_Local & str(rsVta!Cod_Sloc) Then
                sprGr(0).SetText 2, f_a, str(rsVta!imp)
                sprGr(0).SetText 3, f_a, str(rsVta!tick)
                
                rsVta.MoveNext
            End If
            f_a = f_a + 1
        Case "B"
'            sprGr(1).MaxRows = sprGr(1).MaxRows + 1
'            sprGr(1).SetText 1, sprGr(1).MaxRows, Trim(rsVta!cod_local) & str(rsVta!cod_sloc)
            
            sprGr(1).GetText 1, f_b, valor
            If valor = rsVta!Cod_Local & str(rsVta!Cod_Sloc) Then
                sprGr(1).SetText 2, f_b, str(rsVta!imp)
                sprGr(1).SetText 3, f_b, str(rsVta!tick)
                
                rsVta.MoveNext
            End If
            f_b = f_b + 1
        Case "C"
'            sprGr(2).MaxRows = sprGr(2).MaxRows + 1
'            sprGr(2).SetText 1, sprGr(2).MaxRows, Trim(rsVta!cod_local) & str(rsVta!cod_sloc)
            sprGr(2).GetText 1, f_e, valor
            If valor = rsVta!Cod_Local & str(rsVta!Cod_Sloc) Then
                sprGr(2).SetText 2, f_e, str(rsVta!imp)
                sprGr(2).SetText 3, f_e, str(rsVta!tick)
                
                rsVta.MoveNext
            End If
            f_e = f_e + 1
        Case " "
            rsVta.MoveNext
    End Select
    
Loop

End Sub

Private Sub L_DecoUni()
Dim f_a As Integer, f_b As Integer, f_e As Integer, i
Dim valor As Variant

f_a = 1
f_b = 1
f_e = 1
Do While Not rsUni.EOF

    Select Case rsUni!Grupo_Venta
        Case "A"
            sprGr(0).GetText 1, f_a, valor
            If valor = rsUni!Cod_Local & str(rsUni!Cod_Sloc) Then
                sprGr(0).SetText 4, f_a, str(rsUni!cant)
                rsUni.MoveNext
            End If
            f_a = f_a + 1
        Case "B"
            sprGr(1).GetText 1, f_b, valor
            If valor = rsUni!Cod_Local & str(rsUni!Cod_Sloc) Then
                sprGr(1).SetText 4, f_b, str(rsUni!cant)
                rsUni.MoveNext
            End If
            f_b = f_b + 1
        Case "C"
            sprGr(2).GetText 1, f_e, valor
            If valor = rsUni!Cod_Local & str(rsUni!Cod_Sloc) Then
                sprGr(2).SetText 4, f_e, str(rsUni!cant)
                rsUni.MoveNext
            End If
            f_e = f_e + 1
        Case Else
            rsUni.MoveNext
    End Select
    
    
Loop

End Sub


Private Sub L_PonerLocales()
Dim sql As String
Dim RS As Recordset

sql = "SELECT cod_depn,cod_sdep,cod_loc,cod_sloc " _
& " FROM ventas.local_sublocal WHERE cod_depn = '" & CboCodAeropuerto.Text & "' " _
& " and cod_sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "' " _
& "  " _
& " ORDER BY cod_loc,cod_sloc "


If Aplicacion.ObtenerRsDAO(sql, RS) Then
    Do While Not RS.EOF
        sprGr(0).MaxRows = sprGr(0).MaxRows + 1
        sprGr(1).MaxRows = sprGr(1).MaxRows + 1
        sprGr(2).MaxRows = sprGr(2).MaxRows + 1
        
        sprGr(0).SetText 1, sprGr(0).MaxRows, Trim(RS!cod_loc) & str(RS!Cod_Sloc)
        sprGr(1).SetText 1, sprGr(1).MaxRows, Trim(RS!cod_loc) & str(RS!Cod_Sloc)
        sprGr(2).SetText 1, sprGr(2).MaxRows, Trim(RS!cod_loc) & str(RS!Cod_Sloc)
        RS.MoveNext
    Loop
    Aplicacion.CerrarDAO RS
End If



End Sub

Private Sub L_PorcPartGr()
Dim i, j
Dim valor As Variant, tot As Variant

On Error GoTo ErrPart:

For i = 1 To SprTot.MaxRows
    SprTot.GetText 2, i, tot
    sprPart(1).MaxRows = sprPart(1).MaxRows + 1
    SprTot.GetText 1, i, valor
    sprPart(1).SetText 1, i, valor
    sprPart(1).SetText sprPart(1).MaxCols, i, "100"
    If tot > 0 Then
    For j = 0 To 2
        sprGr(j).GetText 2, i, valor
        sprPart(1).SetText j + 2, i, str(100 * valor / tot)
    Next
    End If
Next
Spread.Spread_TotalesLinea sprPart(1)

ErrPart:
    Exit Sub

End Sub

Private Sub L_TratarExcel(titulo As String)
Dim AppExcel As Object
Dim titCol() As String, part(0 To 1) As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer, filaant As Integer
Dim i
Dim tit As Variant
Dim NOMBRE As String

'On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmAcum.caption = Aplicacion.SeteoProceso(frmAcum.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.application.Visible = True
    
    part(0) = " por Local "
    part(1) = " por Grupo "
    ReDim titCol(SprTot.MaxCols)
    Col = 1
    fila = 4
    
    For i = 1 To SprTot.MaxCols
        SprTot.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 2, 1, titulo
    rango = Exl_rangos(2, 2, 1, SprTot.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, "Acumulado de ventas diarias (" & mskFDesde.FormattedText & "-" & mskFHasta.FormattedText & ")"
    rango = Exl_rangos(fila, fila, 1, SprTot.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, "TOTALES"
    rango = Exl_rangos(fila, fila, 1, SprTot.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    fila = fila + 1
    Exl_BajarGrillaExel SprTot, AppExcel, fila, 1, titCol
    rango = Exl_rangos(fila + 1, fila + SprTot.MaxRows, Col + 1, SprTot.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + SprTot.MaxRows
'    AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
    rango = Exl_rangos(fila, fila, Col, SprTot.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    For i = 0 To 2
        fila = fila + 2
        Exl_PonerValor AppExcel, fila, Col, "GRUPO " & Chr(65 + i)
        rango = Exl_rangos(fila, fila, 1, SprTot.MaxCols)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        fila = fila + 1
        Exl_BajarGrillaExel sprGr(i), AppExcel, fila, 1, titCol
        rango = Exl_rangos(fila + 1, fila + SprTot.MaxRows, Col + 1, SprTot.MaxCols)
        Exl_Format AppExcel, rango
        fila = fila + SprTot.MaxRows
    '    AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
        rango = Exl_rangos(fila, fila, Col, SprTot.MaxCols)
        Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Next
    
    For i = 1 To sprPart(0).MaxCols
        sprPart(0).GetText i, 0, tit
        titCol(i) = tit
    Next

    For i = 0 To 1
        fila = fila + 2
        Exl_PonerValor AppExcel, fila, Col, "% PARTICIPACION " & part(i)
        rango = Exl_rangos(fila, fila, 1, SprTot.MaxCols)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        fila = fila + 1
        Exl_BajarGrillaExel sprPart(i), AppExcel, fila, 1, titCol
        rango = Exl_rangos(fila + 1, fila + SprTot.MaxRows, Col + 1, SprTot.MaxCols)
        Exl_Format AppExcel, rango
        fila = fila + SprTot.MaxRows
    '    AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
        rango = Exl_rangos(fila, fila, Col, SprTot.MaxCols)
        Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Next
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    'AppExcel.application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If

    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmAcum.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub L_ProcPartLocal()
Dim i, j
Dim valor As Variant
Dim tot As Variant

'On Error GoTo ErrPart:
tot = 0

SprTot.GetText 2, SprTot.MaxRows, tot
For j = 1 To SprTot.MaxRows - 1
    SprTot.GetText 1, j, valor
    sprPart(0).MaxRows = 1 + sprPart(0).MaxRows
    sprPart(0).SetText 1, sprPart(0).MaxRows, Trim(valor)
    SprTot.GetText 2, j, valor
    sprPart(0).SetText sprPart(0).MaxCols, j, str(100 * valor / tot)
Next

For i = 0 To 2
    sprGr(i).GetText 2, sprGr(i).MaxRows, tot
    If tot > 0 Then
    For j = 1 To sprGr(i).MaxRows - 1
        sprGr(i).GetText 2, j, valor
        sprPart(0).SetText 2 + i, j, str(100 * valor / tot)
    Next
    End If
Next

sprPart(0).MaxRows = sprPart(0).MaxRows + 1
Spread.Spread_TotalesLinea sprPart(0)
For i = 2 To sprPart(0).MaxCols
    sprPart(0).SetText i, sprPart(0).MaxRows, "100"
Next

ErrPart:
    Exit Sub
End Sub

Private Sub L_SumatoriaGrillas(ind As Integer, sprgrT As control)
Dim i, j
Dim valor As Variant, tot As Single

tot = 0
For i = 1 To sprGr(ind).MaxRows
    sprGr(ind).GetText 1, i, valor
    sprgrT.MaxRows = sprgrT.MaxRows + 1
    sprgrT.SetText 1, i, Trim(valor)
    
    For j = 2 To 4
        sprGr(ind).GetText j, i, valor
        tot = valor
        sprGr(ind + 1).GetText j, i, valor
        tot = tot + valor
        sprGr(ind + 2).GetText j, i, valor
        tot = tot + valor
        sprgrT.SetText j, i, str(tot)
    Next
    
Next

Spread.Spread_TotalesGrillas SprTot, 3, 2
For i = 0 To 2
    Spread.Spread_TotalesGrillas sprGr(i), 3, 2
Next

End Sub


Private Sub L_LimpiarGrilla()
Dim i

For i = 0 To 2
    sprGr(i).MaxRows = 0
Next
SprTot.MaxRows = 0
sprPart(0).MaxRows = 0
sprPart(1).MaxRows = 0
End Sub

Private Sub L_Refrescar()
Dim sql As String

'On Error GoTo ErrAcum:

frmAcum.caption = Aplicacion.SeteoProceso(frmAcum.caption)

L_PonerLocales

sql = " SELECT "
sql = sql & " cod_local, "
sql = sql & " cod_sloc, "
sql = sql & " grupo_venta, "
sql = sql & " sum(cantidad) tick, "
sql = sql & " sum(importe) imp "
sql = sql & "FROM " & funcLocal_Vista("venta_lgi", Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion("Vta")
sql = sql & "group by cod_local,cod_sloc,grupo_venta "
sql = sql & " order by grupo_venta,cod_local,cod_sloc"

If Aplicacion.ObtenerRsDAO(sql, rsVta) Then
    If Aplicacion.CantReg(rsVta) > 0 Then
        L_DecoGrupoVta
            L_RefrescarUni
        L_SumatoriaGrillas 0, SprTot
        L_ProcPartLocal
        L_PorcPartGr
        botEjecutar(0).Enabled = False
        frdatos.Enabled = False
    Else
    End If
    Aplicacion.CerrarDAO rsVta
End If

ErrAcum:
    frmAcum.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub

Private Sub L_RefrescarUni()
Dim sql As String


sql = " SELECT "
sql = sql & " cod_local, "
sql = sql & " cod_sloc, "
sql = sql & " grupo_venta, "
sql = sql & " sum(cantidad) cant "
sql = sql & "FROM " & funcLocal_Vista("venta_plg", Year(mskFDesde.FormattedText))
sql = sql & L_Armarcondicion("Uni")
sql = sql & "group by cod_local,cod_sloc,grupo_venta "
sql = sql & " order by grupo_venta,cod_local,cod_sloc "

If Aplicacion.ObtenerRsDAO(sql, rsUni) Then
    If Aplicacion.CantReg(rsUni) > 0 Then
        L_DecoUni
    End If
    Aplicacion.CerrarDAO rsUni
End If

End Sub


Private Function L_Armarcondicion(Tipo As String) As String
Dim Cond As String, condLoc As String
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant, i

fechaDesde = mskFDesde.FormattedText
fechaHasta = mskFHasta.FormattedText

Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta) _
        & " and cod_depn = '" & CboCodAeropuerto.Text & "' " _
        & " and cod_sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "' "
        
        If Tipo = "Vta" Then
            Cond = Cond & " and comitente = 'T' "
        End If
        

L_Armarcondicion = Cond

End Function


Private Sub botEjecutar_Click(Index As Integer)


Select Case Index
    Case 0
        frdatos.Enabled = False
        L_Refrescar
    Case 1
        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        'tabEspigon.Enabled = False
        L_LimpiarGrilla
    Case 2
        Unload Me
End Select


End Sub

Private Sub botExcel_Click()

Select Case LstEspigon.List(CboEspigon.ListIndex)
    Case "INTA"
        L_TratarExcel "ESPIGON INTERNACIONAL A"
    Case "INTB"
        L_TratarExcel "ESPIGON INTERNACIONAL B"
    Case "AEP"
        L_TratarExcel "AEROPARQUE"
End Select
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


Private Sub CboCodAeropuerto_Click()
 Dim sql As String

 sql = " SELECT cod_sdep,descrip FROM baires.subdependencia "
 sql = sql & " WHERE cod_depn = '" & CboCodAeropuerto.Text & "'"
 sql = sql & " ORDER BY cod_sdep"
 
 FuncCbos_LlenarCboLst CboEspigon, LstEspigon, sql

End Sub


Private Sub Form_Load()
Dim sql As String

Width = Screen.Width * 0.6
Height = Screen.Height * 0.58
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height - 500) / 2

mskFDesde.Text = Func.func_Dia1SegunMes_Anio(Month(Date), Year(Date))
mskFHasta.Text = Format$(Date, FTOFECHA)

sql = " SELECT cod_depn,descrip FROM baires.dependencia "
sql = sql & " ORDER BY cod_depn"

FuncCbos_LlenarCbo CboCodAeropuerto, sql

L_LimpiarGrilla

End Sub


Private Sub mskFDesde_LostFocus()

    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    Else
    'If (Year(mskFDesde.FormattedText) < Year(Date)) Then
    '(Month(mskFDesde.FormattedText) < Month(Date))
    '    mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    End If
    
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)


End Sub


Private Sub mskFHasta_LostFocus()

If Not IsDate(mskFHasta.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
Else
'If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) _
'        And Month(mskFHasta.FormattedText) <> Month(mskFHasta.FormattedText) Then
'
'        mskFHasta.Text = mskFDesde.Text
End If

mskFHasta.Text = Format$(mskFHasta.FormattedText, FTOFECHA)
If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
    mskFHasta.Text = mskFDesde.FormattedText
End If

End Sub


