VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVsLocRub 
   Caption         =   "Estimados vs. Ventas por Local-Rubro"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   8625
   Begin VB.CommandButton botExcel 
      Caption         =   "Excel"
      Height          =   540
      Left            =   8010
      Picture         =   "frmVsLocRub.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1785
      Width           =   615
   End
   Begin VB.Frame frMod 
      Caption         =   "Modelos"
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
      Height          =   2130
      Left            =   5400
      TabIndex        =   14
      Top             =   -30
      Width           =   2535
      Begin MSMask.MaskEdBox mskDiaDesde 
         Height          =   285
         Left            =   1485
         TabIndex        =   30
         Top             =   1425
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.ComboBox cboMod 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   990
         Width           =   2295
      End
      Begin VB.TextBox txtMesMod 
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
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   285
         Width           =   735
      End
      Begin MSMask.MaskEdBox mskDiaHasta 
         Height          =   285
         Left            =   1470
         TabIndex        =   31
         Top             =   1755
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "# Día hasta"
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
         Index           =   5
         Left            =   165
         TabIndex        =   33
         Top             =   1755
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "# Día desde"
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
         Left            =   165
         TabIndex        =   32
         Top             =   1425
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Modelo"
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
         Left            =   165
         TabIndex        =   27
         Top             =   645
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
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   285
         Width           =   780
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
      Height          =   480
      Index           =   0
      Left            =   8010
      Picture         =   "frmVsLocRub.frx":0592
      Style           =   1  'Graphical
      TabIndex        =   22
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
      Left            =   7995
      Picture         =   "frmVsLocRub.frx":0694
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   75
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
      Height          =   510
      Index           =   2
      Left            =   8010
      Picture         =   "frmVsLocRub.frx":0796
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1200
      Width           =   615
   End
   Begin TabDlg.SSTab tabEst 
      Height          =   3330
      Left            =   15
      TabIndex        =   18
      Top             =   2175
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   5874
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   7
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "TOTAL Mes"
      TabPicture(0)   =   "frmVsLocRub.frx":0FB8
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "sprTotal"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "T. Mañana"
      TabPicture(1)   =   "frmVsLocRub.frx":0FD4
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sprA"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "T. Tarde"
      TabPicture(2)   =   "frmVsLocRub.frx":0FF0
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "sprB"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "T. Noche"
      TabPicture(3)   =   "frmVsLocRub.frx":100C
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "sprC"
      Tab(3).Control(0).Enabled=   0   'False
      Begin FPSpread.vaSpread sprTotal 
         Height          =   2730
         Left            =   165
         OleObjectBlob   =   "frmVsLocRub.frx":1028
         TabIndex        =   19
         Top             =   450
         Width           =   8295
      End
      Begin FPSpread.vaSpread sprA 
         Height          =   2730
         Left            =   -74850
         OleObjectBlob   =   "frmVsLocRub.frx":17AC
         TabIndex        =   24
         Top             =   435
         Width           =   8295
      End
      Begin FPSpread.vaSpread sprB 
         Height          =   2730
         Left            =   -74850
         OleObjectBlob   =   "frmVsLocRub.frx":1F2E
         TabIndex        =   25
         Top             =   420
         Width           =   8295
      End
      Begin FPSpread.vaSpread sprC 
         Height          =   2730
         Left            =   -74895
         OleObjectBlob   =   "frmVsLocRub.frx":26B0
         TabIndex        =   26
         Top             =   450
         Width           =   8295
      End
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
      Height          =   2115
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   5280
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1680
         Width           =   2070
      End
      Begin VB.Frame Frame2 
         Caption         =   "Comitente"
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
         Height          =   1710
         Index           =   1
         Left            =   3375
         TabIndex        =   10
         Top             =   210
         Width           =   1710
         Begin VB.OptionButton optComi 
            Caption         =   "General"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   3
            Left            =   150
            TabIndex        =   13
            Top             =   315
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton optComi 
            Caption         =   "IOSC y Propios"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   12
            Top             =   1185
            Width           =   1500
         End
         Begin VB.OptionButton optComi 
            Caption         =   "No IOSC"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   11
            Top             =   720
            Width           =   1080
         End
      End
      Begin VB.ComboBox CboEspigon 
         Height          =   315
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1245
         Width           =   2070
      End
      Begin VB.ComboBox CboCodAeropuerto 
         Height          =   315
         ItemData        =   "frmVsLocRub.frx":2E31
         Left            =   1230
         List            =   "frmVsLocRub.frx":2E33
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   765
         Width           =   2070
      End
      Begin VB.ListBox LstEspigon 
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   915
         Visible         =   0   'False
         Width           =   135
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   300
         Left            =   945
         TabIndex        =   4
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskMes 
         Height          =   300
         Left            =   2550
         TabIndex        =   5
         Top             =   315
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.Label LblEspigon 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rubro"
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
         Index           =   1
         Left            =   135
         TabIndex        =   28
         Top             =   1695
         Width           =   1065
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
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   1245
         Width           =   1065
      End
      Begin VB.Label LblCodAeropuerto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aeropuerto"
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
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   1080
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
         Left            =   1755
         TabIndex        =   7
         Top             =   315
         Width           =   780
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
   End
End
Attribute VB_Name = "frmVsLocRub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As Recordset
Dim rsEstim As Recordset

Dim TodosRubro As Boolean
Private Function L_Armarcondicion()
Dim Cond, i
Dim fechaDesde As String
Dim fechaHasta As String
Dim auxCond

'fechaDesde = func_Dia1SegunMes_Anio(mskMes.Text, mskAnio.Text)
'If mskMes.Text = Month(Date) Then
'    fechaHasta = Format$(Date, FTOFECHA)
'Else
'    fechaHasta = func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text)
'End If
fechaDesde = mskDiaDesde.Text & "-" & mskMes.Text & "-" & mskAnio.Text
fechaHasta = mskDiaHasta.Text & "-" & mskMes.Text & "-" & mskAnio.Text

Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)
Cond = Cond & " and v.cod_depn = a.cod_depn and v.cod_sdep = v.cod_sdep and v.cod_local = a.cod_local "

If CboCodAeropuerto.Text <> "" Then
    Cond = Cond & " and v.cod_depn = '" & CboCodAeropuerto.Text & "'"
End If

If CboEspigon.Text <> "" Then
    Cond = Cond & " and cod_ssdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
End If

If cboRubro.Text <> "" Then
    Cond = Cond & " and cod_rubr = '" & cboRubro.Text & "'"
End If

If optComi(0).Value Then
    Cond = Cond & " and Comitente NOT IN ('IBAIR','IOSC') "
ElseIf optComi(1).Value Then
    Cond = Cond & " and Comitente  IN ('IBAIR','IOSC') "
End If

L_Armarcondicion = Cond

End Function

Private Function L_ArmarcondicionEstim()
Dim Cond, i
Dim auxCond

Cond = " WHERE anio = " & mskAnio.Text
Cond = Cond & " And Mes = " & mskMes.Text
Cond = Cond & " And Tipo_porc = 'I' "
Cond = Cond & " And Descrip = '" & cboMod.Text & "'"

If CboCodAeropuerto.Text <> "" Then
    Cond = Cond & " and depn = '" & CboCodAeropuerto.Text & "'"
End If

If CboEspigon.Text <> "" Then
    Cond = Cond & " and sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
End If

If cboRubro.Text <> "" Then
    Cond = Cond & " and rubro = '" & cboRubro.Text & "'"
End If

If optComi(0).Value Then
    Cond = Cond & " and tipo_Comi = 'N' "
ElseIf optComi(1).Value Then
    Cond = Cond & " and tipo_Comi = 'I' "
End If

L_ArmarcondicionEstim = Cond

End Function
Private Sub L_SetearGrillasRub()
Dim SQL As String
Dim rs As Recordset

SQL = "select cod_rubr from baires.rubro order by cod_rubr "
If Aplicacion.ObtenerRsDAO(SQL, rs) Then
    
    Do While Not rs.EOF
        SprTotal.MaxRows = SprTotal.MaxRows + 1
        SprTotal.SetText 2, SprTotal.MaxRows, Trim(rs!cod_rubr)
            
        rs.MoveNext
    Loop
    
    Aplicacion.CerrarDAO rs
End If

End Sub

Private Sub L_LlenarGrillasRub(col As Integer)
Dim fila As Integer
Dim rub As Variant, RubGr As Variant
Dim total As Double

fila = 1
Do While fila <= SprTotal.MaxRows
        
    SprTotal.GetText 1, fila, rub
    total = 0
    Do While rs!cod_local = rub

            Select Case rs!Grupo
                Case "A"
                    If sprA.MaxRows > 0 Then
                    sprA.SetText col, fila, str(rs!imp)
                    End If
                Case "B"
                    If sprB.MaxRows > 0 Then
                    sprB.SetText col, fila, str(rs!imp)
                    End If
                Case "C"
                    If sprC.MaxRows > 0 Then
                    sprC.SetText col, fila, str(rs!imp)
                    End If
            End Select
            total = total + rs!imp
            rs.MoveNext
            If rs.EOF Then Exit Do
               
    Loop
        
    SprTotal.SetText col, fila, str(total)

    Spread_TotalesGrillaAcum SprTotal, 2, 6, fila
    spread_ResaltarCelda SprTotal, 4, fila
    spread_ResaltarCelda SprTotal, 5, fila
    spread_ResaltarCelda SprTotal, 8, fila
    spread_ResaltarCelda SprTotal, 9, fila
    
    If sprA.MaxRows > 0 Then
        Spread_TotalesGrillaAcum sprA, 2, 6, fila
        spread_ResaltarCelda sprA, 4, fila
        spread_ResaltarCelda sprA, 5, fila
        spread_ResaltarCelda sprA, 8, fila
        spread_ResaltarCelda sprA, 9, fila
    End If
    
    If sprB.MaxRows > 0 Then
        Spread_TotalesGrillaAcum sprB, 2, 6, fila
        spread_ResaltarCelda sprB, 4, fila
        spread_ResaltarCelda sprB, 5, fila
        spread_ResaltarCelda sprB, 8, fila
        spread_ResaltarCelda sprB, 9, fila
    End If
    
    If sprC.MaxRows > 0 Then
        Spread_TotalesGrillaAcum sprC, 2, 6, fila
        spread_ResaltarCelda sprC, 4, fila
        spread_ResaltarCelda sprC, 5, fila
        spread_ResaltarCelda sprC, 8, fila
        spread_ResaltarCelda sprC, 9, fila
    End If
    
    fila = fila + 1
Loop
    If col = 2 Then
        Spread_TotalesGrillas SprTotal, SprTotal.MaxCols - 6, 2
        Spread_TotalesGrillas sprA, sprA.MaxCols - 6, 2
        Spread_TotalesGrillas sprB, sprB.MaxCols - 6, 2
        Spread_TotalesGrillas sprC, sprC.MaxCols - 6, 2
        
        spread_ResaltarCelda SprTotal, 4, SprTotal.MaxRows
        spread_ResaltarCelda SprTotal, 5, SprTotal.MaxRows
        spread_ResaltarCelda sprA, 4, sprA.MaxRows
        spread_ResaltarCelda sprA, 5, sprA.MaxRows
        spread_ResaltarCelda sprB, 4, sprB.MaxRows
        spread_ResaltarCelda sprB, 5, sprB.MaxRows
        spread_ResaltarCelda sprC, 4, sprC.MaxRows
        spread_ResaltarCelda sprC, 5, sprC.MaxRows
        
    End If
End Sub


Private Sub L_LlenarGrillasEstimRub()
Dim rub As String
Dim rsGr As Recordset
Dim SQL As String
Dim total As Double

'No hago esta consulta junto con Dia porque hay que separar los
'casos de aquellos modelos que no tienen datos por grupo

SQL = "Select local,grupo, "
SQL = SQL & " sum(porc) PORC"
If Not optComi(3).Value Then
    SQL = SQL & " From estadis.porcentaje_dgr_io "
Else
    SQL = SQL & " From estadis.porcentaje_dgrl "
End If
SQL = SQL & L_ArmarcondicionEstim
'If Month(Date) = mskMes.Text Then
'    sql = sql & " AND DIA BETWEEN " & func_ToDate("01-" & Format$(Date, "mm-yyyy")) & " AND " & func_ToDate(Format$(Date - 1, FTOFECHA))
'Else
'    sql = sql & " AND DIA BETWEEN " & func_ToDate(func_Dia1SegunMes_Anio(mskMes.Text, mskAnio.Text)) & " AND " & func_ToDate(Format$(func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text), FTOFECHA))
'End If

SQL = SQL & " AND DIA BETWEEN " _
& func_ToDate(mskDiaDesde.Text & "-" & mskMes.Text & "-" & mskAnio.Text) _
& " AND " _
& func_ToDate(mskDiaHasta.Text & "-" & mskMes.Text & "-" & mskAnio.Text)

SQL = SQL & " GROUP BY local,grupo "

Call Aplicacion.ObtenerRsDAO(SQL, rsGr)

Do While Not rsEstim.EOF
    SprTotal.MaxRows = SprTotal.MaxRows + 1
    SprTotal.SetText 1, SprTotal.MaxRows, Trim(rsEstim!local)
    
    rub = rsEstim!local
    total = 0
    If Not rsGr.EOF Then
    Do While rub = rsGr!local
        Select Case rsGr!Grupo
            Case "A"
                sprA.MaxRows = sprA.MaxRows + 1
                sprA.SetText 1, sprA.MaxRows, Trim(rub)
                sprA.SetText 3, sprA.MaxRows, str(rsGr!Porc)
                
                If rsGr!Porc > 0 Then
                     Spread_TotalesGrillaAcum sprA, 3, 7, sprA.MaxRows
                End If
                
            Case "B"
                sprB.MaxRows = sprB.MaxRows + 1
                sprB.SetText 1, sprB.MaxRows, Trim(rub)
                sprB.SetText 3, sprB.MaxRows, str(rsGr!Porc)
                
                If rsGr!Porc > 0 Then
                     Spread_TotalesGrillaAcum sprB, 3, 7, sprB.MaxRows
                End If
        
            Case "C"
                sprC.MaxRows = sprC.MaxRows + 1
                sprC.SetText 1, sprC.MaxRows, Trim(rub)
                sprC.SetText 3, sprC.MaxRows, str(rsGr!Porc)
                
                If rsGr!Porc > 0 Then
                     Spread_TotalesGrillaAcum sprC, 3, 7, sprC.MaxRows
                End If
            
        End Select
        'total = total + rsEstim!Porc
        rsGr.MoveNext
        If rsGr.EOF Then
            Exit Do
        End If
    Loop
    End If
    SprTotal.SetText 3, SprTotal.MaxRows, str(rsEstim!Porc)
    
    Spread_TotalesGrillaAcum SprTotal, 3, 7, SprTotal.MaxRows
    
    rsEstim.MoveNext
Loop

Call Aplicacion.CerrarDAO(rsGr)

End Sub


Private Sub L_RefrescarEstimRub()
Dim SQL As String

On Error GoTo ErrEstim:

SQL = "Select local, "
SQL = SQL & " sum(porc) PORC"
If Not optComi(3).Value Then
    SQL = SQL & " From estadis.porcentaje_rd_io "
Else
    SQL = SQL & " From estadis.porcentaje_rdl "
End If
SQL = SQL & L_ArmarcondicionEstim

'If Month(Date) = mskMes.Text Then
'    sql = sql & " AND DIA BETWEEN " & func_ToDate("01-" & Format$(Date, "mm-yyyy")) & " AND " & func_ToDate(Format$(Date - 1, FTOFECHA))
'Else
'    sql = sql & " AND DIA BETWEEN " & func_ToDate(func_Dia1SegunMes_Anio(mskMes.Text, mskAnio.Text)) & " AND " & func_ToDate(Format$(func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text), FTOFECHA))
'End If
    
SQL = SQL & " AND DIA BETWEEN " _
& func_ToDate(mskDiaDesde.Text & "-" & mskMes.Text & "-" & mskAnio.Text) _
& " AND " _
& func_ToDate(mskDiaHasta.Text & "-" & mskMes.Text & "-" & mskAnio.Text)

SQL = SQL & " GROUP BY local "

If Aplicacion.ObtenerRsDAO(SQL, rsEstim) Then
    
    If Aplicacion.CantReg(rsEstim) > 0 Then
        L_LlenarGrillasEstimRub
    End If
    
    Aplicacion.CerrarDAO rsEstim

End If

ErrEstim:
    Exit Sub
End Sub


Private Sub L_RefrescarRub(col As Integer)
Dim SQL As String

On Error GoTo ErrGLC:

frmVsLocRub.caption = Aplicacion.SeteoProceso(frmVsLocRub.caption)

SQL = " SELECT "
SQL = SQL & " v.cod_local,grupo_venta grupo, "
If Not optComi(3).Value Then
    SQL = SQL & " decode(comitente ,'IBAIR','I','IOSC','I','N')  ,"
End If
SQL = SQL & " sum(importe) imp, "
SQL = SQL & " sum(cantidad) cant "
SQL = SQL & "FROM " & funcLocal_Vista("venta_plg", mskAnio.Text)
SQL = SQL & " v, ventas.apertura_sdep a "
SQL = SQL & L_Armarcondicion()
If Not optComi(3).Value Then
    SQL = SQL & " group by v.cod_local,grupo_venta ,decode(comitente ,'IBAIR','I','IOSC','I','N') "
Else
    SQL = SQL & " group by v.cod_local,grupo_venta "
End If
SQL = SQL & " order by v.cod_local "

If Aplicacion.ObtenerRsDAO(SQL, rs) Then
    
    If Aplicacion.CantReg(rs) > 0 Then
        L_LlenarGrillasRub col
        tabEst.Enabled = True
        botEjecutar(0).Enabled = False
        frCab.Enabled = False
        frMod.Enabled = False
    End If

    Aplicacion.CerrarDAO rs

End If

ErrGLC:
    frmVsLocRub.caption = Aplicacion.SeteoFin
    Exit Sub
    


End Sub


Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        If cboMod.Text <> "" Then
            L_RefrescarEstimRub
            L_RefrescarRub 2
        End If
    Case 1
        'Call botPorc_Click(0)
        frCab.Enabled = True
        frMod.Enabled = True
        botEjecutar(0).Enabled = True
        tabEst.Enabled = False
        L_LimpiarGrillas
        'L_RefrescarFrames
    Case 2
        Unload Me
End Select

End Sub

Private Sub L_SetearGrillasDias(anio As Integer, Mes As Integer)
Dim FD As String, FH As String
Dim fdAux As Date, i As Integer


FD = func_Dia1SegunMes_Anio(Mes, anio)
If Mes = Month(Date) Then
    FH = Format$(Date, FTOFECHA)
Else
    FH = func_Dia30SegunMes_Anio(Mes, anio)
End If

fdAux = FD
i = 1
SprTotal.MaxRows = 0
sprA.MaxRows = 0
sprB.MaxRows = 0
sprC.MaxRows = 0
Do While fdAux <= CDate(FH)
        SprTotal.MaxRows = SprTotal.MaxRows + 1
        SprTotal.SetText 1, i, Format$(fdAux, "dd-mm-yy")
        sprA.MaxRows = sprA.MaxRows + 1
        sprA.SetText 1, i, Format$(fdAux, "dd-mm-yy")
        sprB.MaxRows = sprB.MaxRows + 1
        sprB.SetText 1, i, Format$(fdAux, "dd-mm-yy")
        sprC.MaxRows = sprA.MaxRows + 1
        sprC.SetText 1, i, Format$(fdAux, "dd-mm-yy")

        i = i + 1
        fdAux = fdAux + 1

Loop

'Spread_PintarfinSemana sprTotal
'Spread_PintarfinSemana sprA
'Spread_PintarfinSemana sprB
'Spread_PintarfinSemana sprC

End Sub


Private Sub L_LimpiarGrillas()
Dim i

 SprTotal.MaxRows = 0
 sprB.MaxRows = 0
 sprA.MaxRows = 0
 sprC.MaxRows = 0

End Sub

Private Sub botExcel_Click()
Dim Nom As String

Select Case LstEspigon.List(CboEspigon.ListIndex)
    Case "INTA"
        Nom = "ESPIGON INTERNACIONAL A"
    Case "INTB"
        Nom = "ESPIGON INTERNACIONAL B"
    Case "AEP"
        Nom = "AEROPARQUE"
End Select

        L_TratarExcel SprTotal, Nom, "Ventas vs. Estimado -Total Mes- " & mskMes.Text & "-" & mskAnio.Text & " ( al " & Format$(Date - 1, FTOFECHA) & " ) ", "Importes"

End Sub

Private Sub CboCodAeropuerto_Click()
Dim SQL As String

SQL = " SELECT cod_sdep,descrip FROM baires.subdependencia "
SQL = SQL & " WHERE cod_depn = '" & CboCodAeropuerto.Text & "'"
SQL = SQL & " ORDER BY cod_sdep"
 
FuncCbos_LlenarCboLst CboEspigon, LstEspigon, SQL

If CboCodAeropuerto.Text = "EZE" Then
    CboEspigon.AddItem "Inter Llegadas"
    LstEspigon.AddItem "INTAL"
    CboEspigon.AddItem "Inter Salidas"
    LstEspigon.AddItem "INTAS"
End If

cboMod.Clear

End Sub


Private Sub CboCodAeropuerto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    CboCodAeropuerto.ListIndex = -1
End If
End Sub


Private Sub CboEspigon_Click()
Dim SQL As String

SQL = "Select descrip from estadis.porciento_cabezera "
SQL = SQL & " Where anio= " & mskAnio.Text
SQL = SQL & " And mes= " & txtMesMod.Text
SQL = SQL & " And cod_depn= '" & CboCodAeropuerto.Text & "'"
SQL = SQL & " And cod_sdep= '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
SQL = SQL & " And tipo_porc = 'I' "

FuncCbos_LlenarCbo cboMod, SQL



End Sub

Private Sub CboEspigon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    CboEspigon.ListIndex = -1
End If
End Sub


Private Sub cboMod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    cboMod.ListIndex = -1
End If
End Sub


Private Sub cboRubro_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    cboRubro.ListIndex = -1
End If
End Sub

Private Sub Form_Load()
Dim SQL As String
Dim A As Integer

Top = 160
Left = 250
Height = 6100
Width = 9000

SQL = " SELECT cod_depn,descrip FROM baires.dependencia "
SQL = SQL & " ORDER BY cod_depn"

FuncCbos_LlenarCbo CboCodAeropuerto, SQL

SQL = " SELECT cod_rubr,descrip FROM baires.rubro WHERE cod_rubr <> 'REG' "
SQL = SQL & " ORDER BY cod_rubr"

FuncCbos_LlenarCbo cboRubro, SQL

mskAnio.Text = Year(Date)
mskMes.Text = Month(Date)

txtMesMod.Text = Month(Date)

mskDiaDesde.Text = 1
mskDiaHasta.Text = Day(Date - 1)

L_LimpiarGrillas
frmPrincipal.lstForms.AddItem "frmVsRub"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FuncLocal_SacarForm "frmVsRub"
End Sub
Private Sub L_TratarExcel(spr As control, Titulo As String, subTit As String, Info As String)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim fila As Integer, filaant As Integer
Dim i
Dim tit As Variant
Dim nombre As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

frmVsLocRub.caption = Aplicacion.SeteoProceso(frmVsLocRub.caption)

If nombre <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
'    AppExcel.application.Visible = True
    
    ReDim titCol(spr.MaxCols)
    col = 1
    fila = 4
    
    For i = 1 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 2, 1, Titulo
    rango = Exl_rangos(2, 2, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, col, subTit
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    fila = fila + 2
    rango = Exl_rangos(fila, fila + 6, 1, 1)
    
        Exl_PonerValor AppExcel, fila, col, "Modelo :" & cboMod.Text
    
    fila = fila + 2
        
    Exl_PonerValor AppExcel, fila, col, "Información sobre :" & Info
    
    fila = fila + 1
    
    If cboRubro.Text <> "" Then
       Exl_PonerValor AppExcel, fila, col, "Rubro : " & cboRubro.Text
    Else
        Exl_PonerValor AppExcel, fila, col, "Rubro : Todos"
    End If
    
    fila = fila + 2
    
    If optComi(1).Value Then
        Exl_PonerValor AppExcel, fila, col, "Comitente : I.O.S.C y Propios"
    ElseIf optComi(0).Value Then
        Exl_PonerValor AppExcel, fila, col, "Comitente : NO I.O.S.C "
    ElseIf optComi(3).Value Then
        Exl_PonerValor AppExcel, fila, col, "Comitente : Todos "
    End If
    Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
    
    fila = fila + 2
    
    Exl_BajarGrillaExel spr, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + spr.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows
'    AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    If sprA.MaxRows > 1 Then
        fila = fila + 2
        Exl_PonerValor AppExcel, fila, col, "Grupo 'A'"
        rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        
        fila = fila + 2
        Exl_BajarGrillaExel sprA, AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + spr.MaxRows, col + 1, spr.MaxCols)
        Exl_Format AppExcel, rango
        fila = fila + sprA.MaxRows
'        AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
        rango = Exl_rangos(fila, fila, col, spr.MaxCols)
        Exl_ColorInt AppExcel, rango, Exl_Gris
    
        
        fila = fila + 2
        Exl_PonerValor AppExcel, fila, col, "Grupo 'B'"
        rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        
        fila = fila + 2
        Exl_BajarGrillaExel sprB, AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + spr.MaxRows, col + 1, spr.MaxCols)
        Exl_Format AppExcel, rango
        fila = fila + sprA.MaxRows
'        AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
        rango = Exl_rangos(fila, fila, col, spr.MaxCols)
        Exl_ColorInt AppExcel, rango, Exl_Gris
    
        fila = fila + 2
        Exl_PonerValor AppExcel, fila, col, "Grupo 'C'"
        rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        
        fila = fila + 2
        Exl_BajarGrillaExel sprC, AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + spr.MaxRows, col + 1, spr.MaxCols)
        Exl_Format AppExcel, rango
        fila = fila + sprA.MaxRows
'        AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
        rango = Exl_rangos(fila, fila, col, spr.MaxCols)
        Exl_ColorInt AppExcel, rango, Exl_Gris
    
    End If
    Exl.Exl_AnchoCol AppExcel, spr.MaxCols, spr.MaxCols, 1
    Exl.Exl_AnchoCol AppExcel, 1, 1, 8
    Exl.Exl_AnchoCol AppExcel, 5, 5, 7
    Exl.Exl_AnchoCol AppExcel, 9, 9, 7
    
    'AppExcel.application.ActiveSheet.PageSetup.CenterHorizontally = True
    'AppExcel.application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    
    AppExcel.SaveAs nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmVsLocRub.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub mskAnio_LostFocus()
If Val(mskAnio.Text) < 1996 Or Val(mskAnio) > 2050 Then
    mskAnio.Text = Year(Date)
End If

mskDiaDesde.Text = 1
If mskMes.Text = Month(Now) And Year(Now) = mskAnio.Text Then
    mskDiaHasta.Text = Day(Now)
Else
    mskDiaHasta.Text = Day(func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text))
End If

End Sub
Private Sub mskDiaDesde_LostFocus()
Dim dia As Integer

If mskMes.Text = Month(Now) And Year(Now) = mskAnio.Text Then
    dia = Day(Now)
Else
    dia = Day(func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text))
End If

If Val(mskDiaDesde.Text) < 1 Or Val(mskDiaDesde.Text) > dia Then
    mskDiaDesde.Text = 1
End If

If Val(mskDiaDesde.Text) > Val(mskDiaHasta.Text) Then
    mskDiaHasta.Text = mskDiaDesde.Text
End If

End Sub


Private Sub mskDiaHasta_LostFocus()
Dim dia As Integer

If mskMes.Text = Month(Now) And Year(Now) = mskAnio.Text Then
    dia = Day(Now)
Else
    dia = Day(func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text))
End If

If Val(mskDiaHasta.Text) < 1 Or Val(mskDiaHasta.Text) > dia Then
    mskDiaHasta.Text = dia
End If

If Val(mskDiaDesde.Text) > Val(mskDiaHasta.Text) Then
    mskDiaHasta.Text = mskDiaDesde.Text
End If


End Sub


Private Sub mskMes_LostFocus()
Dim SQL As String

If Val(mskMes.Text) < 1 Or Val(mskMes.Text) > 12 Then
    mskMes.Text = Month(Date)
End If
txtMesMod.Text = mskMes.Text

SQL = "Select descrip from estadis.porciento_cabezera "
SQL = SQL & " Where anio= " & mskAnio.Text
SQL = SQL & " And mes= " & txtMesMod.Text
SQL = SQL & " And cod_depn= '" & CboCodAeropuerto.Text & "'"
SQL = SQL & " And cod_sdep= '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
SQL = SQL & " And tipo_porc = 'I' "

FuncCbos_LlenarCbo cboMod, SQL

mskDiaDesde.Text = 1

If mskMes.Text = Month(Now) And Year(Now) = mskAnio.Text Then
    mskDiaHasta.Text = Day(Now)
Else
    mskDiaHasta.Text = Day(func_Dia30SegunMes_Anio(mskMes.Text, mskAnio.Text))
End If
    
End Sub



Private Sub tabEst_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    Select Case tabEst.Tab
        Case 0
            SprTotal.SetFocus
        Case 1
            sprA.SetFocus
        Case 2
            sprB.SetFocus
        Case 3
            sprC.SetFocus
    End Select
    
    
ErrT:
    Exit Sub



End Sub

