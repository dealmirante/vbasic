VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTot 
   Caption         =   "Ventas Total Compañia"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   8265
   Begin VB.CommandButton botExcelAep 
      Caption         =   "Excel"
      Height          =   510
      Left            =   7350
      Picture         =   "frmTot.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4350
      Width           =   765
   End
   Begin VB.CommandButton botPorcTot 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   0
      Left            =   7350
      Picture         =   "frmTot.frx":0592
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1620
      Width           =   765
   End
   Begin VB.CommandButton botPorcTot 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   1
      Left            =   7350
      Picture         =   "frmTot.frx":0F04
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2145
      Width           =   765
   End
   Begin VB.CommandButton botPorcTot 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   2
      Left            =   7350
      Picture         =   "frmTot.frx":155A
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2685
      Width           =   765
   End
   Begin VB.CommandButton botGrTorta 
      Height          =   540
      Left            =   7350
      Picture         =   "frmTot.frx":1C20
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3225
      Width           =   765
   End
   Begin VB.CommandButton botGrEvoTot 
      Height          =   540
      Left            =   7350
      Picture         =   "frmTot.frx":2062
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3810
      Width           =   765
   End
   Begin VB.Frame Frame3 
      Height          =   1545
      Left            =   7080
      TabIndex        =   9
      Top             =   -15
      Width           =   1125
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
         Left            =   255
         Picture         =   "frmTot.frx":24A4
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   195
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
         Left            =   270
         Picture         =   "frmTot.frx":25A6
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   585
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
         Height          =   450
         Index           =   2
         Left            =   270
         Picture         =   "frmTot.frx":26A8
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1005
         Width           =   570
      End
   End
   Begin VB.Frame frdatos 
      Height          =   1545
      Left            =   90
      TabIndex        =   0
      Top             =   -15
      Width           =   6855
      Begin VB.CommandButton botHelpFH 
         Height          =   345
         Left            =   2430
         Picture         =   "frmTot.frx":2ECA
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   855
         Width           =   375
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2445
         Picture         =   "frmTot.frx":303C
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   375
         Width           =   375
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
         Height          =   1215
         Index           =   1
         Left            =   2850
         TabIndex        =   1
         Top             =   180
         Width           =   3810
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
            Left            =   165
            TabIndex        =   8
            Top             =   225
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton optComi 
            Caption         =   "Otros"
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
            Index           =   2
            Left            =   165
            TabIndex        =   7
            Top             =   855
            Width           =   870
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
            Left            =   1815
            TabIndex        =   6
            Top             =   510
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
            Left            =   165
            TabIndex        =   5
            Top             =   525
            Width           =   1080
         End
         Begin VB.ListBox lstComi 
            Height          =   255
            Left            =   3105
            TabIndex        =   4
            Top             =   870
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txtComi 
            Height          =   285
            Left            =   2595
            TabIndex        =   3
            Top             =   855
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.ComboBox cboComi 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1095
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   825
            Width           =   2490
         End
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1275
         TabIndex        =   29
         Top             =   405
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFHasta 
         Height          =   285
         Left            =   1275
         TabIndex        =   30
         Top             =   870
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   327680
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
         Index           =   1
         Left            =   75
         TabIndex        =   32
         Top             =   870
         Width           =   1140
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
         Left            =   75
         TabIndex        =   31
         Top             =   405
         Width           =   1140
      End
   End
   Begin TabDlg.SSTab tabTotal 
      Height          =   3330
      Left            =   90
      TabIndex        =   10
      Top             =   1575
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   5874
      _Version        =   327680
      TabOrientation  =   1
      Tabs            =   5
      TabsPerRow      =   5
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
      TabCaption(0)   =   "Importes"
      TabPicture(0)   =   "frmTot.frx":31AE
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frTotPorc0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frTot0"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "Tickets"
      TabPicture(1)   =   "frmTot.frx":31CA
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frTotPorc1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frTot1"
      Tab(1).Control(1).Enabled=   0   'False
      TabCaption(2)   =   "Pasajeros"
      TabPicture(2)   =   "frmTot.frx":31E6
      Tab(2).ControlCount=   2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frTotPorc2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frTot2"
      Tab(2).Control(1).Enabled=   0   'False
      TabCaption(3)   =   "Prom Imp x Tick"
      TabPicture(3)   =   "frmTot.frx":3202
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "sprTot(3)"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "Prom Imp x Pax"
      TabPicture(4)   =   "frmTot.frx":321E
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "sprTot(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Begin FPSpread.vaSpread sprTot 
         Height          =   2700
         Index           =   3
         Left            =   -74775
         OleObjectBlob   =   "frmTot.frx":323A
         TabIndex        =   11
         Top             =   195
         Width           =   6600
      End
      Begin FPSpread.vaSpread sprTot 
         Height          =   2700
         Index           =   4
         Left            =   -74865
         OleObjectBlob   =   "frmTot.frx":3781
         TabIndex        =   12
         Top             =   225
         Width           =   6600
      End
      Begin VB.Frame frTot0 
         Height          =   3000
         Left            =   105
         TabIndex        =   15
         Top             =   60
         Width           =   6930
         Begin FPSpread.vaSpread sprTot 
            Height          =   2700
            Index           =   0
            Left            =   60
            OleObjectBlob   =   "frmTot.frx":3CC8
            TabIndex        =   16
            Top             =   195
            Width           =   6780
         End
      End
      Begin VB.Frame frTotPorc0 
         Height          =   2985
         Left            =   90
         TabIndex        =   17
         Top             =   45
         Width           =   6945
         Begin FPSpread.vaSpread sprTotPorc 
            Height          =   2700
            Index           =   0
            Left            =   45
            OleObjectBlob   =   "frmTot.frx":423B
            TabIndex        =   18
            Top             =   195
            Width           =   6840
         End
      End
      Begin VB.Frame frTot1 
         Height          =   3000
         Left            =   -74925
         TabIndex        =   19
         Top             =   45
         Width           =   6930
         Begin FPSpread.vaSpread sprTot 
            Height          =   2700
            Index           =   1
            Left            =   105
            OleObjectBlob   =   "frmTot.frx":4647
            TabIndex        =   20
            Top             =   180
            Width           =   6735
         End
      End
      Begin VB.Frame frTotPorc1 
         Height          =   2970
         Left            =   -74925
         TabIndex        =   21
         Top             =   45
         Width           =   6930
         Begin FPSpread.vaSpread sprTotPorc 
            Height          =   2700
            Index           =   1
            Left            =   90
            OleObjectBlob   =   "frmTot.frx":4BE6
            TabIndex        =   22
            Top             =   195
            Width           =   6750
         End
      End
      Begin VB.Frame frTot2 
         Height          =   3000
         Left            =   -74910
         TabIndex        =   13
         Top             =   45
         Width           =   6800
         Begin FPSpread.vaSpread sprTot 
            Height          =   2700
            Index           =   2
            Left            =   90
            OleObjectBlob   =   "frmTot.frx":5001
            TabIndex        =   14
            Top             =   195
            Width           =   6600
         End
      End
      Begin VB.Frame frTotPorc2 
         Height          =   3000
         Left            =   -74880
         TabIndex        =   23
         Top             =   45
         Width           =   6800
         Begin FPSpread.vaSpread sprTotPorc 
            Height          =   2700
            Index           =   2
            Left            =   75
            OleObjectBlob   =   "frmTot.frx":5574
            TabIndex        =   24
            Top             =   210
            Width           =   6600
         End
      End
   End
End
Attribute VB_Name = "frmTot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsData  As Recordset

Dim RsDataPax  As Recordset

Private Sub L_RefrescarTot()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset

On Error GoTo ErrTot:
frmTot.caption = Aplicacion.SeteoProceso(frmTot.caption)

If optComi(3).Value Then
    sql = " SELECT "
    sql = sql & " fch_vta, "
    sql = sql & " cod_depn,"
    sql = sql & " decode(cod_depn,'INT','INT',cod_sdep) cod_sdep, "
''    sql = sql & " cod_sdep,"
    sql = sql & " decode(Sum (cant_tickets),null,0,Sum (cant_tickets)) cant_t, "
    sql = sql & " decode (sum(importe),null,0,sum(importe)) imp, "
    sql = sql & " decode(sum(cant_pax),null,0,sum(cant_pax)) cant_p "
    
    If Year(Date) = Year(CDate(mskFDesde.FormattedText)) Then
        sql = sql & "FROM ESTADIS.pax_espigon "
    Else
        sql = sql & "FROM ESTADIS.pax_espigon_hist "
    End If
    
    sql = sql & L_Armarcondicion
    sql = sql & " group by fch_vta,cod_depn,decode(cod_depn,'INT','INT',cod_sdep)"
    sql = sql & " order by fch_vta,cod_depn,decode(cod_depn,'INT','INT',cod_sdep)"
Else
    sql = " SELECT "
    sql = sql & " fch_ticket fch_vta, "
    sql = sql & " cod_depn,"
    sql = sql & " decode(cod_depn,'INT','INT',cod_sdep) cod_sdep, "
''    sql = sql & " cod_sdep,"
    sql = sql & " decode(Sum (cantidad),null,0,Sum (cantidad)) cant_t, "
    sql = sql & " decode (sum(importe),null,0,sum(importe)) imp "
    
    If Year(Date) = Year(CDate(mskFDesde.FormattedText)) Then
        sql = sql & "FROM ESTADIS.venta_lgi "
    Else
        sql = sql & "FROM ESTADIS.venta_lgi_hist "
    End If
    
    sql = sql & L_Armarcondicion
    sql = sql & " group by fch_ticket,cod_depn,decode(cod_depn,'INT','INT',cod_sdep)"
    sql = sql & " order by fch_ticket,cod_depn,decode(cod_depn,'INT','INT',cod_sdep)"

End If
If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        L_LlenarGrillaTOT
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabTotal.Enabled = True
    End If
    L_PintarGrillas
    Aplicacion.CerrarDAO RsData
    
End If

ErrTot:
    frmTot.caption = Aplicacion.SeteoFin
    Exit Sub
End Sub
Private Function L_Armarcondicion() As String
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String

fechaDesde = mskFDesde.FormattedText
fechaHasta = mskFHasta.FormattedText

If Not optComi(3).Value Then
    Cond = " WHERE fch_ticket between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)
    If optComi(1).Value Then
        Cond = Cond & " and  Comitente IN ('IOSC','IBAIR') "
    ElseIf optComi(0).Value Then
        Cond = Cond & " and  Comitente NOT IN ('IOSC','IBAIR','T') "
    ElseIf optComi(2).Value Then
        Cond = Cond & " and  Comitente = '" & lstComi.List(cboComi.ListIndex) & "'"
    Else
        Cond = Cond & " and  Comitente = 'T' "
    End If

Else
    Cond = " WHERE fch_vta between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)
End If
L_Armarcondicion = Cond

End Function





Private Sub L_LlenarGrillaTOT()
Dim fecha As String
Dim i As Integer
Dim TOTImp As Long
Dim TOTCant_p As Long
Dim TOTCant_t As Long

On Error GoTo ErrTot:

Do While Not RsData.EOF
    fecha = Format$(RsData!fch_vta, FTOFECHA)
    
    If optComi(3).Value Then
        For i = 0 To 4
            sprTot(i).MaxRows = sprTot(i).MaxRows + 1
            sprTot(i).SetText 1, sprTot(i).MaxRows, Format$(fecha, "dd-mm-yy")
        Next
    Else
            sprTot(0).MaxRows = sprTot(0).MaxRows + 1
            sprTot(0).SetText 1, sprTot(0).MaxRows, Format$(fecha, "dd-mm-yy")
    End If
    TOTCant_p = 0
    TOTImp = 0
    TOTCant_t = 0
    Do While fecha = Format$(RsData!fch_vta, FTOFECHA)
        Select Case RsData!cod_sdep
            Case "INTA"
                sprTot(0).SetText 2, sprTot(0).MaxRows, str(RsData!imp)
                If optComi(3).Value Then
                    sprTot(1).SetText 2, sprTot(1).MaxRows, str(RsData!cant_t)
                    sprTot(3).SetText 2, sprTot(3).MaxRows, str(RsData!imp / RsData!cant_t)
                    sprTot(2).SetText 2, sprTot(1).MaxRows, str(RsData!cant_p)
                    sprTot(4).SetText 2, sprTot(4).MaxRows, str(RsData!imp / RsData!cant_p)
                End If
                                
            Case "INTB"
                sprTot(0).SetText 3, sprTot(0).MaxRows, str(RsData!imp)
                If optComi(3).Value Then
                    sprTot(1).SetText 3, sprTot(1).MaxRows, str(RsData!cant_t)
                    sprTot(3).SetText 3, sprTot(3).MaxRows, str(RsData!imp / RsData!cant_t)
                    sprTot(2).SetText 3, sprTot(1).MaxRows, str(RsData!cant_p)
                    sprTot(4).SetText 3, sprTot(4).MaxRows, str(RsData!imp / RsData!cant_p)
                End If
            Case "AEP"
                sprTot(0).SetText 4, sprTot(0).MaxRows, str(RsData!imp)
                If optComi(3).Value Then
                    sprTot(1).SetText 4, sprTot(1).MaxRows, str(RsData!cant_t)
                    sprTot(3).SetText 4, sprTot(3).MaxRows, str(RsData!imp / RsData!cant_t)
                    sprTot(2).SetText 4, sprTot(1).MaxRows, str(RsData!cant_p)
                    sprTot(4).SetText 4, sprTot(4).MaxRows, str(RsData!imp / RsData!cant_p)
                End If
            Case "INT"
                sprTot(0).SetText 5, sprTot(0).MaxRows, str(RsData!imp)
                If optComi(3).Value Then
                    sprTot(1).SetText 5, sprTot(1).MaxRows, str(RsData!cant_t)
                    sprTot(3).SetText 5, sprTot(3).MaxRows, str(RsData!imp / RsData!cant_t)
                    sprTot(2).SetText 5, sprTot(1).MaxRows, str(RsData!cant_p)
                    sprTot(4).SetText 5, sprTot(4).MaxRows, str(RsData!imp / RsData!cant_p)
                End If
                
        End Select
        If optComi(3).Value Then
            TOTCant_p = RsData!cant_p + TOTCant_p
            TOTCant_t = RsData!cant_t + TOTCant_t
        End If
        TOTImp = RsData!imp + TOTImp
        RsData.MoveNext
        If RsData.EOF Then
            Exit Do
        End If
    Loop
    If optComi(3).Value Then
        sprTot(3).SetText sprTot(3).MaxCols, sprTot(3).MaxRows, Format$(TOTImp / TOTCant_t, "#.0")
        sprTot(4).SetText sprTot(4).MaxCols, sprTot(4).MaxRows, Format$(TOTImp / TOTCant_p, "#.0")
    End If
Loop
For i = 0 To 2
    Spread_TotalesGrillas sprTot(i), sprTot(i).MaxCols - 1, 2
Next
    FuncLocal_PromediosFilaSPR sprTot(0), sprTot(1), sprTot(3), sprTot(0).MaxCols
    FuncLocal_PromediosFilaSPR sprTot(0), sprTot(2), sprTot(4), sprTot(0).MaxCols
    
    For i = 0 To 4
        Spread_PintarfinSemana sprTot(i)
    Next
    
ErrTot:
    Exit Sub
End Sub

Private Sub botEjecutar_Click(Index As Integer)

Select Case Index
    Case 0
        If optComi(3).Value Then
            L_SeteoTabPax True
        Else
            L_SeteoTabPax False
        End If
        L_RefrescarTot
        L_RefrescarFrames
    Case 1
        'Call botPorc_Click(0)
        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        tabTotal.Enabled = False
        L_LimpiarGrillas
        L_RefrescarFrames
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

Private Sub L_RefrescarFrames()
    frTotPorc0.Refresh
    frTotPorc1.Refresh
    frTotPorc2.Refresh
End Sub


Private Sub L_LimpiarGrillas()
Dim i
    
For i = 0 To 4
    sprTot(i).MaxRows = 0
Next

End Sub

Private Sub botExcelAep_Click()
    L_TratarExcel "TOTAL COMPAÑIA ", "Informe por Local (del " & mskFDesde.FormattedText & " al " & mskFHasta.FormattedText & ")"
End Sub

Private Sub L_TratarExcel(Titulo As String, subTit As String)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim fila As Integer, filaant As Integer
Dim i
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmTot.caption = Aplicacion.SeteoProceso(frmTot.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
'    AppExcel.application.Visible = True
    
    ReDim titCol(sprTot(0).MaxCols)
    col = 1
    fila = 3
    
    Exl_PonerValor AppExcel, 1, 1, Titulo
    rango = Exl_rangos(1, 1, 1, sprTot(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.application.range(rango).merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, col, subTit
    rango = Exl_rangos(fila, fila, 1, sprTot(0).MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.application.range(rango).merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    
    fila = fila + 1
    
    If optComi(1).Value Then
        Exl_PonerValor AppExcel, fila, col, "Comitente : I.O.S.C y Propios"
    ElseIf optComi(0).Value Then
        Exl_PonerValor AppExcel, fila, col, "Comitente : NO I.O.S.C "
    ElseIf optComi(3).Value Then
        Exl_PonerValor AppExcel, fila, col, "Comitente : TODOS "
    Else
        Exl_PonerValor AppExcel, fila, col, "Comitente : " & cboComi.Text
    End If
    Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
    
    fila = fila + 2
    
    For i = 1 To sprTot(0).MaxCols
        sprTot(0).GetText i, 0, tit
        titCol(i) = tit
    Next

   For i = 0 To 4
       
       fila = fila + 2
       rango = Exl_rangos(fila, fila + 2, 1, 1)

       Exl_PonerValor AppExcel, fila, col, "Información sobre :" & L_NombreDato(i + 1)
       
       fila = fila + 2
       filaant = fila
       Exl_BajarGrillaExel sprTot(i), AppExcel, fila, col, titCol
       fila = fila + sprTot(i).MaxRows
       rango = Exl_rangos(fila, fila, col, sprTot(i).MaxCols)
       Exl_ColorInt AppExcel, rango, Exl_Gris
                                       
       If i > 2 Then
           rango = Exl_rangos(filaant + 1, fila, col + 1, sprTot(i).MaxCols)
           Exl_Format AppExcel, rango
       End If
       'AppExcel.application.Rows(Trim(str(fila + 1)) & ":" & Trim(str(fila + 1))).PageBreak = True
   Next
        AppExcel.application.ActiveSheet.PageSetup.CenterHorizontally = True
        AppExcel.application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
        'AppExcel.application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
        
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If

    AppExcel.saveas NOMBRE & ".xls"
    Set AppExcel = Nothing

End If

ErrorExl:

    frmTot.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Function L_NombreDato(i As Integer) As String
Select Case i
    Case 1
        L_NombreDato = "Importes"
    Case 2
        L_NombreDato = "Tickets"
    Case 3
        L_NombreDato = "Pasajeros"
    Case 4
        L_NombreDato = "Promedios por Ticket"
    Case 5
        L_NombreDato = "Promedios por Pasajeros"
End Select
End Function



Private Sub L_PintarGrillas()
Dim i

For i = 0 To 4
    Spread_PintarfinSemana sprTot(i)
Next

End Sub

Private Sub botGrEvoTot_Click()
    Select Case tabTotal.Tab
        Case 0
            frmGrafEvoLocal.CargarGrafico total, "Importes", sprTot(0), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 1
            frmGrafEvoLocal.CargarGrafico total, "Tickets", sprTot(1), mskFDesde.FormattedText, mskFHasta.FormattedText
        Case 2
            frmGrafEvoLocal.CargarGrafico total, "Pasajeros", sprTot(2), mskFDesde.FormattedText, mskFHasta.FormattedText
    End Select

End Sub

Private Sub botGrTorta_Click()
Dim col_datos As Collection

    Set col_datos = New Collection
    
    Select Case tabTotal.Tab
        Case 0
            L_LLenarColeccion col_datos, sprTot(0)
            FrmGraficos.CargarGrafico total, "Importes", col_datos
        Case 1
            L_LLenarColeccion col_datos, sprTot(1)
            FrmGraficos.CargarGrafico total, "Tickets", col_datos
        Case 2
            L_LLenarColeccion col_datos, sprTot(2)
            FrmGraficos.CargarGrafico total, "Pasajeros", col_datos

    End Select


End Sub

Private Sub L_LLenarColeccion(ByRef col As Collection, spr As Control)
Dim cl_dato As CLlgi
Dim valor As Variant
Dim fecha As Variant
Dim i

spr.Row = spr.ActiveRow

spr.GetText 1, spr.Row, fecha


For i = 2 To spr.MaxCols - 1
    Set cl_dato = New CLlgi
        
    cl_dato.fch = fecha
    
    spr.GetText i, 0, valor
    cl_dato.Locale = valor
    
    spr.GetText i, spr.Row, valor
    cl_dato.DatoGral = valor
    
    col.Add cl_dato
    
Next

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


Private Sub botPorcTot_Click(Index As Integer)
Dim i

Select Case Index
    Case 0
        frTot0.Visible = True
        frTot1.Visible = True
        frTot2.Visible = True
        botPorcTot(0).Enabled = False
        botEjecutar(1).Enabled = True
        frTot0.Refresh
        frTot1.Refresh
        frTot2.Refresh
    Case 1
        frTot0.Visible = False
        frTot1.Visible = False
        frTot2.Visible = False
        FuncLocal_Porcentages sprTotPorc(0), sprTot(0), "FILAS"
        FuncLocal_Porcentages sprTotPorc(1), sprTot(1), "FILAS"
        FuncLocal_Porcentages sprTotPorc(2), sprTot(2), "FILAS"
        botPorcTot(0).Enabled = True
        frTotPorc0.Refresh
        frTotPorc1.Refresh
        frTotPorc2.Refresh
        botEjecutar(1).Enabled = False
    Case 2
        frTot0.Visible = False
        frTot1.Visible = False
        frTot2.Visible = False
        FuncLocal_Porcentages sprTotPorc(0), sprTot(0), "COL"
        FuncLocal_Porcentages sprTotPorc(1), sprTot(1), "COL"
        FuncLocal_Porcentages sprTotPorc(2), sprTot(2), "COL"
        
        botPorcTot(0).Enabled = True
        frTotPorc0.Refresh
        frTotPorc1.Refresh
        frTotPorc2.Refresh
        botEjecutar(1).Enabled = False
        
End Select


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


FuncCbo_LlenarCbosLST cboComi, lstComi, "COMITENTE"

L_RefrescarFrames
L_LimpiarGrillas
frmPrincipal.lstForms.AddItem "frmTot"

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
        If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Or Year(mskFHasta.FormattedText) <> Year(mskFDesde.FormattedText) Then
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


Private Sub optComi_Click(Index As Integer)

Select Case Index
    Case 0, 1, 3
        cboComi.ListIndex = -1
        cboComi.Enabled = False
    Case 2
        cboComi.ListIndex = 0
        cboComi.Enabled = True
    
End Select

End Sub

Private Sub tabTotal_DblClick()
frTot0.Refresh
frTot1.Refresh
frTot2.Refresh
Select Case tabTotal.Tab
    Case 0, 1, 2
        botPorcTot(0).Visible = True
        botPorcTot(1).Visible = True
        botPorcTot(2).Visible = True
        botGrTorta.Visible = True
        botGrEvoTot.Visible = True
    Case 3, 4
        botPorcTot(0).Visible = False
        botPorcTot(1).Visible = False
        botPorcTot(2).Visible = False
        botGrTorta.Visible = False
        botGrEvoTot.Visible = False
End Select


End Sub


