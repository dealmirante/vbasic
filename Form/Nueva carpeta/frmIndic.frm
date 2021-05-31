VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIndic 
   Caption         =   "INDICADORES"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9195
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   8955
      Top             =   1965
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   3810
      Left            =   75
      TabIndex        =   2
      Top             =   1545
      Width           =   8950
      _ExtentX        =   15796
      _ExtentY        =   6720
      _Version        =   327680
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   459
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "TOT BS.AS."
      TabPicture(0)   =   "frmIndic.frx":0000
      Tab(0).ControlCount=   3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "borGrBar(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "botExcel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "sprTotalBsAs"
      Tab(0).Control(2).Enabled=   0   'False
      TabCaption(1)   =   "INTA-L"
      TabPicture(1)   =   "frmIndic.frx":001C
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sprEzeA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "borGrBar(1)"
      Tab(1).Control(1).Enabled=   0   'False
      TabCaption(2)   =   "INTB-S"
      TabPicture(2)   =   "frmIndic.frx":0038
      Tab(2).ControlCount=   2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "sprEzeAS"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "borGrBar(6)"
      Tab(2).Control(1).Enabled=   0   'False
      TabCaption(3)   =   "INTER B"
      TabPicture(3)   =   "frmIndic.frx":0054
      Tab(3).ControlCount=   2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "sprEzeB"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "borGrBar(2)"
      Tab(3).Control(1).Enabled=   0   'False
      TabCaption(4)   =   "AEP"
      TabPicture(4)   =   "frmIndic.frx":0070
      Tab(4).ControlCount=   2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "sprAep"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "borGrBar(3)"
      Tab(4).Control(1).Enabled=   0   'False
      TabCaption(5)   =   "TOT CIA"
      TabPicture(5)   =   "frmIndic.frx":008C
      Tab(5).ControlCount=   3
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "sprTotal"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "borGrBar(5)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Command1"
      Tab(5).Control(2).Enabled=   0   'False
      TabCaption(6)   =   "INTERIOR"
      TabPicture(6)   =   "frmIndic.frx":00A8
      Tab(6).ControlCount=   2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "borGrBar(4)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "SprInt"
      Tab(6).Control(1).Enabled=   0   'False
      Begin FPSpread.vaSpread sprEzeB 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic.frx":00C4
         TabIndex        =   36
         Top             =   450
         Width           =   7695
      End
      Begin FPSpread.vaSpread sprAep 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic.frx":07F9
         TabIndex        =   34
         Top             =   450
         Width           =   7695
      End
      Begin FPSpread.vaSpread sprTotal 
         Height          =   3105
         Left            =   -74775
         OleObjectBlob   =   "frmIndic.frx":0F40
         TabIndex        =   31
         Top             =   450
         Width           =   7695
      End
      Begin FPSpread.vaSpread SprInt 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic.frx":1675
         TabIndex        =   29
         Top             =   450
         Width           =   7695
      End
      Begin FPSpread.vaSpread sprTotalBsAs 
         Height          =   3105
         Left            =   195
         OleObjectBlob   =   "frmIndic.frx":1DBC
         TabIndex        =   26
         Top             =   450
         Width           =   7695
      End
      Begin FPSpread.vaSpread sprEzeA 
         Height          =   3100
         Left            =   -74800
         OleObjectBlob   =   "frmIndic.frx":24F1
         TabIndex        =   14
         Top             =   450
         Width           =   7700
      End
      Begin FPSpread.vaSpread sprEzeAS 
         Height          =   3105
         Left            =   -74805
         OleObjectBlob   =   "frmIndic.frx":2C26
         TabIndex        =   38
         Top             =   450
         Width           =   7695
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   6
         Left            =   -66975
         Picture         =   "frmIndic.frx":335B
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3030
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   2
         Left            =   -66990
         Picture         =   "frmIndic.frx":379D
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3030
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   3
         Left            =   -66990
         Picture         =   "frmIndic.frx":3BDF
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3030
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   5
         Left            =   -66960
         Picture         =   "frmIndic.frx":4021
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2370
         Width           =   780
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -66960
         Picture         =   "frmIndic.frx":4463
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3060
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   4
         Left            =   -66990
         Picture         =   "frmIndic.frx":49F5
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3045
         Width           =   780
      End
      Begin VB.CommandButton botExcel 
         Caption         =   "Excel"
         Height          =   510
         Left            =   8040
         Picture         =   "frmIndic.frx":4E37
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3045
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   0
         Left            =   8040
         Picture         =   "frmIndic.frx":53C9
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2385
         Width           =   780
      End
      Begin VB.CommandButton borGrBar 
         Height          =   510
         Index           =   1
         Left            =   -66975
         Picture         =   "frmIndic.frx":580B
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3030
         Width           =   780
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1500
      Left            =   8220
      TabIndex        =   1
      Top             =   -30
      Width           =   825
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
         Left            =   150
         Picture         =   "frmIndic.frx":5C4D
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   165
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
         Left            =   135
         Picture         =   "frmIndic.frx":5D4F
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   570
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
         Left            =   135
         Picture         =   "frmIndic.frx":5E51
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   990
         Width           =   570
      End
   End
   Begin VB.Frame frdatos 
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   8175
      Begin VB.Frame Frame1 
         Caption         =   "Fechas de  Consultas para Comparación"
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
         Height          =   1485
         Left            =   105
         TabIndex        =   3
         Top             =   30
         Width           =   8040
         Begin VB.CheckBox chkNeto 
            Caption         =   "Venta Neta"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   40
            Top             =   1215
            Value           =   1  'Checked
            Width           =   1410
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
            Left            =   4980
            TabIndex        =   18
            Top             =   210
            Width           =   3015
            Begin VB.CommandButton botHelpFDAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmIndic.frx":6673
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton botHelpFHAnt 
               Height          =   345
               Left            =   2520
               Picture         =   "frmIndic.frx":67E5
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   600
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesdeAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   21
               Top             =   255
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               _Version        =   327680
               PromptInclude   =   0   'False
               MaxLength       =   10
               Mask            =   "##-##-####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFHastaAnt 
               Height          =   285
               Left            =   1380
               TabIndex        =   22
               Top             =   600
               Width           =   1125
               _ExtentX        =   1984
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
               Left            =   150
               TabIndex        =   24
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
               Index           =   3
               Left            =   135
               TabIndex        =   23
               Top             =   600
               Width           =   1185
            End
         End
         Begin VB.OptionButton optFechas 
            Caption         =   "Acumulado Anual "
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   375
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.OptionButton optFechas 
            Caption         =   "Acumulado Rango de Fechas"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   90
            TabIndex        =   15
            Top             =   525
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VB.Frame Frame4 
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
            Height          =   1095
            Left            =   1935
            TabIndex        =   4
            Top             =   225
            Width           =   3015
            Begin VB.CommandButton botHelpFH 
               Height          =   345
               Left            =   2535
               Picture         =   "frmIndic.frx":6957
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   615
               Width           =   375
            End
            Begin VB.CommandButton botHelpFD 
               Height          =   345
               Left            =   2550
               Picture         =   "frmIndic.frx":6AC9
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   255
               Width           =   375
            End
            Begin MSMask.MaskEdBox mskFDesde 
               Height          =   285
               Left            =   1380
               TabIndex        =   7
               Top             =   270
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
               TabIndex        =   8
               Top             =   615
               Width           =   1125
               _ExtentX        =   1984
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
               Left            =   135
               TabIndex        =   10
               Top             =   615
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
               Index           =   0
               Left            =   150
               TabIndex        =   9
               Top             =   270
               Width           =   1185
            End
         End
      End
   End
   Begin VB.Label labInfo 
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
      Height          =   420
      Left            =   150
      TabIndex        =   25
      Top             =   5400
      Width           =   8730
   End
End
Attribute VB_Name = "frmIndic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset
Dim RsDataAnt  As Recordset
Dim RsDataEstim  As Recordset
Dim RsDataVol  As Recordset

Dim UsoPorIntA As Boolean 'cuenta la cantidad de consultas
                       'con porcentajes
Dim UsoPorIntB As Boolean
Dim UsoPorAero As Boolean
Dim UsoPorTotal As Boolean
Dim UsoPorInte As Boolean

Dim Control_BsAs As Boolean
Dim Control_Int As Boolean

Dim Fch_BsAs As String
Dim Fch_Int As String

Private Sub L_ControlFecha()
Dim FH As String
'Dim Fch_BsAs As String
'Dim Fch_Int As String
Dim rs_pax As Recordset
Dim SQL As String

FH = mskFHasta.FormattedText

SQL = "select min(fch_actualizado) Fch from estadis.control_carga Where Cod_sdep <> 'INT' "

If Aplicacion.ObtenerRsDAO(SQL, rs_pax) Then
   Fch_BsAs = CDate(rs_pax!fch)
End If


SQL = "select min(fch_actualizado) Fch from estadis.control_carga Where Cod_sdep = 'INT' "

If Aplicacion.ObtenerRsDAO(SQL, rs_pax) Then
   Fch_Int = CDate(rs_pax!fch)
End If

If CDate(FH) <= CDate(Fch_BsAs) Then
   Control_BsAs = True
 Else
   Control_BsAs = False
End If

If CDate(FH) <= CDate(Fch_Int) Then
   Control_Int = True
 Else
   Control_Int = False
End If


End Sub



Private Function L_Armarcondicion(anio As Integer) As String
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant

If optFechas(0).Value Then
    If anio = Year(Date) Then
        fechaDesde = "01-01-" & Trim(str(anio))
        fechaHasta = Format$(Date - 1, "dd-mm") & "-" & Trim(str(anio))
    Else
        fechaDesde = "01-01-" & Trim(str(anio))
        fechaHasta = Format$(Date - 1, "dd-mm") & "-" & Trim(str(anio))
    End If
Else
    If anio = Year(mskFDesde.FormattedText) Then
        fechaDesde = mskFDesde.FormattedText
        fechaHasta = mskFHasta.FormattedText
    Else
        fechaDesde = mskFDesdeAnt.FormattedText
        fechaHasta = mskFHastaAnt.FormattedText
    End If
End If

Cond = " WHERE fch_vta between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)
Cond = Cond & " AND A.COD_LOCAL = S.COD_LOCAL and  A.COD_depn = S.COD_depn and  A.COD_sdep = S.COD_sdep "
L_Armarcondicion = Cond

End Function

Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant

Do While Not RsData.EOF
        Select Case RsData!cod_depn
            Case DSLoc(1).Dep

                SprAep.SetText 2, 1, str(RsData!imp)
                
                SprAep.SetText 2, 2, str(RsData!cant_t)
                
                SprAep.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                                   
            
            Case DSLoc(2).Dep
                Select Case RsData!cod_sdep
                    Case DSLoc(2).Sdep

                        SprEzeA.SetText 2, 1, str(RsData!imp)
                        
                        SprEzeA.SetText 2, 2, str(RsData!cant_t)
                        
                        SprEzeA.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                        

                    Case DSLoc(3).Sdep
                        SprEzeB.SetText 2, 1, str(RsData!imp)
                        
                        SprEzeB.SetText 2, 2, str(RsData!cant_t)
                        
                        SprEzeB.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                    
                    Case DSLoc(9).Sdep

                        SprEzeA.SetText 2, 1, str(RsData!imp)
                        
                        SprEzeA.SetText 2, 2, str(RsData!cant_t)
                        
                        SprEzeA.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                        

                    Case DSLoc(10).Sdep
                        sprEzeAS.SetText 2, 1, str(RsData!imp)
                        
                        sprEzeAS.SetText 2, 2, str(RsData!cant_t)
                        
                        sprEzeAS.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                        
                
                End Select
                
            Case DSLoc(4).Dep

                        SprInt.SetText 2, 1, str(RsData!imp)
                        
                        SprInt.SetText 2, 2, str(RsData!cant_t)
                        
                        SprInt.SetText 2, 3, str(RsData!imp / RsData!cant_t)
                        
                

        End Select
                
        RsData.MoveNext

Loop


End Sub
Private Sub L_DecoEspigonAnt()
Dim fecha As String
Dim i As Integer
Dim dato As Variant

Do While Not RsDataAnt.EOF
        Select Case RsDataAnt!cod_depn
            Case DSLoc(1).Dep
                SprAep.SetText 3, 1, str(RsDataAnt!imp)
                SprAep.SetText 3, 2, str(RsDataAnt!cant_t)
                SprAep.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
            Case DSLoc(2).Dep
                Select Case RsDataAnt!cod_sdep
                    Case DSLoc(2).Sdep
                        SprEzeA.SetText 3, 1, str(RsDataAnt!imp)
                        SprEzeA.SetText 3, 2, str(RsDataAnt!cant_t)
                        SprEzeA.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                    Case DSLoc(3).Sdep
                        SprEzeB.SetText 3, 1, str(RsDataAnt!imp)
                        SprEzeB.SetText 3, 2, str(RsDataAnt!cant_t)
                        SprEzeB.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                    Case DSLoc(9).Sdep
                        SprEzeA.SetText 3, 1, str(RsDataAnt!imp)
                        SprEzeA.SetText 3, 2, str(RsDataAnt!cant_t)
                        SprEzeA.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                    Case DSLoc(10).Sdep
                        sprEzeAS.SetText 3, 1, str(RsDataAnt!imp)
                        sprEzeAS.SetText 3, 2, str(RsDataAnt!cant_t)
                        sprEzeAS.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
                End Select
            Case DSLoc(4).Dep
                SprInt.SetText 3, 1, str(RsDataAnt!imp)
                SprInt.SetText 3, 2, str(RsDataAnt!cant_t)
                SprInt.SetText 3, 3, str(RsDataAnt!imp / RsDataAnt!cant_t)
             
        End Select
                
        RsDataAnt.MoveNext

Loop

End Sub
Private Sub L_DecoEstim()
Dim ImpAep As Double, ImpA As Double, ImpB As Double
Dim valor As Variant

Do While Not RsDataEstim.EOF
        Select Case RsDataEstim!cod_depn
            Case DSLoc(1).Dep
                Select Case RsDataEstim!tipo_porc
                    Case "I"
                        SprAep.SetText 6, 1, str(RsDataEstim!valor)
                        ImpAep = RsDataEstim!valor
                    Case "T"
                        SprAep.SetText 6, 2, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            SprAep.SetText 6, 3, str(ImpAep / RsDataEstim!valor)
                        End If
                        SprAep.GetText 6, 4, valor
                        If Val(valor) > 0 Then
                            SprAep.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                        End If
                    
                    Case "P"
                        SprAep.SetText 6, 4, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            SprAep.SetText 6, 5, str(ImpAep / RsDataEstim!valor)
                        End If
                
                End Select

            Case DSLoc(2).Dep
                Select Case RsDataEstim!cod_sdep
                    Case DSLoc(2).Sdep
                        Select Case RsDataEstim!tipo_porc
                            Case "I"
                                SprEzeA.SetText 6, 1, str(RsDataEstim!valor)
                                ImpA = RsDataEstim!valor
                            Case "T"
                                SprEzeA.SetText 6, 2, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    SprEzeA.SetText 6, 3, str(ImpA / RsDataEstim!valor)
                                End If
                                SprEzeA.GetText 6, 4, valor
                                If Val(valor) > 0 Then
                                    SprEzeA.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                                End If
                            
                            Case "P"
                                SprEzeA.SetText 6, 4, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    SprEzeA.SetText 6, 5, str(ImpA / RsDataEstim!valor)
                                End If
                            
                        End Select
                        
                    Case DSLoc(3).Sdep
                            Select Case RsDataEstim!tipo_porc
                                Case "I"
                                    SprEzeB.SetText 6, 1, str(RsDataEstim!valor)
                                    ImpB = RsDataEstim!valor
                                Case "T"
                                    SprEzeB.SetText 6, 2, str(RsDataEstim!valor)
                                    If RsDataEstim!valor > 0 Then
                                        SprEzeB.SetText 6, 3, str(ImpB / RsDataEstim!valor)
                                    End If
                                    SprEzeB.GetText 6, 4, valor
                                    If Val(valor) > 0 Then
                                        SprEzeB.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                                    End If
                                    
                                Case "P"
                                    SprEzeB.SetText 6, 4, str(RsDataEstim!valor)
                                    If RsDataEstim!valor > 0 Then
                                        SprEzeB.SetText 6, 5, str(ImpB / RsDataEstim!valor)
                                    End If
                            End Select
                    Case DSLoc(9).Sdep
                        Select Case RsDataEstim!tipo_porc
                            Case "I"
                                SprEzeA.SetText 6, 1, str(RsDataEstim!valor)
                                ImpA = RsDataEstim!valor
                            Case "T"
                                SprEzeA.SetText 6, 2, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    SprEzeA.SetText 6, 3, str(ImpA / RsDataEstim!valor)
                                End If
                                SprEzeA.GetText 6, 4, valor
                                If Val(valor) > 0 Then
                                    SprEzeA.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                                End If
                            
                            Case "P"
                                SprEzeA.SetText 6, 4, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    SprEzeA.SetText 6, 5, str(ImpA / RsDataEstim!valor)
                                End If
                            
                        End Select
                    Case DSLoc(10).Sdep
                        Select Case RsDataEstim!tipo_porc
                            Case "I"
                                sprEzeAS.SetText 6, 1, str(RsDataEstim!valor)
                                ImpA = RsDataEstim!valor
                            Case "T"
                                sprEzeAS.SetText 6, 2, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    sprEzeAS.SetText 6, 3, str(ImpA / RsDataEstim!valor)
                                End If
                                sprEzeAS.GetText 6, 4, valor
                                If Val(valor) > 0 Then
                                    sprEzeAS.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                                End If
                            
                            Case "P"
                                sprEzeAS.SetText 6, 4, str(RsDataEstim!valor)
                                If RsDataEstim!valor > 0 Then
                                    sprEzeAS.SetText 6, 5, str(ImpA / RsDataEstim!valor)
                                End If
                            
                        End Select
                    
                    End Select
            Case DSLoc(4).Dep
                Select Case RsDataEstim!tipo_porc
                    Case "I"
                        SprInt.SetText 6, 1, str(RsDataEstim!valor)
                        ImpAep = RsDataEstim!valor
                    Case "T"
                        SprInt.SetText 6, 2, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            SprInt.SetText 6, 3, str(ImpAep / RsDataEstim!valor)
                        End If
                        SprInt.GetText 6, 4, valor
                        If Val(valor) > 0 Then
                            SprInt.SetText 6, 6, str(RsDataEstim!valor / valor * 100)
                        End If
                    
                    Case "P"
                        SprInt.SetText 6, 4, str(RsDataEstim!valor)
                        If RsDataEstim!valor > 0 Then
                            SprInt.SetText 6, 5, str(ImpAep / RsDataEstim!valor)
                        End If
                                                            
                End Select
        End Select
                
        RsDataEstim.MoveNext

Loop
End Sub


Private Sub L_DecoVolados(col As Integer)
Dim valor As Variant

Do While Not RsDataVol.EOF
        Select Case RsDataVol!cod_depn
            Case DSLoc(1).Dep
                   SprAep.SetText col, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprAep.GetText col, 1, valor
                      SprAep.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                      SprAep.GetText col, 2, valor
                      SprAep.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
            Case DSLoc(2).Dep
                Select Case RsDataVol!cod_sdep
                    Case DSLoc(2).Sdep
                        SprEzeA.SetText col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprEzeA.GetText col, 1, valor
                            SprEzeA.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            SprEzeA.GetText col, 2, valor
                            SprEzeA.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(3).Sdep
                        SprEzeB.SetText col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprEzeB.GetText col, 1, valor
                            SprEzeB.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            SprEzeB.GetText col, 2, valor
                            SprEzeB.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(9).Sdep
                        SprEzeA.SetText col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprEzeA.GetText col, 1, valor
                            SprEzeA.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            SprEzeA.GetText col, 2, valor
                            SprEzeA.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(10).Sdep
                        sprEzeAS.SetText col, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            sprEzeAS.GetText col, 1, valor
                            sprEzeAS.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                            sprEzeAS.GetText col, 2, valor
                            sprEzeAS.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                
                End Select
        
            Case DSLoc(4).Dep
               If (Control_Int And col = 2) Or col = 3 Then
                   SprInt.SetText col, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprInt.GetText col, 1, valor
                      SprInt.SetText col, 5, Format$((valor / RsDataVol!cant_v), "#.00")
                      SprInt.GetText col, 2, valor
                      SprInt.SetText col, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
                 Else
                   labInfo.caption = "Datos de pasajeros transitados del Interior al día " & Fch_Int
                   Timer1.Enabled = True
                   Exit Do
               End If
        End Select
                
        RsDataVol.MoveNext
Loop
End Sub

Private Sub L_DecoVoladosHIST()
Dim valor As Variant

Do While Not RsDataVol.EOF
        Select Case RsDataVol!cod_depn
            Case DSLoc(1).Dep
                   SprAep.SetText 3, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprAep.GetText 3, 1, valor
                      SprAep.SetText 3, 5, str(valor / RsDataVol!cant_v)
                      SprAep.GetText 3, 2, valor
                      SprAep.SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
            Case DSLoc(2).Dep
                Select Case RsDataVol!cod_sdep
                    Case DSLoc(2).Sdep
                        SprEzeA.SetText 3, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprEzeA.GetText 3, 1, valor
                            SprEzeA.SetText 3, 5, str(valor / RsDataVol!cant_v)
                            SprEzeA.GetText 3, 2, valor
                            SprEzeA.SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                    Case DSLoc(3).Sdep
                        SprEzeB.SetText 2, 4, str(RsDataVol!cant_v)
                        If RsDataVol!cant_v > 0 Then
                            SprEzeB.GetText 2, 1, valor
                            SprEzeB.SetText 2, 5, str(valor / RsDataVol!cant_v)
                            SprEzeB.GetText 2, 2, valor
                            SprEzeB.SetText 2, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                        End If
                End Select
                
            Case DSLoc(4).Dep
                   SprInt.SetText 3, 4, str(RsDataVol!cant_v)
                   If RsDataVol!cant_v > 0 Then
                      SprInt.GetText 3, 1, valor
                      SprInt.SetText 3, 5, str(valor / RsDataVol!cant_v)
                      SprInt.GetText 3, 2, valor
                      SprInt.SetText 3, 6, Format$((valor / RsDataVol!cant_v * 100), "00.00")
                   End If
                
        End Select
                
        RsDataVol.MoveNext
Loop
End Sub


Private Sub L_LimpiarGrillas()
Dim i
    
    SprEzeA.MaxRows = 0
    sprEzeAS.MaxRows = 0
    SprEzeB.MaxRows = 0
    SprAep.MaxRows = 0
    SprInt.MaxRows = 0
    SprTotal.MaxRows = 0
    SprTotalBsAs.MaxRows = 0
    
    labInfo.caption = ""
    Timer1.Enabled = False
End Sub

Private Sub L_LlenarTotales()
Dim Nom As String
Dim i As Integer

For i = 1 To 6
    SprTotal.MaxRows = SprTotal.MaxRows + 1
    
    Select Case i
        Case 1
            Nom = "Facturación"
        Case 2
            Nom = "Tickets"
        Case 3
            Nom = "Pr. x Tickets"
        Case 4
            Nom = "Pax Viajados"
       Case 5
            Nom = "Pr x Pax Viaj."
        Case 6
            Nom = "Captación"
    End Select
    
    SprTotal.SetText 1, i, Nom
    
     If i < 4 Or Control_Int Then
        L_RellenarDatos i, 2, Nom
     End If
        L_RellenarDatos i, 3, Nom
        L_RellenarDatos i, 6, Nom

Next
End Sub
Private Sub L_LlenarTotalesBsAs()
Dim Nom As String
Dim i As Integer

For i = 1 To 6
    SprTotalBsAs.MaxRows = SprTotalBsAs.MaxRows + 1
    
    Select Case i
        Case 1
            Nom = "Facturación"
        Case 2
            Nom = "Tickets"
        Case 3
            Nom = "Pr. x Tickets"
        Case 4
            Nom = "Pax Viajados"
        Case 5
            Nom = "Pr x Pax Viaj."
        Case 6
            Nom = "Captación"
    End Select
    
    SprTotalBsAs.SetText 1, i, Nom
    
        L_RellenarDatos_BsAs i, 2, Nom
        L_RellenarDatos_BsAs i, 3, Nom
        L_RellenarDatos_BsAs i, 6, Nom

Next
End Sub

Public Sub L_PonerItems()
SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Facturación"
SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Tickets"
SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Pr. x Tickets"
SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Pax Viajados"
SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Pr x Pax Viaj."
SprAep.MaxRows = SprAep.MaxRows + 1
SprAep.SetText 1, SprAep.MaxRows, "Captación"

SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Facturación"
SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Tickets"
SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Pr. x Tickets"
SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Pax Viajados"
SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Pr x Pax Viaj."
SprEzeA.MaxRows = SprEzeA.MaxRows + 1
SprEzeA.SetText 1, SprEzeA.MaxRows, "Captación"

sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Facturación"
sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Tickets"
sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Pr. x Tickets"
sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Pax Viajados"
sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Pr x Pax Viaj."
sprEzeAS.MaxRows = sprEzeAS.MaxRows + 1
sprEzeAS.SetText 1, sprEzeAS.MaxRows, "Captación"

SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Facturación"
SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Tickets"
SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Pr. x Tickets"
SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Pax Viajados"
SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Pr x Pax Viaj."
SprEzeB.MaxRows = SprEzeB.MaxRows + 1
SprEzeB.SetText 1, SprEzeB.MaxRows, "Captación"

SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Facturación"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Tickets"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Pr. x Tickets"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Pax Viajados"
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Pr x Pax Viaj."
SprInt.MaxRows = SprInt.MaxRows + 1
SprInt.SetText 1, SprInt.MaxRows, "Captación"


End Sub

Private Sub L_Refrescar()
Dim SQL As String
Dim sqlX As String
Dim rs As Recordset

On Error GoTo ErrInd:

frmIndic.caption = Aplicacion.SeteoProceso(frmIndic.caption)


SQL = " SELECT "
SQL = SQL & " s.cod_depn, "
SQL = SQL & " decode(s.cod_depn,'INT','INT',cod_Ssdep) cod_sdep, "
SQL = SQL & " sum(cant_tickets) cant_t, "
SQL = SQL & " sum(importe) imp "
SQL = SQL & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)))
SQL = SQL & " S ,VENTAS.APERTURA_SDEP A "
SQL = SQL & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)))
SQL = SQL & "group by s.cod_depn,decode(s.cod_depn,'INT','INT',cod_Ssdep) "
SQL = SQL & " order by s.cod_depn,decode(s.cod_depn,'INT','INT',cod_Ssdep) "

sqlX = " SELECT "
sqlX = sqlX & " s.cod_depn, "
sqlX = sqlX & " decode(s.cod_depn,'INT','INT',cod_ssdep) cod_sdep, "
sqlX = sqlX & " sum(cant_tickets) cant_t, "
sqlX = sqlX & " sum(importe) imp "
sqlX = sqlX & "FROM " & funcLocal_Vista("pax_local", Year(CDate(mskFDesde.FormattedText)) - 1)
sqlX = sqlX & " S ,VENTAS.APERTURA_SDEP A "
sqlX = sqlX & L_Armarcondicion(Year(CDate(mskFDesde.FormattedText)) - 1)
sqlX = sqlX & "group by s.cod_depn,decode(s.cod_depn,'INT','INT',cod_ssdep) "
sqlX = sqlX & " order by s.cod_depn,decode(s.cod_depn,'INT','INT',cod_ssdep)"


If Aplicacion.ObtenerRsDAO(SQL, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        L_PonerItems
        L_DecoEspigon
        If Aplicacion.ObtenerRsDAO(sqlX, RsDataAnt) Then
            L_DecoEspigonAnt
        End If
        L_TratarVolados
        L_TratarVoladosHIST
        L_TratarEstimado
        L_LlenarTotales
        L_LlenarTotalesBsAs
        
        L_Resaltar
        
        Aplicacion.CerrarDAO RsDataAnt
    End If
    
    Aplicacion.CerrarDAO RsData

End If

ErrInd:
    frmIndic.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub


Private Sub L_RellenarDatos(i As Integer, col As Integer, Nom As String)
Dim valor As Variant, divsor As Variant
Dim Impr As Double
    
    If i = 1 Or i = 2 Or i = 4 Then
        SprEzeA.GetText col, i, valor
        Impr = Impr + valor
        SprEzeB.GetText col, i, valor
        Impr = Impr + valor
        SprAep.GetText col, i, valor
        Impr = Impr + valor
        SprInt.GetText col, i, valor
        Impr = Impr + valor
    ElseIf i = 3 Then
        SprTotal.GetText col, 1, valor
        SprTotal.GetText col, 2, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 5 Then
        SprTotal.GetText col, 1, valor
        SprTotal.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 6 Then
        SprTotal.GetText col, 2, valor
        SprTotal.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor * 100
        End If
    
    End If
    
    SprTotal.SetText col, i, Format$((Impr), "#.00")

End Sub
Private Sub L_RellenarDatos_BsAs(i As Integer, col As Integer, Nom As String)
Dim valor As Variant, divsor As Variant
Dim Impr As Double
    
    If i = 1 Or i = 2 Or i = 4 Then
        SprEzeA.GetText col, i, valor
        Impr = Impr + valor
        SprEzeB.GetText col, i, valor
        Impr = Impr + valor
        SprAep.GetText col, i, valor
        Impr = Impr + valor
        'SprInt.GetText col, i, valor
        'Impr = Impr + valor
    ElseIf i = 3 Then
        SprTotalBsAs.GetText col, 1, valor
        SprTotalBsAs.GetText col, 2, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 5 Then
        SprTotalBsAs.GetText col, 1, valor
        SprTotalBsAs.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor
        End If
    ElseIf i = 6 Then
        SprTotalBsAs.GetText col, 2, valor
        SprTotalBsAs.GetText col, 4, divsor
        If divsor > 0 Then
            Impr = valor / divsor * 100
        End If
    
    End If
    
    SprTotalBsAs.SetText col, i, Format$((Impr), "#.00")

End Sub

Private Sub L_Resaltar()
Dim fila As Integer

For fila = 1 To SprTotal.MaxRows
    Spread.spread_ResaltarCelda SprTotal, 4, fila
    Spread.spread_ResaltarCelda SprTotal, 5, fila
    Spread.spread_ResaltarCelda SprTotal, 7, fila
    Spread.spread_ResaltarCelda SprTotal, 8, fila
    
    Spread.spread_ResaltarCelda SprEzeA, 4, fila
    Spread.spread_ResaltarCelda SprEzeA, 5, fila
    Spread.spread_ResaltarCelda SprEzeA, 7, fila
    Spread.spread_ResaltarCelda SprEzeA, 8, fila

    Spread.spread_ResaltarCelda sprEzeAS, 4, fila
    Spread.spread_ResaltarCelda sprEzeAS, 5, fila
    Spread.spread_ResaltarCelda sprEzeAS, 7, fila
    Spread.spread_ResaltarCelda sprEzeAS, 8, fila

    Spread.spread_ResaltarCelda SprEzeB, 4, fila
    Spread.spread_ResaltarCelda SprEzeB, 5, fila
    Spread.spread_ResaltarCelda SprEzeB, 7, fila
    Spread.spread_ResaltarCelda SprEzeB, 8, fila

    Spread.spread_ResaltarCelda SprAep, 4, fila
    Spread.spread_ResaltarCelda SprAep, 5, fila
    Spread.spread_ResaltarCelda SprAep, 7, fila
    Spread.spread_ResaltarCelda SprAep, 8, fila

    Spread.spread_ResaltarCelda SprInt, 4, fila
    Spread.spread_ResaltarCelda SprInt, 5, fila
    Spread.spread_ResaltarCelda SprInt, 7, fila
    Spread.spread_ResaltarCelda SprInt, 8, fila

    Spread.spread_ResaltarCelda SprTotalBsAs, 4, fila
    Spread.spread_ResaltarCelda SprTotalBsAs, 5, fila
    Spread.spread_ResaltarCelda SprTotalBsAs, 7, fila
    Spread.spread_ResaltarCelda SprTotalBsAs, 8, fila

Next

End Sub

Private Sub L_SeteoEjecutar(valor As Boolean, dato As String)

If valor Then
    Select Case dato
        Case EZEA
            UsoPorIntA = True
        Case EZEB
            UsoPorIntB = True
        Case AERO
            UsoPorAero = True
        Case INTE
            UsoPorInte = True
        Case total
            UsoPorTotal = True
    End Select
    botEjecutar(1).Enabled = False
Else
    Select Case dato
        Case EZEA
            UsoPorIntA = False
        Case EZEB
            UsoPorIntB = False
        Case AERO
            UsoPorAero = False
        Case INTE
            UsoPorInte = True
        Case total
            UsoPorTotal = False
    End Select
    
    If Not (UsoPorIntA Or UsoPorIntA Or UsoPorAero Or UsoPorTotal Or UsoPorInte) Then
        botEjecutar(1).Enabled = True
    End If
End If

End Sub


Private Sub L_TratarEstimado()
Dim SQL As String
Dim Tipo As String
Dim anio As Integer, Mes As Integer
Dim FD As String, FH As String

Aplicacion.ComienzoTrans

If optFechas(0).Value Then
    Tipo = "A"
    anio = Year(mskFDesde.FormattedText)
    Mes = Month(mskFDesde.FormattedText)
    FD = "01-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText)
    FH = mskFDesde.FormattedText
Else
    Tipo = "M"
    anio = Year(mskFDesde.FormattedText)
    Mes = Month(mskFDesde.FormattedText)
    FD = mskFDesde.FormattedText
    FH = mskFHasta.FormattedText
End If

SQL = "BEGIN estadis.Estimado_acum ( "
SQL = SQL & "'" & Tipo & "',"
SQL = SQL & anio & ","
SQL = SQL & Mes & ","
SQL = SQL & "" & Func.func_ToDate(FD) & ","
SQL = SQL & "" & Func.func_ToDate(FH) & " ); "
SQL = SQL & " End;"

If Aplicacion.EjecutarDAO(SQL) Then
    SQL = "SELECT cod_depn,"
    SQL = SQL & " decode(cod_depn,'INT','INT',cod_sdep) cod_sdep,"
    SQL = SQL & " tipo_porc,"
    SQL = SQL & " sum(valor) valor "
    SQL = SQL & " FROM estadis.resultados_estim "
    SQL = SQL & " group by cod_depn,decode(cod_depn,'INT','INT',cod_sdep),tipo_porc"
    
    If Aplicacion.ObtenerRsDAO(SQL, RsDataEstim) Then
        L_DecoEstim
    End If
    Aplicacion.TerminarConExitoTrans
    Aplicacion.CerrarDAO RsDataEstim
Else
    Aplicacion.TerminarConErrorTrans
End If

End Sub

Private Sub L_TratarVolados()

Dim SQL As String
Dim FD As String, FH As String

If optFechas(0).Value Then
    FD = "01-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText)
    FH = mskFDesde.FormattedText
Else
    FD = mskFDesde.FormattedText
    FH = mskFHasta.FormattedText
End If

L_ControlFecha


    If Control_BsAs Then
    SQL = " SELECT "
    SQL = SQL & "P.cod_depn,"
    SQL = SQL & "decode(P.cod_depn,'INT','INT',cod_ssdep) cod_sdep,"
'    SQL = SQL & " SUM(pax_AT_VS_vol.PAX_VOL) CANT_V "
    SQL = SQL & " SUM(cantidad) CANT_V "
    SQL = SQL & " From estadis.pax_volados P , ventas.apertura_sdep A"
    SQL = SQL & " Where P.local = A.cod_local and p.cod_depn = a.cod_depn and p.cod_sdep = a.cod_sdep  And "
    SQL = SQL & " fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
    SQL = SQL & " GROUP BY p.COD_DEPN,decode(P.cod_depn,'INT','INT',cod_ssdep)"
    
    If Aplicacion.ObtenerRsDAO(SQL, RsDataVol) Then
        L_DecoVolados 2
        Aplicacion.CerrarDAO RsDataVol
    End If
        
    Else
        labInfo.caption = "Datos de pasajeros transitados Actualizados al día " & Fch_BsAs
        Timer1.Enabled = True
    End If

End Sub

Private Sub L_TratarVoladosHIST()
Dim SQL As String
Dim FD As String, FH As String

If optFechas(0).Value Then
    FD = "01-01-" & Year(mskFDesde.FormattedText) - 1
    FH = Day(mskFDesde.FormattedText) & "-" & Month(mskFDesde.FormattedText) & "-" & Year(mskFDesde.FormattedText) - 1
Else
    FD = mskFDesdeAnt.FormattedText
    FH = mskFHastaAnt.FormattedText
End If

        SQL = " SELECT "
        SQL = SQL & "P.cod_depn,"
        SQL = SQL & "decode(P.cod_depn,'INT','INT',cod_ssdep) cod_sdep,"
        SQL = SQL & " SUM(p.cantidad) CANT_V "
        SQL = SQL & " From estadis.pax_volados P, ventas.apertura_sdep A "
        SQL = SQL & " Where P.local = A.cod_local and p.cod_depn = a.cod_depn and p.cod_sdep = a.cod_sdep  And"
        SQL = SQL & " fch_vuelo Between " & func_ToDate(FD) & " AND " & func_ToDate(FH)
        SQL = SQL & " GROUP BY P.COD_DEPN,decode(P.cod_depn,'INT','INT',cod_ssdep) "
        
        If Aplicacion.ObtenerRsDAO(SQL, RsDataVol) Then
            L_DecoVolados 3
            Aplicacion.CerrarDAO RsDataVol
        End If

End Sub


Private Sub L_VentaNeta()
Dim SQL As String


    
End Sub

Private Sub borGrBar_Click(Index As Integer)
Dim col_datos As Collection

    Set col_datos = New Collection
    Select Case Index
        Case 0
            L_LLenarColeccion col_datos, SprTotalBsAs
            FrmGrafbar.CargarGrafico total, "", col_datos
        Case 1
            L_LLenarColeccion col_datos, SprEzeA
            FrmGrafbar.CargarGrafico EZEA, "", col_datos
        Case 2
            L_LLenarColeccion col_datos, SprEzeB
            FrmGrafbar.CargarGrafico EZEB, "", col_datos
        Case 3
            L_LLenarColeccion col_datos, SprAep
            FrmGrafbar.CargarGrafico AERO, "", col_datos
        Case 4
            L_LLenarColeccion col_datos, SprInt
            FrmGrafbar.CargarGrafico INTE, "", col_datos
        Case 5
            L_LLenarColeccion col_datos, SprTotal
            FrmGrafbar.CargarGrafico total, "", col_datos
    End Select
        
    
End Sub

Private Sub botEjecutar_Click(Index As Integer)
Select Case Index
    Case 0
        L_Refrescar
    Case 1

        frdatos.Enabled = True
        botEjecutar(0).Enabled = True
        tabEspigon.Enabled = False
        L_LimpiarGrillas
        
    Case 2
        Unload Me
End Select
End Sub



Private Sub botExcel_Click()
    If optFechas(0).Value Then
        L_TratarExcel SprTotal, "INDICADORES ACUMULADOS ANUALES (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
        "del 1 de Ene al " & Format$(mskFDesde.FormattedText, "dd-mmm"), "", Exl_Blanco
    Else
        L_TratarExcel SprTotal, "INDICADORES ACUMULADOS MES CORRIENTE (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
        " del " & Format$(mskFDesde.FormattedText, "dd-mmm") & " al " & Format$(mskFHasta.FormattedText, "dd-mmm"), "", Exl_Blanco 'Exl_Ros '
    End If
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


Private Sub botHelpFDAnt_Click()
Dim fch As Date

If mskFDesdeAnt.Text <> "" Then
    fch = mskFDesdeAnt.FormattedText
Else
    fch = Date
End If

frmFecha.MuestroFormFecha fch

mskFDesdeAnt.Text = Format$(fch, FTOFECHA)

mskFDesdeAnt.SetFocus


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



Private Sub botHelpFHAnt_Click()
Dim fch As Date

If mskFHastaAnt.Text <> "" Then
    fch = mskFHastaAnt.FormattedText
Else
    fch = Date
End If

frmFecha.MuestroFormFecha fch

mskFHastaAnt.Text = Format$(fch, FTOFECHA)

mskFHastaAnt.SetFocus


End Sub

Private Sub chkNeto_Click()

If chkNeto.Value = 1 Then
    L_VentaNeta
Else

End If
End Sub

Private Sub Command1_Click()
    If optFechas(0).Value Then
        L_TratarExcel SprTotal, "INDICADORES ACUMULADOS ANUALES (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
        "del 1 de Ene al " & Format$(mskFDesde.FormattedText, "dd-mmm"), "", Exl_Blanco
    Else
        L_TratarExcel SprTotal, "INDICADORES ACUMULADOS MES CORRIENTE (" & Format$(mskFDesde.FormattedText, "yyyy") & ")", _
        " del " & Format$(mskFDesde.FormattedText, "dd-mmm") & " al " & Format$(mskFHasta.FormattedText, "dd-mmm"), "", Exl_Blanco 'Exl_Ros '
    End If
End Sub

Private Sub Form_Activate()
'FuncLocal_SeteoTABS tabEspigon

End Sub

Private Sub L_LLenarColeccion(ByRef col As Collection, spr As control)
Dim cl_dato As CLlgi
Dim valor As Variant
Dim item As Variant

spr.Row = spr.ActiveRow
spr.GetText 1, spr.Row, item

Set cl_dato = New CLlgi
    
spr.GetText 2, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
col.Add cl_dato
    
Set cl_dato = New CLlgi
spr.GetText 3, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
col.Add cl_dato

Set cl_dato = New CLlgi
spr.GetText 6, spr.Row, valor
cl_dato.Locale = item
cl_dato.DatoGral = valor
col.Add cl_dato

End Sub

Private Sub L_TratarExcel(spr As control, Titulo As String, subTit As String, Info As String, color As Integer)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim col As Integer
Dim fila As Integer
Dim i
Dim tit As Variant
Dim nombre As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

frmIndic.caption = Aplicacion.SeteoProceso(frmIndic.caption)

If nombre <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.Application.Visible = True
    
    ReDim titCol(spr.MaxCols)
    col = 1
    fila = 3
    
    For i = 1 To spr.MaxCols
        spr.GetText i, 0, tit
        titCol(i) = tit
    Next
    
    Exl_PonerValor AppExcel, 1, 1, Titulo
    rango = Exl_rangos(1, 1, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, color
    
    Exl_PonerValor AppExcel, fila, col, subTit
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    
    fila = fila + 1
    
    Exl_PonerValor AppExcel, fila, col, "Total Companía"
    rango = Exl_rangos(fila, fila, 1, spr.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    AppExcel.Application.Range(rango).Merge
        
    fila = fila + 1
    
    Exl_BajarGrillaExel spr, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + spr.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '-------------------------
    fila = fila + 1
    
    Exl_PonerValor AppExcel, fila, col, "Total Buenos Aires"
    rango = Exl_rangos(fila, fila, 1, SprTotalBsAs.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    AppExcel.Application.Range(rango).Merge
        
    fila = fila + 1
    
    Exl_BajarGrillaExel SprTotalBsAs, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprTotalBsAs.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "INTERNACIONAL A Llegadas"
    rango = Exl_rangos(fila, fila, 1, SprEzeA.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprEzeA, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprEzeA.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "INTERNACIONAL A Salidas "
    rango = Exl_rangos(fila, fila, 1, sprEzeAS.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel sprEzeAS, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + sprEzeAS.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
        
    '-----------------------------------
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "INTERNACIONAL B "
    rango = Exl_rangos(fila, fila, 1, SprEzeB.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba
    
    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprEzeB, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprEzeB.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '--------------------------------
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "AEROPARQUE "
    rango = Exl_rangos(fila, fila, 1, SprAep.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba
    
    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprAep, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprAep.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
        
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris

    '------------------------
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, col, "INTERIOR "
    rango = Exl_rangos(fila, fila, 1, SprInt.MaxCols)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    Exl_ColorInt AppExcel, rango, color
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LArr
    Exl_LineasPart AppExcel, rango, Exl_Linsimple, Exl_LAba

    AppExcel.Application.Range(rango).Merge
    
    fila = fila + 1
    Exl_BajarGrillaExel SprInt, AppExcel, fila, col, titCol
    rango = Exl_rangos(fila + 1, fila + SprInt.MaxRows, col + 1, spr.MaxCols)
    Exl_Format AppExcel, rango
    fila = fila + spr.MaxRows + 1
    rango = Exl_rangos(fila, fila, col, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, color
    
    rango = Exl_rangos(fila - spr.MaxRows, fila - 1, 6, spr.MaxCols)
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    '-----------------------------------

    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.Zoom = False
    
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
    AppExcel.SaveAs nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    frmIndic.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub


Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6300
Me.Width = 9300

If Day(Date) = 1 Then
    mskFDesde.Text = Format$(Func.func_Dia1SegunMes_Anio(Month(Date - 1), Year(Date - 1)), FTOFECHA)
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
    
    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio - 1, FTOFECHA)
    mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio - 1, FTOFECHA)
    
Else
    mskFDesde.Text = "01-" & Format$(Month(Date), "0#") & "-" & Format$(Year(Date), "####")
    mskFHasta.Text = Format$(Date - 1, FTOFECHA)
    
    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
    mskFHastaAnt.Text = Format$(CDate(mskFHasta.FormattedText) - Aplicacion.anio, FTOFECHA)
    
End If


L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmIndic"

End Sub


Private Sub Form_Unload(Cancel As Integer)
FuncLocal_SacarForm "frmIndic"
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
    Else
    'If Year(CDate(mskFDesdeAnt.FormattedText)) >= Year(Date) Then
    '    mskFDesdeAnt.Text = Format$(CDate(mskFDesde.FormattedText) - Aplicacion.anio, FTOFECHA)
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

Private Sub optFechas_Click(Index As Integer)
Select Case Index
    Case 0
        mskFHasta.Visible = False
        botHelpFH.Visible = False
        Label1(0).caption = "Fecha Tope"
        Label1(1).Visible = False
        frPerAnt.Visible = False
    Case 1
        mskFHasta.Visible = True
        botHelpFH.Visible = True
        Label1(0).caption = "Fecha Desde"
        Label1(1).Visible = True
        frPerAnt.Visible = True
End Select
End Sub

Private Sub tabEspigon_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    Select Case tabEspigon.Tab
        Case 0
            SprTotal.SetFocus
        Case 1
            SprEzeA.SetFocus
        Case 2
            sprEzeAS.SetFocus
        Case 3
            SprEzeB.SetFocus
        Case 4
            SprAep.SetFocus
    End Select
    
    
ErrT:
    Exit Sub



End Sub

Private Sub Timer1_Timer()
    
If Not Control_BsAs Then
    labInfo.Visible = Not labInfo.Visible
 ElseIf Not Control_Int Then
     If tabEspigon.Tab > 4 Then
        labInfo.Visible = Not labInfo.Visible
     Else
        labInfo.Visible = False
     End If
End If


End Sub


