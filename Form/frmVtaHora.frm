VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVtaHora 
   Caption         =   "Venta Horaria"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   9195
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
         Left            =   135
         Picture         =   "frmVtaHora.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmVtaHora.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmVtaHora.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   990
         Width           =   570
      End
   End
   Begin TabDlg.SSTab tabEspigon 
      Height          =   3915
      Left            =   150
      TabIndex        =   2
      Top             =   1365
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   6906
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   5
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
      TabCaption(0)   =   "RESUMEN"
      TabPicture(0)   =   "frmVtaHora.frx":0A26
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "botExcel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "sprTotCIA"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "EZE-INTA"
      TabPicture(1)   =   "frmVtaHora.frx":0A42
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "botExcelEzeA"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "tabLocA"
      Tab(1).Control(1).Enabled=   0   'False
      TabCaption(2)   =   "EZE-INTB"
      TabPicture(2)   =   "frmVtaHora.frx":0A5E
      Tab(2).ControlCount=   2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "botExcelEzeb"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "SSTab1"
      Tab(2).Control(1).Enabled=   0   'False
      TabCaption(3)   =   "AEROP."
      TabPicture(3)   =   "frmVtaHora.frx":0A7A
      Tab(3).ControlCount=   2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTab2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "botExcelAep"
      Tab(3).Control(1).Enabled=   -1  'True
      Begin FPSpread.vaSpread sprTotCIA 
         Height          =   2880
         Left            =   180
         OleObjectBlob   =   "frmVtaHora.frx":0A96
         TabIndex        =   39
         Top             =   720
         Width           =   7620
      End
      Begin VB.CommandButton botExcelEzeA 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67035
         Picture         =   "frmVtaHora.frx":1144
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   705
         Width           =   765
      End
      Begin VB.CommandButton botExcelEzeb 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -66990
         Picture         =   "frmVtaHora.frx":16D6
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   645
         Width           =   765
      End
      Begin VB.CommandButton botExcelAep 
         Caption         =   "Excel"
         Height          =   510
         Left            =   -67020
         Picture         =   "frmVtaHora.frx":1C68
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   630
         Width           =   765
      End
      Begin TabDlg.SSTab tabLocA 
         Height          =   3345
         Left            =   -74895
         TabIndex        =   8
         Top             =   405
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5900
         _Version        =   327680
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   441
         BackColor       =   12632256
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "L Lleg."
         TabPicture(0)   =   "frmVtaHora.frx":21FA
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprEzeA(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "L06"
         TabPicture(1)   =   "frmVtaHora.frx":2216
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprEzeA(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "L Sal"
         TabPicture(2)   =   "frmVtaHora.frx":2232
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprEzeA(2)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "L08"
         TabPicture(3)   =   "frmVtaHora.frx":224E
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprEzeA(3)"
         Tab(3).Control(0).Enabled=   0   'False
         TabCaption(4)   =   "TOTAL"
         TabPicture(4)   =   "frmVtaHora.frx":226A
         Tab(4).ControlCount=   1
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "sprTot(0)"
         Tab(4).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2820
            Index           =   0
            Left            =   -74895
            OleObjectBlob   =   "frmVtaHora.frx":2286
            TabIndex        =   9
            Top             =   405
            Width           =   7605
         End
         Begin FPSpread.vaSpread sprTot 
            Height          =   2880
            Index           =   0
            Left            =   105
            OleObjectBlob   =   "frmVtaHora.frx":2933
            TabIndex        =   30
            Top             =   375
            Width           =   7620
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2820
            Index           =   1
            Left            =   -74910
            OleObjectBlob   =   "frmVtaHora.frx":2FE1
            TabIndex        =   33
            Top             =   405
            Width           =   7605
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2820
            Index           =   2
            Left            =   -74925
            OleObjectBlob   =   "frmVtaHora.frx":368E
            TabIndex        =   34
            Top             =   375
            Width           =   7605
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2820
            Index           =   3
            Left            =   -74925
            OleObjectBlob   =   "frmVtaHora.frx":3D3B
            TabIndex        =   35
            Top             =   360
            Width           =   7605
         End
      End
      Begin VB.CommandButton botExcel 
         Caption         =   "Excel"
         Height          =   510
         Left            =   7890
         Picture         =   "frmVtaHora.frx":43E8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   780
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3375
         Left            =   -74895
         TabIndex        =   10
         Top             =   405
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   5953
         _Version        =   327680
         Tabs            =   4
         TabsPerRow      =   5
         TabHeight       =   441
         BackColor       =   12632256
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "L01"
         TabPicture(0)   =   "frmVtaHora.frx":497A
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "sprEzeB(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "L02"
         TabPicture(1)   =   "frmVtaHora.frx":4996
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprEzeB(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "L03"
         TabPicture(2)   =   "frmVtaHora.frx":49B2
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "sprEzeB(2)"
         Tab(2).Control(0).Enabled=   0   'False
         TabCaption(3)   =   "TOTAL"
         TabPicture(3)   =   "frmVtaHora.frx":49CE
         Tab(3).ControlCount=   1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "sprTot(1)"
         Tab(3).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprTot 
            Height          =   2880
            Index           =   1
            Left            =   -74895
            OleObjectBlob   =   "frmVtaHora.frx":49EA
            TabIndex        =   31
            Top             =   420
            Width           =   7650
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2865
            Index           =   0
            Left            =   90
            OleObjectBlob   =   "frmVtaHora.frx":5098
            TabIndex        =   28
            Top             =   375
            Width           =   7605
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   6
            Left            =   -74880
            OleObjectBlob   =   "frmVtaHora.frx":5736
            TabIndex        =   13
            Top             =   390
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   5
            Left            =   -74880
            OleObjectBlob   =   "frmVtaHora.frx":5C63
            TabIndex        =   12
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   4
            Left            =   -74880
            OleObjectBlob   =   "frmVtaHora.frx":6190
            TabIndex        =   11
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2865
            Index           =   1
            Left            =   -74910
            OleObjectBlob   =   "frmVtaHora.frx":66BD
            TabIndex        =   36
            Top             =   375
            Width           =   7605
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2865
            Index           =   2
            Left            =   -74925
            OleObjectBlob   =   "frmVtaHora.frx":6D5B
            TabIndex        =   37
            Top             =   375
            Width           =   7605
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3420
         Left            =   -74880
         TabIndex        =   14
         Top             =   390
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   6033
         _Version        =   327680
         Tab             =   2
         TabsPerRow      =   5
         TabHeight       =   441
         BackColor       =   12632256
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "L09"
         TabPicture(0)   =   "frmVtaHora.frx":73F9
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "sprAep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "L14"
         TabPicture(1)   =   "frmVtaHora.frx":7415
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "sprAep(1)"
         Tab(1).Control(0).Enabled=   0   'False
         TabCaption(2)   =   "TOTAL"
         TabPicture(2)   =   "frmVtaHora.frx":7431
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "sprTot(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread sprTot 
            Height          =   2880
            Index           =   2
            Left            =   90
            OleObjectBlob   =   "frmVtaHora.frx":744D
            TabIndex        =   32
            Top             =   405
            Width           =   7620
         End
         Begin FPSpread.vaSpread sprAep 
            Height          =   2925
            Index           =   0
            Left            =   -74895
            OleObjectBlob   =   "frmVtaHora.frx":7AFB
            TabIndex        =   29
            Top             =   375
            Width           =   7605
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2655
            Index           =   3
            Left            =   -74850
            OleObjectBlob   =   "frmVtaHora.frx":8199
            TabIndex        =   15
            Top             =   420
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeB 
            Height          =   2655
            Index           =   4
            Left            =   -74850
            OleObjectBlob   =   "frmVtaHora.frx":86C6
            TabIndex        =   16
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   7
            Left            =   -74880
            OleObjectBlob   =   "frmVtaHora.frx":8BF3
            TabIndex        =   17
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   8
            Left            =   -74880
            OleObjectBlob   =   "frmVtaHora.frx":9120
            TabIndex        =   18
            Top             =   405
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprEzeA 
            Height          =   2655
            Index           =   9
            Left            =   -74880
            OleObjectBlob   =   "frmVtaHora.frx":964D
            TabIndex        =   19
            Top             =   390
            Width           =   6375
         End
         Begin FPSpread.vaSpread sprAep 
            Height          =   2925
            Index           =   1
            Left            =   -74910
            OleObjectBlob   =   "frmVtaHora.frx":9B7A
            TabIndex        =   38
            Top             =   345
            Width           =   7605
         End
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
         Caption         =   "Fecha de  Consultas"
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
         Height          =   1290
         Left            =   120
         TabIndex        =   3
         Top             =   30
         Width           =   8040
         Begin VB.Frame Frame2 
            Caption         =   "Grupos"
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
            Height          =   615
            Index           =   0
            Left            =   3825
            TabIndex        =   23
            Top             =   150
            Width           =   2940
            Begin VB.ComboBox cboGrupo 
               Height          =   315
               Left            =   660
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   225
               Width           =   1845
            End
         End
         Begin VB.CommandButton botHelpFD 
            Height          =   345
            Left            =   3015
            Picture         =   "frmVtaHora.frx":A218
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   360
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskFDesde 
            Height          =   285
            Left            =   1440
            TabIndex        =   21
            Top             =   375
            Width           =   1500
            _ExtentX        =   2646
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
            Caption         =   "Fecha "
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
            Left            =   210
            TabIndex        =   22
            Top             =   375
            Width           =   1185
         End
      End
   End
End
Attribute VB_Name = "frmVtaHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsData  As Recordset

Dim RsDataAnt  As Recordset
Dim RsDataEstim  As Recordset
Dim RsDataVol  As Recordset

Dim ConsultaComo As Integer



Private Function L_NombreLocal(Esp As String, ind As Integer)
Select Case Esp
    Case AERO
        Select Case ind
            Case 0
                L_NombreLocal = "Local 9"
            Case 1
                L_NombreLocal = "Local 14"
        End Select
    Case EZEA
        Select Case ind
            Case 0
                L_NombreLocal = "Local 5"
            Case 1
                L_NombreLocal = "Local 6"
            Case 2
                L_NombreLocal = "Local 7"
            Case 3
                L_NombreLocal = "Local 8"
        
        End Select
    Case EZEB
        Select Case ind
            Case 0
                L_NombreLocal = "Local 1"
            Case 1
                L_NombreLocal = "Local 2"
            Case 2
                L_NombreLocal = "Local 3"
        
        End Select
    
End Select
End Function


Private Sub L_TratarExcel(titulo As String, subTit As String, Esp As String, CantCol As Integer)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer, filaant As Integer
Dim i As Integer
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmVtaHora.caption = Aplicacion.SeteoProceso(frmVtaHora.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.application.Visible = True
    
    ReDim titCol(CantCol)
    Col = 1
    fila = 3
    
    Exl_PonerValor AppExcel, 1, 1, titulo
    rango = Exl_rangos(1, 1, 1, CantCol)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, CantCol)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    
    fila = fila + 2
    Exl_PonerValor AppExcel, fila, Col, "Grupo :" & cboGrupo.Text
    
    Select Case Esp
        Case AERO
            For i = 1 To CantCol
                sprAep(0).GetText i, 0, tit
                titCol(i) = tit
            Next
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información del Espigón "
               
               fila = fila + 2
               
               Exl_BajarGrillaExel SprTot(2), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + SprTot(2).MaxRows
               rango = Exl_rangos(fila, fila, Col, SprTot(2).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               rango = Exl_rangos(filaant, fila, 4, 4)
               Exl.Exl_Format AppExcel, rango
               rango = Exl_rangos(filaant, fila, 7, 7)
               Exl.Exl_Format AppExcel, rango

           For i = 0 To 1
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información del Local : " & L_NombreLocal(AERO, i)
               
               fila = fila + 2
               
               Exl_BajarGrillaExel sprAep(i), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + sprAep(i).MaxRows
               rango = Exl_rangos(fila, fila, Col, sprAep(i).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               rango = Exl_rangos(filaant, fila, 4, 4)
               Exl.Exl_Format AppExcel, rango
               rango = Exl_rangos(filaant, fila, 7, 7)
               Exl.Exl_Format AppExcel, rango
               
           Next
           Exl_AnchoCol AppExcel, 1, 1, 10
           
        Case EZEA
            For i = 1 To CantCol
                sprEzeA(0).GetText i, 0, tit
                titCol(i) = tit
            Next
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información del Espigón "
               
               fila = fila + 2
               
               Exl_BajarGrillaExel SprTot(0), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + SprTot(0).MaxRows
               rango = Exl_rangos(fila, fila, Col, SprTot(0).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               rango = Exl_rangos(filaant, fila, 4, 4)
               Exl.Exl_Format AppExcel, rango
               rango = Exl_rangos(filaant, fila, 7, 7)
               Exl.Exl_Format AppExcel, rango
       
           For i = 0 To 3
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información sobre :" & L_NombreLocal(EZEA, i)
               
               fila = fila + 2

               Exl_BajarGrillaExel sprEzeA(i), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + sprEzeA(i).MaxRows
               rango = Exl_rangos(fila, fila, Col, sprEzeA(i).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               rango = Exl_rangos(filaant, fila, 4, 4)
               Exl.Exl_Format AppExcel, rango
               rango = Exl_rangos(filaant, fila, 7, 7)
               Exl.Exl_Format AppExcel, rango
               
           Next
           Exl_AnchoCol AppExcel, 1, 1, 10
        
        Case EZEB
            For i = 1 To CantCol
                sprEzeB(0).GetText i, 0, tit
                titCol(i) = tit
            Next
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información del Espigón "
               
               fila = fila + 2
               
               Exl_BajarGrillaExel SprTot(1), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + SprTot(1).MaxRows
               rango = Exl_rangos(fila, fila, Col, SprTot(1).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               rango = Exl_rangos(filaant, fila, 4, 4)
               Exl.Exl_Format AppExcel, rango
               rango = Exl_rangos(filaant, fila, 7, 7)
               Exl.Exl_Format AppExcel, rango
       
           For i = 0 To 2
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información sobre :" & L_NombreLocal(EZEB, i)
               
               fila = fila + 2

               Exl_BajarGrillaExel sprEzeB(i), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + sprEzeB(i).MaxRows
               rango = Exl_rangos(fila, fila, Col, sprEzeB(i).MaxCols)
               Exl_ColorInt AppExcel, rango, Exl_Gris
                                               
               rango = Exl_rangos(filaant, fila, 4, 4)
               Exl.Exl_Format AppExcel, rango
               rango = Exl_rangos(filaant, fila, 7, 7)
               Exl.Exl_Format AppExcel, rango
               
           Next
           Exl_AnchoCol AppExcel, 1, 1, 10
        
        Case "TOT"
            For i = 1 To CantCol
                sprTotCIA.GetText i, 0, tit
                titCol(i) = tit
            Next
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información Total  "
               Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
               
               fila = fila + 2
               
               Exl_BajarGrillaExel sprTotCIA, AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + 25
               rango = Exl_rangos(fila, fila, Col, sprTotCIA.MaxCols + 1)
               Exl_ColorInt AppExcel, rango, Exl_Gris
               Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
               Exl_LetraColor AppExcel, rango, Exl_Blanco
               
               Exl_AnchoCol AppExcel, 1, 1, 10
    End Select
    With AppExcel.Application.ActiveSheet.PageSetup
'        .PrintTitleRows = "$1:$" & Trim(str(filaTit))
        .CenterHorizontally = True
        .TopMargin = Exl_TopMargen
'        .BottomMargin = Exl_BotMargen
'        .CenterFooter = "Página &P de &N"
'        .PrintGridlines = False
    End With
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
        
    AppExcel.SaveAs NOMBRE & ".xls"
'    AppExcel.Workbooks.Open nombre & ".xls"
    Set AppExcel = Nothing
    
End If

ErrorExl:

    frmVtaHora.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub

            
Private Sub L_TratarExcelProceso(titulo As String, subTit As String, Esp As String, CantCol As Integer)
Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer, filaant As Integer
Dim i As Integer
Dim tit As Variant
Dim NOMBRE As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

frmVtaHora.caption = Aplicacion.SeteoProceso(frmVtaHora.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    
    'AppExcel.Application.Visible = True
    
    ReDim titCol(CantCol)
    Col = 1
    fila = 3
    
    Exl_PonerValor AppExcel, 1, 1, titulo
    rango = Exl_rangos(1, 1, 1, CantCol)
    Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
    AppExcel.Application.Range(rango).Merge
    Exl_Lineas AppExcel, rango, Exl_Linsimple
    Exl_ColorInt AppExcel, rango, Exl_Gris
    
    Exl_PonerValor AppExcel, fila, Col, subTit
    rango = Exl_rangos(fila, fila, 1, CantCol)
    Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
    Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
    
'    AppExcel.Application.Range(rango).Merge
'    Exl_Lineas AppExcel, rango, Exl_Linsimple
        
    
    fila = fila + 2
'    Exl_PonerValor AppExcel, fila, col, "Grupo :" & cboGrupo.Text
    
    Select Case Esp
        Case AERO
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información del Espigón "
               Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
               
               fila = fila + 2
               
               L_BajarGrillaExel SprTot(2), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + 25
               rango = Exl_rangos(fila, fila, Col, SprTot(2).MaxCols + 1)
               Exl_ColorInt AppExcel, rango, Exl_Gris
               Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
               Exl_LetraColor AppExcel, rango, Exl_Blanco
               'rango = Exl_rangos(filaant, fila, 8, 8)
               'Exl.Exl_Format AppExcel, rango

           For i = 0 To 1
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, L_NombreLocal(AERO, i)
               Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
               
               fila = fila + 2
               
               L_BajarGrillaExel sprAep(i), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + 25
               rango = Exl_rangos(fila, fila, Col, sprAep(i).MaxCols + 1)
               Exl_ColorInt AppExcel, rango, Exl_Gris
               Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
               Exl_LetraColor AppExcel, rango, Exl_Blanco
                                               
               'rango = Exl_rangos(filaant, fila, 8, 8)
               'Exl.Exl_Format AppExcel, rango
               
           Next
           
        Case EZEA
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información del Espigón "
               Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
               
               fila = fila + 2
               
               L_BajarGrillaExel SprTot(0), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + 25
               rango = Exl_rangos(fila, fila, Col, SprTot(0).MaxCols + 1)
               Exl_ColorInt AppExcel, rango, Exl_Gris
               Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
               Exl_LetraColor AppExcel, rango, Exl_Blanco
                                               
               'rango = Exl_rangos(filaant, fila, 8, 8)
               'Exl.Exl_Format AppExcel, rango
       
           For i = 0 To 3
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, L_NombreLocal(EZEA, i)
               Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
               
               fila = fila + 2

               L_BajarGrillaExel sprEzeA(i), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + 25
               rango = Exl_rangos(fila, fila, Col, sprEzeA(i).MaxCols + 1)
               Exl_ColorInt AppExcel, rango, Exl_Gris
               Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
               Exl_LetraColor AppExcel, rango, Exl_Blanco
                                               
               'rango = Exl_rangos(filaant, fila, 8, 8)
               'Exl.Exl_Format AppExcel, rango
               
           Next
        
        Case EZEB
            For i = 1 To CantCol
                sprEzeB(0).GetText i, 0, tit
                titCol(i) = tit
            Next
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información del Espigón "
               Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
               
               fila = fila + 2
               
               L_BajarGrillaExel SprTot(1), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + 25
               rango = Exl_rangos(fila, fila, Col, SprTot(1).MaxCols + 1)
               Exl_ColorInt AppExcel, rango, Exl_Gris
               Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
               Exl_LetraColor AppExcel, rango, Exl_Blanco
                                               
               'rango = Exl_rangos(filaant, fila, 8, 8)
               'Exl.Exl_Format AppExcel, rango
       
           For i = 0 To 2
               
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, L_NombreLocal(EZEB, i)
               Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
               
               fila = fila + 2

               L_BajarGrillaExel sprEzeB(i), AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + 25
               rango = Exl_rangos(fila, fila, Col, sprEzeB(i).MaxCols + 1)
               Exl_ColorInt AppExcel, rango, Exl_Gris
               Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
               Exl_LetraColor AppExcel, rango, Exl_Blanco
                                               
               'rango = Exl_rangos(filaant, fila, 8, 8)
               'Exl.Exl_Format AppExcel, rango
               
           Next
        Case "TOT"
            For i = 1 To CantCol
                sprTotCIA.GetText i, 0, tit
                titCol(i) = tit
            Next
               fila = fila + 2
               rango = Exl_rangos(fila, fila + 2, 1, 1)
        
               Exl_PonerValor AppExcel, fila, Col, "Información Total  "
               Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
               
               fila = fila + 2
               
               L_BajarGrillaExel sprTotCIA, AppExcel, fila, Col, titCol
               filaant = fila
               fila = fila + 25
               rango = Exl_rangos(fila, fila, Col, sprTotCIA.MaxCols + 1)
               Exl_ColorInt AppExcel, rango, Exl_Gris
               Exl_Letra AppExcel, rango, NEGRITA, 10, "Arial"
               Exl_LetraColor AppExcel, rango, Exl_Blanco
        
    End Select
           
           Exl_AnchoCol AppExcel, 1, 1, 10
           Exl_AnchoCol AppExcel, 3, 4, 17
           Exl_AnchoCol AppExcel, 7, 7, 17
           Exl_AnchoCol AppExcel, 9, 9, 17
    
    With AppExcel.Application.ActiveSheet.PageSetup
'        .PrintTitleRows = "$1:$" & Trim(str(filaTit))
        .CenterHorizontally = True
        .TopMargin = Exl_TopMargen
'        .BottomMargin = Exl_BotMargen
'        .CenterFooter = "Página &P de &N"
'        .PrintGridlines = False
        .Orientation = 2
    End With
    If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
        AppExcel.PrintOut
    End If
        
    AppExcel.SaveAs NOMBRE & ".xls"
'    AppExcel.Workbooks.Open nombre & ".xls"
    Set AppExcel = Nothing
    
End If

ErrorExl:

    frmVtaHora.caption = Aplicacion.SeteoFin
    Exit Sub
    
End Sub

Public Sub L_BajarGrillaExel(spr As control, apl As Object, FilaInit As Integer, ColInit As Integer, titCol() As String)
Dim gr_fila As Integer
Dim gr_col As Integer, gr_aux_col As Integer
Dim dato As Variant, hora As Integer
Dim rango As String, RangoArmado As String

hora = 0
gr_fila = 1

      Exl_PonerValor apl, FilaInit, 1, Format(mskFDesde.FormattedText, "dddd")
      Exl_PonerValor apl, FilaInit, 2, "Rango"
      Exl_PonerValor apl, FilaInit, 3, "Dotación Teórica"
      Exl_PonerValor apl, FilaInit, 4, "Dotación Real"
      Exl_PonerValor apl, FilaInit, 5, "Importes"
      Exl_PonerValor apl, FilaInit, 6, "Ticket"
      Exl_PonerValor apl, FilaInit, 7, "Vuelos/Bodega"
      Exl_PonerValor apl, FilaInit, 8, "Promedio Ticket"
      Exl_PonerValor apl, FilaInit, 9, "Productividad x Pers"

'For gr_fila = 0 To spr.MaxRows
 For hora = 0 To 23
    If hora = 0 Then
    RangoArmado = Format$(hora, "0") & " - " & Format$(hora + 1, "00")
    Else
    RangoArmado = Format$(hora, "00") & " - " & Format$(hora + 1, "00")
    End If
    
    spr.GetText 2, gr_fila, dato
        
    If dato = RangoArmado Then
    Exl_PonerValor apl, FilaInit + hora + 1, 2, dato
    
    spr.GetText 1, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 1, dato
        
    spr.GetText 3, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 5, dato
    
    spr.GetText 4, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 6, dato
        
    spr.GetText 5, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 8, dato
        
        If gr_fila <= spr.MaxRows Then
            gr_fila = gr_fila + 1
        End If
    Else
        Exl_PonerValor apl, FilaInit + hora + 1, 1, Format(mskFDesde.FormattedText, "dd-mm-yy")
        Exl_PonerValor apl, FilaInit + hora + 1, 2, RangoArmado
    End If
Next
    Exl_PonerValor apl, FilaInit + hora + 1, 1, "Total"
    spr.GetText 3, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 5, dato
    
    spr.GetText 4, gr_fila, dato
    Exl_PonerValor apl, FilaInit + hora + 1, 6, dato

rango = Exl_rangos(FilaInit, FilaInit, ColInit, spr.MaxCols + 1)

Exl_Justificacion apl, rango, Exl_CentroCol, Exl_CentroVert, False

Exl_ColorInt apl, rango, Exl_Gris
Exl_Letra apl, rango, NEGRITA, 10, "Ms Serif"
Exl_LetraColor apl, rango, Exl_Blanco
rango = Exl_rangos(FilaInit, 25 + FilaInit, ColInit, spr.MaxCols + 1)

Exl_Lineas apl, rango, Exl_Linsimple

rango = Exl_rangos(FilaInit + 6, FilaInit + 6, ColInit, spr.MaxCols + 1)
apl.Application.Range(rango).Borders(4).Weight = 4
rango = Exl_rangos(FilaInit + 14, FilaInit + 14, ColInit, spr.MaxCols + 1)
apl.Application.Range(rango).Borders(4).Weight = 4

rango = Exl_rangos(FilaInit + 1, FilaInit + 25, 5, 5)
apl.Application.Range(rango).NumberFormat = "$ #,##0_ ;[Red]-#,##0 "
rango = Exl_rangos(FilaInit + 1, FilaInit + 25, 8, 9)
apl.Application.Range(rango).NumberFormat = "$ #,##0_ ;[Red]-#,##0 "

End Sub

            

Private Function L_Armarcondicion(anio As Integer) As String
Dim Cond
Dim fechaDesde As String
Dim fechaHasta As String
Dim cant


'Cond = " WHERE fch_vta between " & func_ToDate(fechaDesde) & " And " & func_ToDate(fechaHasta)

L_Armarcondicion = Cond

End Function

Private Sub L_LimpiarGrillas()
Dim i
       
sprAep(0).MaxRows = 0
sprAep(1).MaxRows = 0

sprEzeA(0).MaxRows = 0
sprEzeA(1).MaxRows = 0
sprEzeA(2).MaxRows = 0
sprEzeA(3).MaxRows = 0

sprEzeB(0).MaxRows = 0
sprEzeB(1).MaxRows = 0
sprEzeB(2).MaxRows = 0

SprTot(0).MaxRows = 0
SprTot(1).MaxRows = 0
SprTot(2).MaxRows = 0

sprTotCIA.MaxRows = 0

End Sub

Private Sub L_PonerenGrilla(spr As control, item As String, valor As Single)
    spr.MaxRows = spr.MaxRows + 1
    spr.SetText 1, spr.MaxRows, Trim(item)
    spr.SetText 2, spr.MaxRows, str(valor)
End Sub

Private Sub L_Refrescar()
Dim sql As String

'On Error GoTo ErrInd:

frmVtaHora.caption = Aplicacion.SeteoProceso(frmVtaHora.caption)

If ConsultaComo = 1 Then
If cboGrupo.Text = "TODOS" Then
    sql = "SELECT cod_depn,"
    sql = sql & " cod_sdep,"
    sql = sql & " decode(cod_local,'L21','L05','L06','L05','L07','L22','L08','L22',cod_Local) cod_local,"
    sql = sql & " rango,"
    sql = sql & " sum(ticket) ticket,"
    sql = sql & " sum(imp) imp, "
    sql = sql & " fch_proceso "
    If Month(Date) = Month(mskFDesde.FormattedText) And Year(Date) = Year(mskFDesde.FormattedText) Then
        sql = sql & " FROM estadis.View_Ticket_hora "
    Else
        sql = sql & " FROM estadis.View_Ticket_hora_hist "
    End If
    sql = sql & " WHERE fch_ticket = " & Func.func_ToDate(mskFDesde.FormattedText)
    sql = sql & " GROUP BY fch_proceso,cod_depn,cod_sdep,decode(cod_local,'L21','L05','L06','L05','L07','L22','L08','L22',cod_Local),rango "
    sql = sql & " ORDER BY fch_proceso,cod_depn,cod_sdep,cod_local,rango"
Else
    sql = "SELECT cod_depn,"
    sql = sql & " cod_sdep,"
    sql = sql & " decode(cod_local,'L21','L05','L06','L05','L07','L22','L08','L22',cod_Local) cod_local,"
    sql = sql & " rango,"
    sql = sql & " ticket,"
    sql = sql & " imp, "
    sql = sql & " fch_proceso"
    If Month(Date) = Month(mskFDesde.FormattedText) And _
       Year(Date) = Year(mskFDesde.FormattedText) Then
        sql = sql & " FROM estadis.View_Ticket_hora "
    Else
        sql = sql & " FROM estadis.View_Ticket_hora_hist "
    End If
    sql = sql & " WHERE fch_ticket = " & Func.func_ToDate(mskFDesde.FormattedText)
    sql = sql & " And grupo = '" & cboGrupo.Text & "' "
    sql = sql & " ORDER BY fch_proceso,cod_depn,cod_sdep,decode(cod_local,'L21','L05','L06','L05','L07','L22','L08','L22',cod_Local),rango"
End If
Else
    sql = "SELECT cod_depn,"
    sql = sql & " cod_sdep,"
    sql = sql & " decode(cod_local,'L21','L05','L06','L05','L07','L22','L08','L22',cod_Local) cod_local,"
    sql = sql & " rango,"
    sql = sql & " sum(ticket) ticket,"
    sql = sql & " sum(imp) imp, "
    sql = sql & " fch_proceso "
    If Month(Date) = Month(mskFDesde.FormattedText) And _
       Year(Date) = Year(mskFDesde.FormattedText) Then
        sql = sql & " FROM estadis.View_Ticket_hora "
    Else
        sql = sql & " FROM estadis.View_Ticket_hora_hist "
    End If
    sql = sql & " WHERE fch_proceso = " & Func.func_ToDate(mskFDesde.FormattedText)
    sql = sql & " GROUP BY fch_proceso,cod_depn,cod_sdep,decode(cod_local,'L21','L05','L06','L05','L07','L22','L08','L22',cod_Local),rango "
    sql = sql & " ORDER BY fch_proceso,cod_depn,cod_sdep,cod_local,rango"
End If

If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    
    If Aplicacion.CantReg(RsData) > 0 Then
        frdatos.Enabled = False
        botEjecutar(0).Enabled = False
        tabEspigon.Enabled = True
        L_DecoEspigon
    End If
    Aplicacion.CerrarDAO RsData
    
    L_TratarTotales

End If

ErrInd:
    frmVtaHora.caption = Aplicacion.SeteoFin
    Exit Sub

End Sub

Private Sub L_Resaltar()
Dim fila As Integer
Dim i

For fila = 1 To 6
    For i = 0 To 3
        Spread.spread_ResaltarCelda sprEzeA(i), 4, fila
        Spread.spread_ResaltarCelda sprEzeA(i), 5, fila
    Next
    For i = 0 To 2
        Spread.spread_ResaltarCelda sprEzeB(i), 4, fila
        Spread.spread_ResaltarCelda sprEzeB(i), 5, fila
    Next
    For i = 0 To 1
        Spread.spread_ResaltarCelda sprAep(i), 4, fila
        Spread.spread_ResaltarCelda sprAep(i), 5, fila
    Next
Next
End Sub


Private Sub L_DecoEspigon()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant

Do While Not RsData.EOF
    Select Case RsData!cod_depn
        Case DSLoc(1).Dep
            Select Case RsData!cod_local
                Case Left(DSLoc(1).locales(1), 3)
                    sprAep(0).MaxRows = sprAep(0).MaxRows + 1
                    sprAep(0).SetText 1, sprAep(0).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
                    sprAep(0).SetText 2, sprAep(0).MaxRows, Trim(RsData!rango)
                    sprAep(0).SetText 3, sprAep(0).MaxRows, str(RsData!imp)
                    sprAep(0).SetText 4, sprAep(0).MaxRows, str(RsData!ticket)
                    If RsData!ticket > 0 Then
                        sprAep(0).SetText 5, sprAep(0).MaxRows, str(RsData!imp / RsData!ticket)
                    End If
                    Spread_TotalesGrillaAcum sprAep(0), 3, 6, sprAep(0).MaxRows
                    Spread_TotalesGrillaAcum sprAep(0), 4, 7, sprAep(0).MaxRows
                
                Case Left(DSLoc(1).locales(2), 3)
                    sprAep(1).MaxRows = sprAep(1).MaxRows + 1
                    sprAep(1).SetText 1, sprAep(1).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
                    sprAep(1).SetText 2, sprAep(1).MaxRows, Trim(RsData!rango)
                    sprAep(1).SetText 3, sprAep(1).MaxRows, str(RsData!imp)
                    sprAep(1).SetText 4, sprAep(1).MaxRows, str(RsData!ticket)
                    If RsData!ticket > 0 Then
                        sprAep(1).SetText 5, sprAep(1).MaxRows, str(RsData!imp / RsData!ticket)
                    End If
                    Spread_TotalesGrillaAcum sprAep(1), 3, 6, sprAep(1).MaxRows
                    Spread_TotalesGrillaAcum sprAep(1), 4, 7, sprAep(1).MaxRows

            End Select
        Case DSLoc(2).Dep
            Select Case RsData!Cod_Sdep
                Case DSLoc(2).Sdep
                    Select Case RsData!cod_local
                        Case Left(DSLoc(2).locales(1), 3)
                            sprEzeA(0).MaxRows = sprEzeA(0).MaxRows + 1
                            sprEzeA(0).SetText 1, sprEzeA(0).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
                            sprEzeA(0).SetText 2, sprEzeA(0).MaxRows, Trim(RsData!rango)
                            sprEzeA(0).SetText 3, sprEzeA(0).MaxRows, str(RsData!imp)
                            sprEzeA(0).SetText 4, sprEzeA(0).MaxRows, str(RsData!ticket)
                            If RsData!ticket > 0 Then
                                sprEzeA(0).SetText 5, sprEzeA(0).MaxRows, str(RsData!imp / RsData!ticket)
                            End If
                            Spread_TotalesGrillaAcum sprEzeA(0), 3, 6, sprEzeA(0).MaxRows
                            Spread_TotalesGrillaAcum sprEzeA(0), 4, 7, sprEzeA(0).MaxRows
                        
                        Case Left(DSLoc(2).locales(3), 3)
                            sprEzeA(1).MaxRows = sprEzeA(1).MaxRows + 1
                            sprEzeA(1).SetText 1, sprEzeA(1).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
                            sprEzeA(1).SetText 2, sprEzeA(1).MaxRows, Trim(RsData!rango)
                            sprEzeA(1).SetText 3, sprEzeA(1).MaxRows, str(RsData!imp)
                            sprEzeA(1).SetText 4, sprEzeA(1).MaxRows, str(RsData!ticket)
                            If RsData!ticket > 0 Then
                                sprEzeA(1).SetText 5, sprEzeA(1).MaxRows, str(RsData!imp / RsData!ticket)
                            End If
                            Spread_TotalesGrillaAcum sprEzeA(1), 3, 6, sprEzeA(1).MaxRows
                            Spread_TotalesGrillaAcum sprEzeA(1), 4, 7, sprEzeA(1).MaxRows
                        
                        Case Left(DSLoc(2).locales(5), 3)
                            sprEzeA(2).MaxRows = sprEzeA(2).MaxRows + 1
                            sprEzeA(2).SetText 1, sprEzeA(2).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
                            sprEzeA(2).SetText 2, sprEzeA(2).MaxRows, Trim(RsData!rango)
                            sprEzeA(2).SetText 3, sprEzeA(2).MaxRows, str(RsData!imp)
                            sprEzeA(2).SetText 4, sprEzeA(2).MaxRows, str(RsData!ticket)
                            If RsData!ticket > 0 Then
                                sprEzeA(2).SetText 5, sprEzeA(2).MaxRows, str(RsData!imp / RsData!ticket)
                            End If
                            Spread_TotalesGrillaAcum sprEzeA(2), 3, 6, sprEzeA(2).MaxRows
                            Spread_TotalesGrillaAcum sprEzeA(2), 4, 7, sprEzeA(2).MaxRows
                        
                        Case Left(DSLoc(2).locales(7), 3)
                            sprEzeA(3).MaxRows = sprEzeA(3).MaxRows + 1
                            sprEzeA(3).SetText 1, sprEzeA(3).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
                            sprEzeA(3).SetText 2, sprEzeA(3).MaxRows, Trim(RsData!rango)
                            sprEzeA(3).SetText 3, sprEzeA(3).MaxRows, str(RsData!imp)
                            sprEzeA(3).SetText 4, sprEzeA(3).MaxRows, str(RsData!ticket)
                            If RsData!ticket > 0 Then
                                sprEzeA(3).SetText 5, sprEzeA(3).MaxRows, str(RsData!imp / RsData!ticket)
                            End If
                            Spread_TotalesGrillaAcum sprEzeA(3), 3, 6, sprEzeA(3).MaxRows
                            Spread_TotalesGrillaAcum sprEzeA(3), 4, 7, sprEzeA(3).MaxRows
                    
                    End Select
                Case DSLoc(3).Sdep
                    Select Case RsData!cod_local
                        Case Left(DSLoc(3).locales(1), 3)
                            sprEzeB(0).MaxRows = sprEzeB(0).MaxRows + 1
                            sprEzeB(0).SetText 1, sprEzeB(0).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
                            sprEzeB(0).SetText 2, sprEzeB(0).MaxRows, Trim(RsData!rango)
                            sprEzeB(0).SetText 3, sprEzeB(0).MaxRows, str(RsData!imp)
                            sprEzeB(0).SetText 4, sprEzeB(0).MaxRows, str(RsData!ticket)
                            If RsData!ticket > 0 Then
                                sprEzeB(0).SetText 5, sprEzeB(0).MaxRows, str(RsData!imp / RsData!ticket)
                            End If
                            Spread_TotalesGrillaAcum sprEzeB(0), 3, 6, sprEzeB(0).MaxRows
                            Spread_TotalesGrillaAcum sprEzeB(0), 4, 7, sprEzeB(0).MaxRows
                        
                        Case Left(DSLoc(3).locales(4), 3)
                            sprEzeB(1).MaxRows = sprEzeB(1).MaxRows + 1
                            sprEzeB(1).SetText 1, sprEzeB(1).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
                            sprEzeB(1).SetText 2, sprEzeB(1).MaxRows, Trim(RsData!rango)
                            sprEzeB(1).SetText 3, sprEzeB(1).MaxRows, str(RsData!imp)
                            sprEzeB(1).SetText 4, sprEzeB(1).MaxRows, str(RsData!ticket)
                            If RsData!ticket > 0 Then
                                sprEzeB(1).SetText 5, sprEzeB(1).MaxRows, str(RsData!imp / RsData!ticket)
                            End If
                            Spread_TotalesGrillaAcum sprEzeB(1), 3, 6, sprEzeB(1).MaxRows
                            Spread_TotalesGrillaAcum sprEzeB(1), 4, 7, sprEzeB(1).MaxRows
                        
                        Case Left(DSLoc(3).locales(5), 3)
                            sprEzeB(2).MaxRows = sprEzeB(2).MaxRows + 1
                            sprEzeB(2).SetText 1, sprEzeB(2).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
                            sprEzeB(2).SetText 2, sprEzeB(2).MaxRows, Trim(RsData!rango)
                            sprEzeB(2).SetText 3, sprEzeB(2).MaxRows, str(RsData!imp)
                            sprEzeB(2).SetText 4, sprEzeB(2).MaxRows, str(RsData!ticket)
                            If RsData!ticket > 0 Then
                                sprEzeB(2).SetText 5, sprEzeB(2).MaxRows, str(RsData!imp / RsData!ticket)
                            End If
                            Spread_TotalesGrillaAcum sprEzeB(2), 3, 6, sprEzeB(2).MaxRows
                            Spread_TotalesGrillaAcum sprEzeB(2), 4, 7, sprEzeB(2).MaxRows
                    
                    End Select
            End Select
        
    End Select
                                
        RsData.MoveNext

Loop

For i = 0 To 3
    Spread_TotalesGrillas sprEzeA(i), sprEzeA(i).MaxCols - 5, 2
Next
For i = 0 To 2
    Spread_TotalesGrillas sprEzeB(i), sprEzeB(i).MaxCols - 5, 2
Next
For i = 0 To 1
    Spread_TotalesGrillas sprAep(i), sprAep(i).MaxCols - 5, 2
Next


End Sub

Private Sub L_DecoTotal()
Dim fecha As String
Dim i As Integer, indDep As Integer
Dim dato As Variant

Do While Not RsData.EOF
    Select Case RsData!cod_depn
        Case DSLoc(1).Dep
            SprTot(2).MaxRows = SprTot(2).MaxRows + 1
            SprTot(2).SetText 1, SprTot(2).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
            SprTot(2).SetText 2, SprTot(2).MaxRows, Trim(RsData!rango)
            SprTot(2).SetText 3, SprTot(2).MaxRows, str(RsData!imp)
            SprTot(2).SetText 4, SprTot(2).MaxRows, str(RsData!ticket)
            If RsData!ticket > 0 Then
                SprTot(2).SetText 5, SprTot(2).MaxRows, str(RsData!imp / RsData!ticket)
            End If
            Spread_TotalesGrillaAcum SprTot(2), 3, 6, SprTot(2).MaxRows
            Spread_TotalesGrillaAcum SprTot(2), 4, 7, SprTot(2).MaxRows
        
        Case DSLoc(2).Dep
            Select Case RsData!Cod_Sdep
                Case DSLoc(2).Sdep
                    SprTot(0).MaxRows = SprTot(0).MaxRows + 1
                    SprTot(0).SetText 1, SprTot(0).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
                    SprTot(0).SetText 2, SprTot(0).MaxRows, Trim(RsData!rango)
                    SprTot(0).SetText 3, SprTot(0).MaxRows, str(RsData!imp)
                    SprTot(0).SetText 4, SprTot(0).MaxRows, str(RsData!ticket)
                    If RsData!ticket > 0 Then
                        SprTot(0).SetText 5, SprTot(0).MaxRows, str(RsData!imp / RsData!ticket)
                    End If
                    Spread_TotalesGrillaAcum SprTot(0), 3, 6, SprTot(0).MaxRows
                    Spread_TotalesGrillaAcum SprTot(0), 4, 7, SprTot(0).MaxRows
                
                Case DSLoc(3).Sdep
                    SprTot(1).MaxRows = SprTot(1).MaxRows + 1
                    SprTot(1).SetText 1, SprTot(1).MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
                    SprTot(1).SetText 2, SprTot(1).MaxRows, Trim(RsData!rango)
                    SprTot(1).SetText 3, SprTot(1).MaxRows, str(RsData!imp)
                    SprTot(1).SetText 4, SprTot(1).MaxRows, str(RsData!ticket)
                    If RsData!ticket > 0 Then
                        SprTot(1).SetText 5, SprTot(1).MaxRows, str(RsData!imp / RsData!ticket)
                    End If
                    Spread_TotalesGrillaAcum SprTot(1), 3, 6, SprTot(1).MaxRows
                    Spread_TotalesGrillaAcum SprTot(1), 4, 7, SprTot(1).MaxRows
           End Select
        Case "TOT"
            sprTotCIA.MaxRows = sprTotCIA.MaxRows + 1
            sprTotCIA.SetText 1, sprTotCIA.MaxRows, Format$(RsData!fch_proceso, "dd-mm-yy")
            sprTotCIA.SetText 2, sprTotCIA.MaxRows, Trim(RsData!rango)
            sprTotCIA.SetText 3, sprTotCIA.MaxRows, str(RsData!imp)
            sprTotCIA.SetText 4, sprTotCIA.MaxRows, str(RsData!ticket)
            If RsData!ticket > 0 Then
                sprTotCIA.SetText 5, sprTotCIA.MaxRows, str(RsData!imp / RsData!ticket)
            End If
            Spread_TotalesGrillaAcum sprTotCIA, 3, 6, sprTotCIA.MaxRows
            Spread_TotalesGrillaAcum sprTotCIA, 4, 7, sprTotCIA.MaxRows
        
    End Select
                                
        RsData.MoveNext

Loop

For i = 0 To 2
    Spread_TotalesGrillas SprTot(i), SprTot(i).MaxCols - 5, 2
Next
    Spread_TotalesGrillas sprTotCIA, sprTotCIA.MaxCols - 5, 2

End Sub


Private Sub L_TratarTotales()
Dim sql As String
Dim rs As Recordset

  If ConsultaComo = 1 Then
    sql = "SELECT cod_depn,"
    sql = sql & " cod_sdep,"
    sql = sql & " rango,"
    sql = sql & " sum(ticket) ticket,"
    sql = sql & " sum(imp) imp, "
    sql = sql & " fch_proceso "
    If Month(Date) = Month(mskFDesde.FormattedText) And Year(Date) = Year(mskFDesde.FormattedText) Then
        sql = sql & " FROM estadis.View_Ticket_hora "
    Else
        sql = sql & " FROM estadis.View_Ticket_hora_hist "
    End If

    sql = sql & " WHERE fch_ticket = " & Func.func_ToDate(mskFDesde.FormattedText)
    If cboGrupo.Text <> "TODOS" Then
        sql = sql & " And grupo = '" & cboGrupo.Text & "' "
    End If
    sql = sql & " GROUP BY cod_depn,cod_sdep,rango,fch_proceso "
    sql = sql & " Union "
    sql = sql & " SELECT 'TOT' cod_depn,"
    sql = sql & " 'TOT' cod_sdep,"
    sql = sql & " rango,"
    sql = sql & " sum(ticket) ticket,"
    sql = sql & " sum(imp) imp, "
    sql = sql & " fch_proceso "
    If Month(Date) = Month(mskFDesde.FormattedText) And Year(Date) = Year(mskFDesde.FormattedText) Then
        sql = sql & " FROM estadis.View_Ticket_hora "
    Else
        sql = sql & " FROM estadis.View_Ticket_hora_hist "
    End If
    sql = sql & " WHERE fch_ticket = " & Func.func_ToDate(mskFDesde.FormattedText)
    sql = sql & " And cod_depn <> 'INT' "
    If cboGrupo.Text <> "TODOS" Then
        sql = sql & " And grupo = '" & cboGrupo.Text & "' "
    End If
    sql = sql & " GROUP BY rango,fch_proceso "
    sql = sql & " ORDER BY fch_proceso,cod_depn,cod_sdep,rango"
  
  Else
    sql = "SELECT cod_depn,"
    sql = sql & " cod_sdep,"
    sql = sql & " rango,"
    sql = sql & " sum(ticket) ticket,"
    sql = sql & " sum(imp) imp, "
    sql = sql & " fch_proceso "
    If Month(Date) = Month(mskFDesde.FormattedText) And Year(Date) = Year(mskFDesde.FormattedText) Then
        sql = sql & " FROM estadis.View_Ticket_hora "
    Else
        sql = sql & " FROM estadis.View_Ticket_hora_hist "
    End If

    sql = sql & " WHERE fch_proceso = " & Func.func_ToDate(mskFDesde.FormattedText)
    sql = sql & " GROUP BY cod_depn,cod_sdep,rango,fch_proceso "
    sql = sql & " Union "
    sql = sql & " SELECT 'TOT' cod_depn,"
    sql = sql & " 'TOT' cod_sdep,"
    sql = sql & " rango,"
    sql = sql & " sum(ticket) ticket,"
    sql = sql & " sum(imp) imp, "
    sql = sql & " fch_proceso "
    If Month(Date) = Month(mskFDesde.FormattedText) And Year(Date) = Year(mskFDesde.FormattedText) Then
        sql = sql & " FROM estadis.View_Ticket_hora "
    Else
        sql = sql & " FROM estadis.View_Ticket_hora_hist "
    End If
    sql = sql & " WHERE fch_proceso = " & Func.func_ToDate(mskFDesde.FormattedText)
    sql = sql & " And cod_depn <> 'INT' "
    sql = sql & " GROUP BY rango,fch_proceso "
    sql = sql & " ORDER BY fch_proceso,cod_depn,cod_sdep,rango "
  End If
If Aplicacion.ObtenerRsDAO(sql, RsData) Then
    If Aplicacion.CantReg(RsData) > 0 Then
        L_DecoTotal
    End If
    Aplicacion.CerrarDAO RsData
End If
End Sub

Public Sub ModoConsulta(Modo As Integer)

ConsultaComo = Modo

Me.Show

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
If ConsultaComo = 1 Then
    L_TratarExcel " TOTAL BUENOS AIRES ", "Informe de Ventas Horarias del día (" & mskFDesde.FormattedText & " )", "TOT", sprTotCIA.MaxCols
Else
    L_TratarExcelProceso " TOTAL BUENOS AIRES ", "Informe de Ventas Horarias del día " & mskFDesde.FormattedText & " ", "TOT", sprTotCIA.MaxCols + 1
End If

End Sub

Private Sub botExcelAep_Click()
If ConsultaComo = 1 Then
 L_TratarExcel "AEROPARQUE", "Informe de Ventas Horarias del día (" & mskFDesde.FormattedText & " )", AERO, SprTot(2).MaxCols
Else
 L_TratarExcelProceso "AEROPARQUE", "Informe de Ventas Horarias del día " & mskFDesde.FormattedText & " ", AERO, SprTot(2).MaxCols + 1
End If
End Sub

Private Sub botExcelEzeA_Click()

If ConsultaComo = 1 Then
    L_TratarExcel "ESPIGON INTERNACIONAL 'A'", "Informe de Ventas Horarias del día (" & mskFDesde.FormattedText & " )", EZEA, SprTot(0).MaxCols
Else
    L_TratarExcelProceso "ESPIGON INTERNACIONAL 'A'", "Informe de Ventas Horarias del día " & mskFDesde.FormattedText & " ", EZEA, SprTot(0).MaxCols + 1
End If

End Sub

Private Sub botExcelEzeb_Click()
If ConsultaComo = 1 Then
    L_TratarExcel "ESPIGON INTERNACIONAL 'B'", "Informe de Ventas Horarias del día (" & mskFDesde.FormattedText & " )", EZEB, SprTot(1).MaxCols
Else
    L_TratarExcelProceso "ESPIGON INTERNACIONAL 'B'", "Informe de Ventas Horarias del día " & mskFDesde.FormattedText & " ", EZEB, SprTot(1).MaxCols + 1
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


Private Sub Form_Activate()
FuncLocal_SeteoTABS tabEspigon
tabLocA.TabVisible(1) = False
tabLocA.TabVisible(3) = False
End Sub

Private Sub Form_Load()
Dim i

Me.Left = 50
Me.Top = 100
Me.Height = 6000
Me.Width = 9300

'If Day(Date) = 1 Then
'    mskFDesde.Text = Format$(Func.func_Dia1SegunMes_Anio(Month(Date - 1), Year(Date - 1)), FTOFECHA)
'Else
'    mskFDesde.Text = "01-" & Month(Date) & "-" & Format$(Year(Date), "####")
'End If
mskFDesde.Text = Format$(Date - 1, FTOFECHA)

cboGrupo.AddItem "TODOS"
cboGrupo.AddItem "A"
cboGrupo.AddItem "B"
cboGrupo.AddItem "C"

cboGrupo.ListIndex = 0

L_LimpiarGrillas

frmPrincipal.lstForms.AddItem "frmIndic"

If ConsultaComo = 2 Then
    cboGrupo.Visible = False
Else

End If
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

End Sub





Private Sub tabEspigon_Click(PreviousTab As Integer)
On Error GoTo ErrT:

    Select Case tabEspigon.Tab
        Case 0
            'sprTotal.SetFocus
        Case 1
            'sprEzeA.SetFocus
        Case 2
            'sprEzeB.SetFocus
        Case 3
            'sprAep.SetFocus
    End Select
    
    
ErrT:
    Exit Sub



End Sub

