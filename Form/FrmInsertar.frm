VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form FrmAdmDeFlight 
   Caption         =   "Administración de Pasajeros Viajados"
   ClientHeight    =   5775
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   6045
   Icon            =   "FrmInsertar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5775
   ScaleWidth      =   6045
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   150
      TabIndex        =   16
      Top             =   0
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   17
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "k"
            Object.ToolTipText     =   "Nueva Seleción"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "l"
            Object.ToolTipText     =   "Buscar"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "a"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "b"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "c"
            Object.ToolTipText     =   "Registro Siguiente"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "e"
            Object.ToolTipText     =   "Ultimo Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "f"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "g"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "h"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "i"
            Object.ToolTipText     =   "Abortar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "m"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "n"
            Object.ToolTipText     =   "Vista en Grilla"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "j"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "FrmInsertar.frx":0442
   End
   Begin VB.TextBox txtReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4125
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   480
      Width           =   465
   End
   Begin VB.TextBox txtCantReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4935
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   480
      Width           =   480
   End
   Begin VB.Frame FraCabecera 
      Height          =   5160
      Left            =   135
      TabIndex        =   18
      Top             =   510
      Width           =   5835
      Begin VB.Frame FraTipoVuelo 
         Caption         =   "Tipo de vuelo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1455
         Left            =   4185
         TabIndex        =   25
         Top             =   300
         Width           =   1530
         Begin VB.OptionButton Opttipovuelo 
            Caption         =   "Todos"
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
            Index           =   2
            Left            =   195
            TabIndex        =   5
            Top             =   990
            Width           =   900
         End
         Begin VB.OptionButton Opttipovuelo 
            Caption         =   "Llegada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   195
            TabIndex        =   3
            Top             =   285
            Width           =   1005
         End
         Begin VB.OptionButton Opttipovuelo 
            Caption         =   "Salida"
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
            Index           =   1
            Left            =   195
            TabIndex        =   4
            Top             =   630
            Width           =   870
         End
      End
      Begin VB.Frame FramAerop 
         Height          =   1470
         Left            =   150
         TabIndex        =   22
         Top             =   285
         Width           =   3945
         Begin VB.ListBox LstEspigon 
            Height          =   255
            Left            =   3285
            TabIndex        =   32
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin MSMask.MaskEdBox MskFecha 
            Height          =   315
            Left            =   1815
            TabIndex        =   0
            Top             =   180
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Format          =   "dd-mm-yyyy"
            Mask            =   "##-##-####"
            PromptChar      =   " "
         End
         Begin VB.CommandButton CmdFecha 
            Height          =   390
            Left            =   3330
            Picture         =   "FrmInsertar.frx":045E
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   150
            Width           =   465
         End
         Begin VB.ComboBox CboCodAeropuerto 
            Height          =   315
            ItemData        =   "FrmInsertar.frx":05E8
            Left            =   1785
            List            =   "FrmInsertar.frx":05EA
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   615
            Width           =   1995
         End
         Begin VB.ComboBox CboEspigon 
            Height          =   315
            Left            =   1785
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1020
            Width           =   2010
         End
         Begin VB.Label LblFecha 
            BackColor       =   &H00E0E0E0&
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
            Height          =   300
            Left            =   150
            TabIndex        =   28
            Top             =   195
            Width           =   1605
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
            Left            =   150
            TabIndex        =   24
            Top             =   615
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
            Left            =   165
            TabIndex        =   23
            Top             =   1035
            Width           =   1575
         End
      End
      Begin VB.Frame FramPax 
         Height          =   3210
         Left            =   150
         TabIndex        =   19
         Top             =   1755
         Width           =   5580
         Begin VB.OptionButton OptTrip 
            Caption         =   "N"
            Height          =   240
            Index           =   1
            Left            =   5070
            TabIndex        =   39
            Top             =   2355
            Width           =   390
         End
         Begin VB.OptionButton OptTrip 
            Caption         =   "E"
            Height          =   240
            Index           =   0
            Left            =   4620
            TabIndex        =   38
            Top             =   2340
            Width           =   360
         End
         Begin VB.ComboBox cboLocal 
            Height          =   315
            Left            =   4245
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   615
            Width           =   1065
         End
         Begin MSMask.MaskEdBox txtPaxExtranjeros 
            Height          =   300
            Left            =   3300
            TabIndex        =   10
            Top             =   1500
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin VB.ListBox LstCia 
            Height          =   255
            Left            =   4680
            TabIndex        =   33
            Top             =   285
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox TxtHoraReal 
            Height          =   315
            Left            =   1800
            TabIndex        =   9
            Top             =   1065
            Width           =   1425
         End
         Begin VB.ComboBox CboCia 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   210
            Width           =   3510
         End
         Begin VB.TextBox TxtVuelo 
            Height          =   300
            Left            =   1815
            TabIndex        =   7
            Top             =   630
            Width           =   1440
         End
         Begin MSMask.MaskEdBox txtPaxNacionales 
            Height          =   300
            Left            =   3300
            TabIndex        =   11
            Top             =   1935
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TxtPaxTransito 
            Height          =   300
            Left            =   3315
            TabIndex        =   14
            Top             =   2700
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtTrip 
            Height          =   300
            Left            =   3300
            TabIndex        =   12
            Top             =   2310
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cantidad de Tripulación :"
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
            Left            =   150
            TabIndex        =   37
            Top             =   2325
            Width           =   3105
         End
         Begin VB.Label LblPaxTransito 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cantidad de Pax en Tránsito :"
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
            Left            =   150
            TabIndex        =   36
            Top             =   2730
            Width           =   3105
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Local :"
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
            Index           =   1
            Left            =   3480
            TabIndex        =   35
            Top             =   630
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
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
            Left            =   4710
            TabIndex        =   34
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label LblHoraReal 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hora Real :"
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
            Left            =   150
            TabIndex        =   31
            Top             =   1065
            Width           =   1590
         End
         Begin VB.Label LblCia 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Companía :"
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
            Left            =   150
            TabIndex        =   30
            Top             =   210
            Width           =   1605
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vuelo :"
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
            Left            =   150
            TabIndex        =   29
            Top             =   645
            Width           =   1605
         End
         Begin VB.Label LblPaxExt 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cantidad de Pax Extranjeros :"
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
            Left            =   135
            TabIndex        =   21
            Top             =   1515
            Width           =   3105
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cantidad de Pax Nacionales :"
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
            Left            =   150
            TabIndex        =   20
            Top             =   1950
            Width           =   3105
         End
      End
      Begin VB.Label de 
         Caption         =   "  de"
         Height          =   255
         Left            =   4425
         TabIndex        =   26
         Top             =   15
         Width           =   390
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "chk"
      Height          =   195
      Left            =   4560
      TabIndex        =   17
      Top             =   525
      Visible         =   0   'False
      Width           =   990
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":05EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":0906
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":0C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":0F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":1254
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":156E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":1888
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":1BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":1EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":21D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":22E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":23FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmInsertar.frx":2C2C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAdmDeFlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As Recordset

Dim cl_volados As CLVolados

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$

Public Sub altas()
    SetearBotonesAltas
    Modo = "ALTA"
    FrmAdmDeFlight.caption = FrmAdmDeFlight.caption & " - Altas -"
    Me.Show 1
End Sub
Private Sub L_AltasDatos()

If L_TodoCargado Then
   FrmAdmDeFlight.caption = Aplicacion.SeteoProceso(FrmAdmDeFlight.caption)

   Aplicacion.ComienzoTrans

   MeLlenarObjeto

   If cl_volados.Insert_Pax_Volados() Then
      Aplicacion.TerminarConExitoTrans
      chk.Value = False
    
      NuevaSeleccion
    
    Else
     Aplicacion.TerminarConErrorTrans
   End If


   FrmAdmDeFlight.caption = Aplicacion.SeteoFin

  Else
   MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
 End If

End Sub

Private Sub L_DatosGrilla()
Dim i
Dim nro As Integer

On Error GoTo DG:

nro = frmGridFly.DatosGrilla(sqlGral$)

If nro > 0 Then
    rs.MoveFirst
    For i = 1 To nro - 1
        rs.MoveNext
    Next
    MellenarPantalla
    txtReg.Text = nro
    MeSetearBotonesToolBar

DG:
    Exit Sub
End If

End Sub

Private Sub L_SetearLocal()
Dim sql As String
Dim tipo As String
Dim rs As Recordset

On Error GoTo Err:

Call Aplicacion.SeteoProceso("")
If Opttipovuelo(0).Value Then 'Llegada
    tipo = "L"
Else 'Salida
    tipo = "S"
End If

sql = "SELECT cod_loc From BAIRES.LOCAL "
sql = sql & " WHERE COD_DEPN = '" & CboCodAeropuerto.Text & "' "
sql = sql & " AND COD_SDEP = '" & LstEspigon.List(CboEspigon.ListIndex) & "' "
sql = sql & " AND TIPO_LOC in ('" & tipo & "','T') "
sql = sql & " "

cboLocal.Clear
If LstEspigon.List(CboEspigon.ListIndex) = "INTA" And tipo = "L" Then
            cboLocal.AddItem "L05"
            cboLocal.AddItem "L06"
            cboLocal.ListIndex = 0
Else
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            cboLocal.AddItem rs!cod_loc
            rs.MoveNext
        Loop
        cboLocal.ListIndex = 0
    End If
    Aplicacion.CerrarDAO rs
End If
End If

Err:
    Call Aplicacion.SeteoFin
    Exit Sub
    
End Sub

Private Sub MeImpDatos()
Dim nom As String, NombreArchivo As String


'On Error GoTo ErrFoto:
'
'Aplicacion.SeteoProceso ("")
'
'    NombreArchivo = RutaFotos & "P" & txtLegajo.Text & ".bmp"
'    Nom = txtApe.Text
'
'    If Dir(NombreArchivo) <> "" Then
'        Image1.Picture = LoadPicture(NombreArchivo)
'        Printer.PaintPicture Image1, 8000, 2000, 2800, 2200
'    End If
'
'
'Printer.FontBold = True
'
'Printer.CurrentX = 10
'Printer.CurrentY = 10
'Printer.FontSize = 10
'
'Printer.Print "  "
'
'Printer.CurrentX = 1000
'Printer.CurrentY = 1000
'Printer.FontSize = 18
'
'Printer.Print txtApe.Text & ", " & txtNom.Text
'
'Printer.CurrentX = 1000
'Printer.CurrentY = 2000
'Printer.FontSize = 10
''Printer.FontBold = False
'Printer.Print "Legajo  : "
'
'Printer.FontBold = False

'Printer.CurrentX = 10
'Printer.CurrentY = 10
'Printer.FontSize = 10
'
'Printer.Print "  "
'
'Printer.CurrentX = 2000
'Printer.CurrentY = 2000
'Printer.FontSize = 10
'Printer.Print txtLegajo.Text
'
'Printer.EndDoc
'
'ErrFoto:
'    Aplicacion.SeteoFin
'    Exit Sub
        
End Sub


Public Sub modificacion()
Modo = "MOD"
FrmAdmDeFlight.caption = FrmAdmDeFlight.caption & " - Modificación y Bajas -"
Me.Show 1
End Sub
Private Sub NuevaSeleccion()
Dim i%

If Modo = "MOD" Then
    
    SetBotonesGeneral False
    CboCia.ListIndex = -1
    MskFecha.Text = ""
Else
    If chk.Value Then
        If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
            L_AltasDatos
        End If
    End If

End If
'Limpiar campos de pantallas

TxtVuelo.Text = ""
TxtHoraReal.Text = ""
txtPaxExtranjeros.Text = 0
txtPaxNacionales.Text = 0
TxtPaxTransito.Text = 0
txtTrip.Text = 0

If Modo = "MOD" Then
   CboCodAeropuerto.ListIndex = -1
   Opttipovuelo.item(2).Visible = True
   Opttipovuelo(2) = True
End If

Set cl_volados = New CLVolados
MskFecha.SetFocus
chk.Value = False

End Sub

Private Sub MeAbortarMod()
    
    SeteoBotonesMod True
    
    Tollbar.Buttons(2).Enabled = False
    
    MeSetearBotonesToolBar
    
    MellenarPantalla
    
End Sub

Private Sub MeActualizar()

If L_TodoCargado Then

Aplicacion.SeteoProceso ("Actualizando")

Aplicacion.ComienzoTrans

MeLlenarObjeto


If cl_volados.update_pax_volados() Then
    Aplicacion.TerminarConExitoTrans
    SeteoBotonesMod True

    If MeReconsultar > 0 Then

    Tollbar.Buttons(2).Enabled = False

    MeSetearBotonesToolBar
    Else
            NuevaSeleccion
    End If

Else
    Aplicacion.TerminarConErrorTrans
End If


Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub MeCargarDatos()
Dim sql As String

'FrmAdmDeFlight.caption = Aplicacion.SeteoProceso(FrmAdmDeFlight.caption)
Aplicacion.SeteoProceso (FrmAdmDeFlight.caption)
        
    CondConsulta = ArmarCondicion

    sqlGral = ""
    'sqlGral$ = sqlGral$ & " SELECT * FROM estadis.pax_volados"
    sqlGral = sqlGral & "SELECT DISTINCT PXN.fch_vuelo "
    sqlGral = sqlGral & ",pxn.cod_depn "
    sqlGral = sqlGral & ",pxn.cod_sdep "
    sqlGral = sqlGral & ",pxn.cod_cia_aerea  "
    sqlGral = sqlGral & ",pxn.cod_vuelo "
    sqlGral = sqlGral & ",pxn.tipo "
    sqlGral = sqlGral & ",pxn.hora_vuelo "
    sqlGral = sqlGral & ",pxn.local "
    sqlGral = sqlGral & ",PXN.cantidad CN "
    sqlGral = sqlGral & ",PXE.cantidad CE "
    sqlGral = sqlGral & ",PXE.transito CT "
    sqlGral = sqlGral & "FROM estadis.pax_volados PXN ,estadis.pax_volados PXE "
    sqlGral = sqlGral & "WHERE "
    sqlGral = sqlGral & "PXN.nacionalidad='N' and "
    sqlGral = sqlGral & "PXE.cantidad = (select cantidad from estadis.pax_volados PXE "
    sqlGral = sqlGral & "WHERE "
    sqlGral = sqlGral & "PXE.fch_vuelo=PXN.fch_vuelo "
    sqlGral = sqlGral & "and PXE.cod_cia_aerea=PXN.cod_cia_aerea "
    sqlGral = sqlGral & "and PXE.cod_vuelo=PXN.cod_vuelo "
    sqlGral = sqlGral & "and PXE.cod_depn=PXN.cod_depn "
    sqlGral = sqlGral & "and PXE.cod_sdep=PXN.cod_sdep "
    sqlGral = sqlGral & "AND PXE.TIPO = PXN.TIPO "
    sqlGral = sqlGral & "and PXE.nacionalidad = 'E') "
    sqlGral = sqlGral & "and  "
    sqlGral = sqlGral & "PXE.fch_vuelo=PXN.fch_vuelo "
    sqlGral = sqlGral & "and PXE.cod_cia_aerea=PXN.cod_cia_aerea "
    sqlGral = sqlGral & "and PXE.cod_vuelo=PXN.cod_vuelo "
    sqlGral = sqlGral & "and PXE.cod_depn=PXN.cod_depn "
    sqlGral = sqlGral & "and PXE.cod_sdep=PXN.cod_sdep "
    sqlGral = sqlGral & "AND PXE.TIPO = PXN.TIPO "
    sqlGral = sqlGral & CondConsulta
    sqlGral = sqlGral & "Order By fch_vuelo,cod_depn,cod_sdep "
    
If Aplicacion.ObtenerRsDAO(sqlGral, rs) Then
    txtCantReg.Text = Aplicacion.CantReg(rs)
    If txtCantReg.Text > 0 Then
        txtReg.Text = 1
        SetBotonesGeneral True
        MellenarPantalla
        MeSetearBotonesToolBar
    Else
        txtReg.Text = 0
    End If
End If

FrmAdmDeFlight.caption = Aplicacion.SeteoFin
Aplicacion.SeteoFin
'FraCabecera.Enabled = False

End Sub
Private Sub MeEliminar()

Set cl_volados = New CLVolados
MeLlenarObjeto

Aplicacion.SeteoProceso (FrmAdmDeFlight.caption)

Aplicacion.ComienzoTrans

If cl_volados.delete_pax_volados Then
    Aplicacion.TerminarConExitoTrans
    If MeReconsultar > 0 Then

    Tollbar.Buttons(2).Enabled = False

    MeSetearBotonesToolBar
    Else
            NuevaSeleccion
    End If
Else
    Aplicacion.TerminarConErrorTrans
End If

Aplicacion.SeteoFin

End Sub

Private Sub MeLlenarObjeto()

cl_volados.fch_vuelo = MskFecha.FormattedText

If Opttipovuelo(0).Value Then
   cl_volados.tipo = "L"
 Else
   cl_volados.tipo = "S"
End If

cl_volados.cod_depn = CboCodAeropuerto.Text
cl_volados.Cod_Sdep = LstEspigon.List(CboEspigon.ListIndex)
cl_volados.Cod_Cia_Aerea = LstCia.List(CboCia.ListIndex)
cl_volados.Cod_Vuelo = TxtVuelo.Text
cl_volados.hora_vuelo = TxtHoraReal.Text
If OptTrip(0).Value = True Then
   cl_volados.cantE = txtPaxExtranjeros.Text - txtTrip.Text
 Else
   cl_volados.cantE = txtPaxExtranjeros.Text
End If

If OptTrip(1).Value = True Then
   cl_volados.cantN = txtPaxNacionales.Text - txtTrip.Text
 Else
   cl_volados.cantN = txtPaxNacionales.Text
End If

'cl_volados.cantN = txtPaxNacionales.Text
cl_volados.cantT = TxtPaxTransito.Text
cl_volados.locales = cboLocal.Text


End Sub



Private Function L_TodoCargado() As Boolean
 
 If MskFecha.Text <> "" _
    And CboCodAeropuerto.Text <> "" _
    And CboEspigon <> "" And CboCia.Text <> "" _
    And TxtVuelo.Text <> "" And TxtHoraReal.Text <> "" _
    And txtPaxExtranjeros.Text <> "" _
    And txtPaxNacionales.Text <> "" _
    And (Opttipovuelo(0).Value Or Opttipovuelo(1).Value) Then
    L_TodoCargado = True
  Else
    L_TodoCargado = False
 End If
 
End Function

Private Sub MellenarPantalla()
Dim Trans As Integer

MskFecha.Text = Format$(rs!fch_vuelo, FTOFECHA)
Func_SetearCboSTR CboCodAeropuerto, rs!cod_depn
Func_SetearCboConLst CboEspigon, LstEspigon, rs!Cod_Sdep

Select Case rs!tipo
     Case Is = "S"
        Opttipovuelo.item(1) = True
        Opttipovuelo.item(2).Visible = False
     Case Is = "L"
        Opttipovuelo.item(0) = True
        Opttipovuelo.item(2).Visible = False
End Select

L_SetearLocal

Func_SetearCboConLst CboCia, LstCia, Trim(str(rs!Cod_Cia_Aerea))
TxtVuelo.Text = rs!Cod_Vuelo
TxtHoraReal.Text = rs!hora_vuelo
txtPaxNacionales.Text = rs!cn
Trans = IIf(IsNull(rs!ct), 0, rs!ct)
TxtPaxTransito.Text = Trans
txtPaxExtranjeros.Text = (rs!ce - Trans)
Func_SetearCboSTR cboLocal, rs!local

'cboLocal.Text = IIf(IsNull(rs!local), "", rs!local)

End Sub

Private Sub SetBotonesGeneral(valor As Boolean)
    
    Tollbar.Buttons(1).Enabled = valor
    Tollbar.Buttons(2).Enabled = Not valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = valor
    Tollbar.Buttons(6).Enabled = valor
    Tollbar.Buttons(7).Enabled = valor

    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor

    Tollbar.Buttons(16).Enabled = valor
'    TollBar.Buttons(13).Enabled = Not valor
'

'habilitar frames

FraCabecera.Enabled = Not valor
Opttipovuelo(2).Enabled = Not valor

    
    If Not valor Then
        txtReg.Text = 0
        txtCantReg.Text = 0
    End If

End Sub

Private Function ArmarCondicion()
Dim Con$

Con$ = ""
If MskFecha.Text <> "" Then
   Con$ = Con$ & " and PXN.fch_vuelo = " & func_ToDate(MskFecha.FormattedText)
End If

If CboCodAeropuerto.Text <> "" Then
   Con$ = Con$ & " and PXN.cod_depn = '" & CboCodAeropuerto.Text & "'"
End If

If CboEspigon.Text <> "" Then
   Con$ = Con$ & " and PXN.cod_sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
End If


If Opttipovuelo(2) Then
   'No agrega ninguna condición para el tipo de vuelo
   Else
     If Opttipovuelo(0) Then
        Con$ = Con$ & " and PXN.tipo = 'L'"
       Else
        Con$ = Con$ & " and PXN.tipo = 'S'"
     End If
End If

If CboCia.Text <> "" Then
   Con$ = Con$ & " and PXN.cod_cia_aerea = " & LstCia.List(CboCia.ListIndex)
End If

If TxtVuelo.Text <> "" Then
   Con$ = Con$ & " and PXN.cod_vuelo = " & TxtVuelo.Text
End If

If TxtHoraReal.Text <> "" Then
   Con$ = Con$ & " and PXN.hora_vuelo = " & TxtHoraReal.Text
End If

If txtPaxExtranjeros.Text <> "" Then
   Con$ = Con$ & " and PXE.cantidad = " & txtPaxExtranjeros.Text
End If

If txtPaxNacionales.Text <> "" Then
   Con$ = Con$ & " and PXN.cantidad = " & txtPaxNacionales.Text
End If

'If Con$ <> "" Then
'    Con$ = "  " & Mid(Con$, 5, Len(Con$))
'End If

ArmarCondicion = Con$

End Function




Private Sub MePrepararMod()
    
    SeteoBotonesMod False

End Sub

Private Function MeReconsultar() As Integer
Dim sql$
Dim i%
If Aplicacion.ObtenerRsDAO(sqlGral$, rs) Then
        txtCantReg.Text = Aplicacion.CantReg(rs)
        If Val(txtReg.Text) > Val(txtCantReg.Text) Then
            txtReg.Text = txtCantReg.Text
        End If
        
        For i% = 1 To txtReg.Text - 1
            rs.MoveNext
        Next
        If txtCantReg.Text > 0 Then
            MellenarPantalla
        End If
        MeSetearBotonesToolBar
        MeReconsultar = txtCantReg.Text
End If

End Function


Private Sub MeSetearBotonesToolBar()
Dim i%
Dim but As Button

If txtCantReg.Text = 0 Then
'    TollBar.Buttons(1).Enabled = False
'    TollBar.Buttons(2).Enabled = False
'    TollBar.Buttons(3).Enabled = False
'    TollBar.Buttons(4).Enabled = False
'    TollBar.Buttons(6).Enabled = False
'    TollBar.Buttons(7).Enabled = False
ElseIf txtCantReg.Text = 1 Then
    Tollbar.Buttons(4).Enabled = False
    Tollbar.Buttons(5).Enabled = False
    Tollbar.Buttons(6).Enabled = False
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
ElseIf txtReg.Text = txtCantReg.Text Then
    Tollbar.Buttons(4).Enabled = True
    Tollbar.Buttons(5).Enabled = True
    Tollbar.Buttons(6).Enabled = False
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
ElseIf txtReg.Text = 1 Then
    Tollbar.Buttons(4).Enabled = False
    Tollbar.Buttons(5).Enabled = False
    Tollbar.Buttons(6).Enabled = True
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
Else
    Tollbar.Buttons(4).Enabled = True
    Tollbar.Buttons(5).Enabled = True
    Tollbar.Buttons(6).Enabled = True
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
    
End If
    


End Sub



Private Sub SetearBotonesAltas()
    
    Tollbar.Buttons(1).Enabled = True
    Tollbar.Buttons(11).Enabled = True
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Visible = False
    
    
    Tollbar.Buttons(2).Visible = False
    
    Tollbar.Buttons(4).Visible = False
    Tollbar.Buttons(5).Visible = False
    Tollbar.Buttons(6).Visible = False
    Tollbar.Buttons(7).Visible = False
    
    Tollbar.Buttons(9).Visible = False
    Tollbar.Buttons(10).Visible = False
    
    Tollbar.Buttons(15).Visible = False
    Tollbar.Buttons(16).Visible = False
    
    txtCantReg.Visible = False
    txtReg.Visible = False
    de.Visible = False
    Opttipovuelo(2).Visible = False
    MskFecha.Text = Format(Date - 1, FTOFECHA)

End Sub

Private Sub SeteoBotonesMod(valor As Boolean)
    
    
    Tollbar.Buttons(1).Enabled = valor
    Tollbar.Buttons(2).Enabled = valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = valor
    Tollbar.Buttons(6).Enabled = valor
    Tollbar.Buttons(7).Enabled = valor

    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor

    Tollbar.Buttons(12).Enabled = Not valor
    Tollbar.Buttons(13).Enabled = Not valor

    Tollbar.Buttons(15).Enabled = valor
    Tollbar.Buttons(16).Enabled = valor

'habilitar o des frames y/o campos
FraCabecera.Enabled = Not valor

CboCia.Enabled = valor
TxtVuelo.Enabled = valor
MskFecha.Enabled = valor
End Sub





Private Sub CboCia_Change()
chk.Value = True
End Sub

Private Sub cboCia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
ElseIf KeyAscii = 32 Then
    CboCia.ListIndex = -1
End If
End Sub


Private Sub CboCodAeropuerto_Click()
Dim sql As String

sql = " SELECT cod_sdep,descrip FROM baires.subdependencia "
sql = sql & " WHERE cod_depn = '" & CboCodAeropuerto.Text & "'"
sql = sql & " ORDER BY cod_sdep"
 
FuncCbos_LlenarCboLst CboEspigon, LstEspigon, sql

cboLocal.Clear
End Sub

Private Sub CboCodAeropuerto_Change()
chk.Value = True
End Sub


Private Sub CboCodAeropuerto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub CboEspigon_Change()
chk.Value = True
End Sub


Private Sub CboEspigon_Click()
cboLocal.Clear
End Sub


Private Sub CboEspigon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub cboLocal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub CmdFecha_Click()
Dim fch As Date

If MskFecha.Text <> "" Then
    fch = MskFecha.FormattedText
Else
    fch = Date
End If

frmFecha.MuestroFormFecha fch

MskFecha.Text = Format$(fch, FTOFECHA)

MskFecha.SetFocus

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 4 Then
    Select Case KeyCode
        Case 78 'Nueva sel
            Call TollBar_ButtonClick(Tollbar.Buttons(1))
        Case 71 'Guardar
            If Tollbar.Buttons(12).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(12))
            End If
        Case 83 'Salir
            Call TollBar_ButtonClick(Tollbar.Buttons(17))
        Case 66 'Buscar
            If Modo = "MOD" Then
            Call TollBar_ButtonClick(Tollbar.Buttons(2))
            End If
    End Select
    If Modo = "MOD" And Val(txtCantReg.Text) > 0 Then
    Select Case KeyCode
        Case 37 'Izq
            If Tollbar.Buttons(5).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(5))
            End If
        Case 38 'Arriba
            Call TollBar_ButtonClick(Tollbar.Buttons(7))
        Case 40 'Abajo
            Call TollBar_ButtonClick(Tollbar.Buttons(4))
        Case 39 'Der
            If Tollbar.Buttons(6).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(6))
            End If
        Case 82 'Grilla
            If Tollbar.Buttons(16).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(16))
            End If
    
    End Select
    End If
End If

End Sub

Private Sub Form_Load()
Dim sql As String
Top = 1000
Left = 1400
Width = 6200
Height = 6000


Set cl_volados = New CLVolados

TxtPaxTransito.Text = 0
txtTrip.Text = 0
txtPaxNacionales.Text = 0
txtPaxExtranjeros.Text = 0

OptTrip(0).Value = True
'arma el SQL para llenar el combo de Aeropuerto
sql = " SELECT cod_depn,descrip FROM baires.dependencia "
sql = sql & " ORDER BY cod_depn"

FuncCbos_LlenarCbo CboCodAeropuerto, sql

'arma el SQL para llenar el combo de Compania
sql = " SELECT cod_compania,descrip FROM ventas.companias WHERE ACTIVO = 'S' "
sql = sql & " ORDER BY DESCRIP "

FuncCbos_LlenarCboLst CboCia, LstCia, sql

End Sub












Private Sub mskFecha_Change()
'chk.Value = True
End Sub

Private Sub MskFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub mskFecha_LostFocus()

If MskFecha.Text <> "" Then
If Not IsDate(MskFecha.FormattedText) Then
    MskFecha.Text = Format(Date - 1, FTOFECHA)
End If
    
    MskFecha.Text = MskFecha.FormattedText
End If
End Sub


Private Sub Opttipovuelo_Click(Index As Integer)

If Modo = "ALTA" Then
    L_SetearLocal
End If

End Sub

Private Sub Opttipovuelo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub TollBar_ButtonClick(ByVal Button As ComctlLib.Button)
Dim a%
Dim pos As String
Dim saltear As Boolean

saltear = False

pos = txtReg.Text

Select Case Button.Key
    Case "a"
        Func_MoverPrimero rs, pos
    Case "b"
        Func_MoverAnterior rs, pos
    Case "c"
        Func_MoverSiguiente rs, pos
    Case "e"
        Func_MoverUltimo rs, pos
    Case "f"
         saltear = True
         MePrepararMod
    Case "g"
         saltear = True
         MeEliminar
    Case "h"
        saltear = True
        If Modo = "MOD" Then
            MskFecha.Enabled = False
            CboCia.Enabled = False
            TxtVuelo.Enabled = False
            MeActualizar
        Else
            L_AltasDatos
        End If
    Case "i"
        saltear = True
        MeAbortarMod
    Case "j"
        saltear = True
        If chk.Value Then
            If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
                If Modo = "MOD" Then
                    MeActualizar
                Else
                    L_AltasDatos
                End If
            End If
        End If
        Unload Me
    Case "k"
        saltear = True
        NuevaSeleccion
    Case "l"
        saltear = True
        MeCargarDatos
    Case "m"
        saltear = True
        MeImpDatos
    Case "n"
        saltear = True
        L_DatosGrilla
    
End Select

If Not saltear Then
    txtReg.Text = pos
    MellenarPantalla
    MeSetearBotonesToolBar
End If

End Sub




Private Sub TxtHoraReal_Change()
'chk.Value = True
End Sub

Private Sub TxtHoraReal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub




Private Sub TxtPaxExtranjeros_Change()
'chk.Value = True
Label3.caption = Val(txtPaxExtranjeros.Text) + Val(txtPaxNacionales) - Val(txtTrip.Text)
End Sub


Private Sub txtPaxExtranjeros_GotFocus()
    txtPaxExtranjeros.SelStart = 0
    txtPaxExtranjeros.SelLength = Len(txtPaxExtranjeros.Text)
End Sub


Private Sub TxtPaxExtranjeros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub TxtPaxNacionales_Change()
'chk.Value = True
Label3.caption = Val(txtPaxExtranjeros.Text) + Val(txtPaxNacionales) - Val(txtTrip.Text)
End Sub


Private Sub txtPaxNacionales_GotFocus()
    txtPaxNacionales.SelStart = 0
    txtPaxNacionales.SelLength = Len(txtPaxNacionales.Text)
End Sub


Private Sub TxtPaxNacionales_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
'    If Modo = "ALTA" Then
'        Call Tollbar_ButtonClick(Tollbar.Buttons(12))
'        CboCia.SetFocus
'    Else
'        Call Tollbar_ButtonClick(Tollbar.Buttons(2))
'    End If
End If
End Sub


Private Sub TxtPaxTransito_GotFocus()
    TxtPaxTransito.SelStart = 0
    TxtPaxTransito.SelLength = Len(TxtPaxTransito.Text)

End Sub

Private Sub TxtPaxTransito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Modo = "ALTA" Then
    Call TollBar_ButtonClick(Tollbar.Buttons(12))
    CboCia.SetFocus
Else
    Call TollBar_ButtonClick(Tollbar.Buttons(2))
End If
End If
End Sub

Private Sub TxtPaxTransito_LostFocus()

If Not IsNumeric(TxtPaxTransito.Text) Then
   TxtPaxTransito.Text = 0
End If

End Sub


Private Sub txtTrip_Change()
Label3.caption = Val(txtPaxExtranjeros.Text) + Val(txtPaxNacionales) - Val(txtTrip.Text)

End Sub

Private Sub txtTrip_GotFocus()

    txtTrip.SelStart = 0
    txtTrip.SelLength = Len(txtTrip.Text)
 
End Sub

Private Sub txtTrip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub txtTrip_LostFocus()

If Not IsNumeric(txtTrip.Text) Then
   txtTrip.Text = 0
 Else

End If

End Sub


Private Sub TxtVuelo_Change()
'chk.Value = True
End Sub


Private Sub TxtVuelo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub TxtVuelo_LostFocus()
Dim SqlHora As String
Dim RsHora As Recordset
Dim tipo As String

SqlHora = " SELECT hora_vuelo FROM estadis.destinos"
SqlHora = SqlHora & " WHERE cod_cia_aerea = '" & LstCia.List(CboCia.ListIndex) & "'"
If Opttipovuelo(0).Value = True Then
  tipo = "L"
 Else
  tipo = "S"
End If
SqlHora = SqlHora & " AND  tipo = '" & tipo & "'"
SqlHora = SqlHora & " AND cod_vuelo = " & TxtVuelo.Text

If Aplicacion.ObtenerRsDAO(SqlHora, RsHora) Then
    If Aplicacion.CantReg(RsHora) > 0 Then
       TxtHoraReal.Text = RsHora!hora_vuelo
    Else
       TxtHoraReal.Text = ""
    End If
End If

End Sub
