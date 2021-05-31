VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmModifica 
   Caption         =   "Modificación de pasajeros atendidos / volados "
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "N"
            Object.ToolTipText     =   "Nueva Seleción"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "B"
            Object.ToolTipText     =   "Buscar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "G"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "A"
            Object.ToolTipText     =   "Abortar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "P"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "S"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame FraPasajeros 
      Caption         =   "Datos a modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2220
      Left            =   120
      TabIndex        =   7
      Top             =   1845
      Width           =   7260
      Begin VB.ComboBox CboNuevoVuelo 
         Height          =   315
         Left            =   3705
         TabIndex        =   28
         Text            =   "Combo1"
         Top             =   240
         Width           =   1155
      End
      Begin VB.CheckBox ChkNuevoVuelo 
         Height          =   195
         Left            =   4920
         TabIndex        =   27
         Top             =   270
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CheckBox ChkMover 
         Height          =   195
         Left            =   3225
         TabIndex        =   17
         Top             =   1980
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Frame FraOriginal 
         Caption         =   "Datos del vuelo original"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1260
         Left            =   210
         TabIndex        =   16
         Top             =   675
         Width           =   3000
         Begin VB.TextBox TxtVoladosOriginal 
            Height          =   300
            Left            =   1875
            TabIndex        =   24
            Top             =   765
            Width           =   660
         End
         Begin VB.TextBox TxtAtendidosOriginal 
            Height          =   285
            Left            =   1875
            TabIndex        =   23
            Top             =   300
            Width           =   660
         End
         Begin VB.Label LblVoladosOriginal 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Volados :"
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
            Left            =   375
            TabIndex        =   26
            Top             =   765
            Width           =   1380
         End
         Begin VB.Label LblAtendidosOriginal 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Atendidos :"
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
            Left            =   375
            TabIndex        =   25
            Top             =   315
            Width           =   1395
         End
      End
      Begin VB.Frame FraFinal 
         Caption         =   "Datos del vuelo a modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1260
         Left            =   4080
         TabIndex        =   12
         Top             =   675
         Width           =   3000
         Begin MSMask.MaskEdBox txtatendidosfinal 
            Height          =   300
            Left            =   1770
            TabIndex        =   31
            Top             =   315
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   529
            _Version        =   327680
            PromptInclude   =   0   'False
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin VB.TextBox TxtVoladosFinal 
            Height          =   285
            Left            =   1815
            TabIndex        =   13
            Top             =   765
            Width           =   690
         End
         Begin VB.Label LblAtendidosFinal 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Atendidos :"
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
            Left            =   360
            TabIndex        =   15
            Top             =   330
            Width           =   1335
         End
         Begin VB.Label LblVoladosFinal 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Volados :"
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
            Left            =   360
            TabIndex        =   14
            Top             =   765
            Width           =   1350
         End
      End
      Begin VB.CommandButton CmdUnoAdelante 
         Caption         =   ">"
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
         Left            =   3450
         Picture         =   "FrmModifica.frx":0000
         TabIndex        =   11
         Top             =   690
         Width           =   375
      End
      Begin VB.CommandButton CmdTodosAdelante 
         Caption         =   ">>"
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
         Left            =   3450
         Picture         =   "FrmModifica.frx":0442
         TabIndex        =   10
         Top             =   1050
         Width           =   375
      End
      Begin VB.CommandButton CmdAtrasUno 
         Caption         =   "<"
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
         Left            =   3450
         TabIndex        =   9
         Top             =   1410
         Width           =   375
      End
      Begin VB.CommandButton CmdAtrasTodos 
         Caption         =   "<<"
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
         Left            =   3465
         TabIndex        =   8
         Top             =   1785
         Width           =   375
      End
      Begin VB.Label LblNuevoVuelo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nuevo vuelo :"
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
         Left            =   2130
         TabIndex        =   29
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame FraCabecera 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7275
      Begin VB.TextBox TxtCiaAerea 
         Height          =   285
         Left            =   3840
         TabIndex        =   30
         Top             =   270
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox TxtCompania 
         Height          =   285
         Left            =   1725
         TabIndex        =   20
         Top             =   1005
         Width           =   1830
      End
      Begin VB.TextBox TxtVueloOriginal 
         Height          =   285
         Left            =   5190
         TabIndex        =   19
         Top             =   1020
         Width           =   1125
      End
      Begin VB.TextBox TxtEspigon 
         Height          =   285
         Left            =   5205
         TabIndex        =   6
         Top             =   615
         Width           =   1830
      End
      Begin MSMask.MaskEdBox MskFecha 
         Height          =   285
         Left            =   1740
         TabIndex        =   4
         Top             =   210
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Format          =   "dd-mm-yyyy"
         Mask            =   "##-##-####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtAeropuerto 
         Height          =   285
         Left            =   1740
         TabIndex        =   3
         Top             =   615
         Width           =   1830
      End
      Begin VB.Label LblCompania 
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
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   1005
         Width           =   1380
      End
      Begin VB.Label LblVueloOriginal 
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
         Left            =   3720
         TabIndex        =   21
         Top             =   1005
         Width           =   1365
      End
      Begin VB.Label LblEspigon 
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
         Height          =   285
         Left            =   3720
         TabIndex        =   5
         Top             =   615
         Width           =   1350
      End
      Begin VB.Label LblAeropuerto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aeropuerto :"
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
         Left            =   255
         TabIndex        =   2
         Top             =   615
         Width           =   1365
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
         Left            =   270
         TabIndex        =   1
         Top             =   225
         Width           =   1350
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   540
      Top             =   345
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmModifica.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmModifica.frx":0B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmModifica.frx":0EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmModifica.frx":11D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmModifica.frx":12E4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmModifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cl_vs As CLVersus
Dim NoPasarOriginal As Integer
Dim NoPasarFinal As Integer
Dim NoPasarTotal As Integer
Dim GuardarOriginal As Integer


Public Sub L_AltasDatos()

End Sub


Public Sub MeAbortarMod()

  If ChkMover.Value = 1 Then
    If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
       MeActualizar
     Else
       Unload Me
       'TxtAtendidosOriginal.Text = GuardarOriginal
    End If
  End If

End Sub


Public Sub MeActualizar()

Aplicacion.SeteoProceso ("Actualizando")

Aplicacion.ComienzoTrans

MeLlenarObjeto


If cl_vs.update_pax_at_vs_vol() Then
    Aplicacion.TerminarConExitoTrans
    ChkMover.Value = 0
    'SeteoBotonesMod True
 Else
    Aplicacion.TerminarConErrorTrans
End If

Aplicacion.SeteoFin

End Sub


Public Sub MeCargarDatos()
 Dim sql As String
 Dim RS As Recordset
  
'Preparo la barra de tareas
SeteoBotonesBuscar
 
 sql = " SELECT pax_at,pax_vol FROM estadis.pax_at_vs_vol "
 sql = sql & " WHERE cod_cia_aerea = " & cl_vs.cod_cia_aerea
 sql = sql & " AND cod_vuelo = " & CboNuevoVuelo.Text
 sql = sql & " AND cod_depn = '" & cl_vs.cod_depn & "'"
 sql = sql & " AND cod_sdep = '" & cl_vs.cod_sdep & "'"
 sql = sql & " AND fch_vuelo = " & func_ToDate(cl_vs.fch_vuelo)
  
If Aplicacion.ObtenerRsDAO(sql, RS) Then
    If Aplicacion.CantReg(RS) > 0 Then
       RS.MoveFirst
       txtatendidosfinal.Text = str(RS!pax_at)
       TxtVoladosFinal.Text = RS!pax_vol
       NoPasarOriginal = TxtAtendidosOriginal
       NoPasarFinal = txtatendidosfinal
       NoPasarTotal = NoPasarOriginal + NoPasarFinal
       GuardarOriginal = NoPasarOriginal
       'guardarFinal = NoPasarFinal
      Else
       txtatendidosfinal.Text = 0
       TxtVoladosFinal.Text = 0
       NoPasarOriginal = TxtAtendidosOriginal
       NoPasarFinal = txtatendidosfinal
    End If
       SetearBotones
       Aplicacion.CerrarDAO RS
  Else
End If

ChkNuevoVuelo.Value = 1

End Sub

Public Sub MeLlenarObjeto()

cl_vs.fch_vuelo = MskFecha.FormattedText
cl_vs.cod_depn = TxtAeropuerto.Text
cl_vs.cod_sdep = TxtEspigon.Text
cl_vs.cod_cia_aerea = TxtCiaAerea.Text
cl_vs.cod_vuelo = TxtVueloOriginal.Text
cl_vs.pax_at = TxtAtendidosOriginal.Text

cl_vs.cod_nuevo_vuelo = CboNuevoVuelo.Text
cl_vs.nuevo_pax_at = txtatendidosfinal.Text
End Sub

Public Sub MePrepararMod()

End Sub


Public Sub NuevaSeleccion()
Dim i%

If ChkMover.Value = 1 Then
  If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
     MeActualizar
   Else
    TxtAtendidosOriginal.Text = GuardarOriginal
  End If
End If

'Limpiar campos de pantallas
CboNuevoVuelo.Enabled = True
txtatendidosfinal.Enabled = True

CboNuevoVuelo.ListIndex = -1


txtatendidosfinal.Text = 0
TxtVoladosFinal.Text = 0

'Set cl_vs = New CLVersus

ChkMover.Value = False
SetBotonesGeneral False
End Sub

Public Sub SeteoBotonesBuscar()
    
    
    Tollbar.Buttons(1).Enabled = Not valor
    Tollbar.Buttons(2).Enabled = valor
    
    Tollbar.Buttons(4).Enabled = Not valor
    Tollbar.Buttons(5).Enabled = Not valor

    Tollbar.Buttons(7).Enabled = valor
    Tollbar.Buttons(8).Enabled = Not valor

'habilitar o des frames y/o campos
FraOriginal.Enabled = Not valor
FraFinal.Enabled = Not valor

'habilitar cajas de texto

TxtAtendidosOriginal.Enabled = valor
TxtVoladosOriginal.Enabled = valor
txtatendidosfinal.Enabled = Not valor
TxtVoladosFinal.Enabled = valor

' deshabilitar combo NuevoVuelo

CboNuevoVuelo.Enabled = valor


SetearBotones

End Sub

Private Sub SeteoBotonesMod(valor As Boolean)
    
    Tollbar.Buttons(1).Enabled = valor
    Tollbar.Buttons(2).Enabled = Not valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = valor

    Tollbar.Buttons(7).Enabled = valor
    Tollbar.Buttons(8).Enabled = Not valor

'habilitar o des frames y/o campos
FraOriginal.Enabled = valor
FraFinal.Enabled = Not valor

'habilitar cajas de texto

TxtAtendidosOriginal.Enabled = valor
TxtVoladosOriginal.Enabled = valor
txtatendidosfinal.Enabled = Not valor
TxtVoladosFinal.Enabled = valor

MousePointer = 0

End Sub
 
Private Sub SetBotonesGeneral(valor As Boolean)
    
    Tollbar.Buttons(1).Enabled = valor 'Nueva Seleccion
    Tollbar.Buttons(2).Enabled = valor     'Buscar
    
    Tollbar.Buttons(4).Enabled = valor  'Grabar
    Tollbar.Buttons(5).Enabled = valor  'Abortar
    
    Tollbar.Buttons(7).Enabled = valor  'Imprimir
    Tollbar.Buttons(8).Enabled = Not valor  'Salir

   
'Seteo de botones inicial

CmdAtrasUno.Enabled = valor
CmdAtrasTodos.Enabled = valor
CmdUnoAdelante.Enabled = valor
CmdTodosAdelante.Enabled = valor

'Seteo de frames inicial

FraPasajeros.Enabled = Not valor
FraCabecera.Enabled = valor

'Seteo de cajas de texto

FraOriginal.Enabled = valor
FraFinal.Enabled = valor

'Seteo de combo nuevo vuelo

CboNuevoVuelo.Enabled = Not valor


'Seteo el chk a verdadero para una nueva seleccion
ChkMover.Value = 0
End Sub

Public Sub cargar(cl_versus As CLVersus)
 Set cl_vs = cl_versus
 Show 1
End Sub




Public Sub SetearBotones()

Select Case Val(TxtAtendidosOriginal.Text)
   Case Is > 0
        CmdUnoAdelante.Enabled = True
        CmdTodosAdelante.Enabled = True
   Case Is = 0
        CmdUnoAdelante.Enabled = False
        CmdTodosAdelante.Enabled = False
End Select


Select Case Val(txtatendidosfinal.Text)
   Case Is > 0
        CmdAtrasUno.Enabled = True
        CmdAtrasTodos.Enabled = True
   Case Is = 0
        CmdAtrasUno.Enabled = False
        CmdAtrasTodos.Enabled = False
End Select

End Sub

Private Sub CboNuevoVuelo_Click()
    SeteoBotonesMod False
End Sub

Private Sub CmdAtrasTodos_Click()

TxtAtendidosOriginal.Text = Val(TxtAtendidosOriginal.Text) + Val(txtatendidosfinal.Text)
txtatendidosfinal.Text = 0

NoPasarFinal = Val(txtatendidosfinal)
NoPasarOriginal = Val(TxtAtendidosOriginal)

SetearBotones

ChkMover.Value = 1

End Sub

Private Sub CmdAtrasUno_Click()

TxtAtendidosOriginal.Text = Val(TxtAtendidosOriginal.Text) + 1
txtatendidosfinal.Text = Val(txtatendidosfinal.Text) - 1

NoPasarFinal = Val(txtatendidosfinal)
NoPasarOriginal = Val(TxtAtendidosOriginal)

SetearBotones

ChkMover.Value = 1

End Sub



Private Sub CmdTodosAdelante_Click()

txtatendidosfinal.Text = Val(TxtAtendidosOriginal.Text) + Val(txtatendidosfinal.Text)
TxtAtendidosOriginal.Text = 0

NoPasarFinal = Val(txtatendidosfinal)
NoPasarOriginal = Val(TxtAtendidosOriginal)

SetearBotones
ChkMover.Value = 1
End Sub

Private Sub CmdUnoAdelante_Click()

TxtAtendidosOriginal.Text = Val(TxtAtendidosOriginal.Text) - 1
txtatendidosfinal.Text = Val(txtatendidosfinal.Text) + 1

NoPasarFinal = Val(txtatendidosfinal)
NoPasarOriginal = Val(TxtAtendidosOriginal)

SetearBotones

ChkMover.Value = 1
End Sub

Private Sub Form_Load()
Dim sql As String

SetBotonesGeneral False

'Llenar combo de vuelos

Top = 1000
Left = 500

sql = ""
sql = " SELECT cod_vuelo "
sql = sql & " FROM estadis.pax_at_vs_vol "
sql = sql & " WHERE cod_cia_aerea = " & cl_vs.cod_cia_aerea
sql = sql & " AND cod_depn = '" & cl_vs.cod_depn & "'"
sql = sql & " AND cod_sdep = '" & cl_vs.cod_sdep & "'"
sql = sql & " AND cod_vuelo <> " & cl_vs.cod_vuelo
sql = sql & " ORDER BY cod_vuelo "

FuncCbos_LlenarCbo CboNuevoVuelo, sql



MskFecha.Text = cl_vs.fch_vuelo
TxtAeropuerto.Text = cl_vs.cod_depn
TxtEspigon.Text = cl_vs.cod_sdep
TxtCiaAerea.Text = cl_vs.cod_cia_aerea
TxtCompania.Text = cl_vs.des_cia
TxtVueloOriginal.Text = cl_vs.cod_vuelo
TxtAtendidosOriginal.Text = cl_vs.pax_at
TxtVoladosOriginal.Text = cl_vs.pax_vol


End Sub


Private Sub TollBar_ButtonClick(ByVal Button As ComctlLib.Button)

Select Case Button.Key
    Case "G" 'Grabar
        MeActualizar
    Case "A"  'Abortar
        If ChkMover.Value = 1 Then
          If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
             MeActualizar
          End If
        End If
        Unload Me
    Case "S"  'Salir
        If ChkMover.Value = 1 Then
            If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
               MeActualizar
            End If
        End If
        Unload Me
        'FrmVersus.MeLlenargrilla
    Case "N"  'Nueva Seleccion
        NuevaSeleccion
    Case "B" 'Buscar
        MeCargarDatos
    Case "P"  'Imprimir
End Select

End Sub


Private Sub TxtAtendidosFinal_KeyPress(KeyAscii As Integer)
Dim diferencia As Integer
Dim tecla As Integer

diferencia = 0

tecla = KeyAscii


If tecla = 13 Then
  
  Select Case txtatendidosfinal
     
     Case Is = ""
       txtatendidosfinal = 0
       diferencia = NoPasarFinal - txtatendidosfinal
       TxtAtendidosOriginal = TxtAtendidosOriginal + diferencia
     Case Is = NoPasarFinal
       diferencia = NoPasarFinal - txtatendidosfinal
       TxtAtendidosOriginal = TxtAtendidosOriginal + diferencia
     Case Is > NoPasarFinal
         If txtatendidosfinal <= NoPasarTotal Then
            diferencia = txtatendidosfinal - NoPasarFinal
            TxtAtendidosOriginal = TxtAtendidosOriginal - diferencia
          Else
            MsgBox ("Cantidad de pasajeros no valida")
            txtatendidosfinal = NoPasarFinal
        End If
     Case Is < NoPasarFinal
        If txtatendidosfinal >= 0 Then
           diferencia = NoPasarFinal - txtatendidosfinal
           TxtAtendidosOriginal = TxtAtendidosOriginal + diferencia
         Else
           MsgBox ("Cantidad de pasajeros no valida")
           txtatendidosfinal = NoPasarFinal
        End If

  End Select
  SetearBotones
  NoPasarFinal = Val(txtatendidosfinal)
  NoPasarOriginal = Val(TxtAtendidosOriginal)
  ChkMover.Value = 1
 Else
End If

End Sub


Private Sub TxtAtendidosFinal_LostFocus()
  
  If Not IsNumeric(txtatendidosfinal.Text) Then
    txtatendidosfinal.Text = NoPasarFinal
  End If
  
End Sub


