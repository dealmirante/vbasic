VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmAdmModEsp 
   Caption         =   "Administración de Modelos de Estimados"
   ClientHeight    =   6075
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   5655
   Icon            =   "FrmAdmModEsp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6075
   ScaleWidth      =   5655
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   105
      TabIndex        =   9
      Top             =   15
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   16
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
            Key             =   "d"
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
            Key             =   "j"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame frEsp 
      Caption         =   "Porc por Espigón"
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
      Height          =   3000
      Left            =   60
      TabIndex        =   12
      Top             =   3045
      Width           =   5400
      Begin FPSpread.vaSpread sprPorc 
         Height          =   2715
         Left            =   120
         OleObjectBlob   =   "FrmAdmModEsp.frx":0442
         TabIndex        =   6
         Top             =   210
         Width           =   5055
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "chk"
      Height          =   195
      Left            =   4155
      TabIndex        =   10
      Top             =   2610
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3495
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox txtCantReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4380
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   510
      Width           =   480
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
      Height          =   2550
      Left            =   60
      TabIndex        =   11
      Top             =   480
      Width           =   5385
      Begin VB.Frame frMod 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   405
         TabIndex        =   16
         Top             =   270
         Width           =   3945
         Begin VB.TextBox txtNom 
            Height          =   300
            Left            =   1410
            MaxLength       =   30
            TabIndex        =   2
            Top             =   705
            Width           =   2175
         End
         Begin MSMask.MaskEdBox mskAnio 
            Height          =   300
            Left            =   1425
            TabIndex        =   0
            Top             =   0
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
            Left            =   1425
            TabIndex        =   1
            Top             =   345
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   529
            _Version        =   327680
            PromptInclude   =   0   'False
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre"
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
            Left            =   105
            TabIndex        =   19
            Top             =   720
            Width           =   1185
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
            Left            =   120
            TabIndex        =   18
            Top             =   15
            Width           =   1185
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
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1185
         End
      End
      Begin MSMask.MaskEdBox mskImp 
         Height          =   300
         Left            =   1815
         TabIndex        =   3
         Top             =   1350
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   12
         Format          =   "$ #,###.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskTick 
         Height          =   300
         Left            =   1815
         TabIndex        =   4
         Top             =   1740
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   8
         Format          =   "#,###"
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskPax 
         Height          =   285
         Left            =   1815
         TabIndex        =   5
         Top             =   2115
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   8
         Format          =   "#,###"
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pax Viajados"
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
         Left            =   510
         TabIndex        =   20
         Top             =   2115
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tickets"
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
         Index           =   6
         Left            =   510
         TabIndex        =   15
         Top             =   1755
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe"
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
         Left            =   510
         TabIndex        =   14
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Label de 
         Caption         =   "de"
         Height          =   255
         Left            =   4050
         TabIndex        =   13
         Top             =   135
         Width           =   195
      End
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
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmModEsp.frx":0F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmModEsp.frx":12A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmModEsp.frx":15BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmModEsp.frx":18D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmModEsp.frx":1BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmModEsp.frx":1F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmModEsp.frx":2222
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmModEsp.frx":253C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmModEsp.frx":2856
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmModEsp.frx":2B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmModEsp.frx":2C82
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAdmModEsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As Recordset

Dim cl_Estim As CLEstimado

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$

Public Sub Altas()
If Aplicacion.Nivel = 0 Then
    Modo = "ALTA"
    SetearBotonesAltas
    Me.Show 1
Else
    MsgBox "No tiene autorización para Crear modelos Iniciales", vbOKOnly + vbExclamation, "ATENCION"
End If
End Sub

Private Sub L_AltasDatos()

If L_TodoCargado Then
    FrmAdmModEsp.caption = Aplicacion.SeteoProceso(FrmAdmModEsp.caption)

    Aplicacion.ComienzoTrans

    MeLlenarObjeto

    If cl_Estim.Insert_Estim() Then
        Aplicacion.TerminarConExitoTrans
        chk.Value = False
    
        NuevaSeleccion
    
    Else
        Aplicacion.TerminarConErrorTrans
    End If

    FrmAdmModEsp.caption = Aplicacion.SeteoFin

Else
'    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

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

Public Sub Modificacion()
Modo = "MOD"
Me.Show 1
End Sub
Private Sub NuevaSeleccion()
Dim i, j

If Modo = "MOD" Then
    SetBotonesGeneral False
    mskAnio.Text = ""
    mskMes.Text = ""
    mskImp.Text = ""
    mskPax.Text = ""
    mskTick.Text = ""
    txtNom.Text = ""
Else
    If chk.Value Then
        If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
            L_AltasDatos
        End If
    End If

    mskImp.Text = 0
    mskPax.Text = 0
    mskTick.Text = 0
    txtNom.Text = L_NombreSujerido
End If
'Limpiar campos de pantallas
Set cl_Estim = New CLEstimado
    
    For i = 1 To 9
        For j = 2 To 4
            sprPorc.SetText j, i, ""
        Next
    Next

mskAnio.SetFocus
chk.Value = 0

End Sub

Private Sub MeAbortarMod()
    
    SeteoBotonesMod True
    
    Tollbar.Buttons(2).Enabled = False
    
    MeSetearBotonesToolBar
    
    MellenarPantalla
    
End Sub

Private Sub MeActualizar()

Dim ViejoOrgan$
Dim Viejocargo%

If L_TodoCargado Then

Aplicacion.SeteoProceso ("Actualizando")

Aplicacion.ComienzoTrans

MeLlenarObjeto

If cl_Estim.Update_Estim Then '
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
Dim sql$

FrmAdmModEsp.caption = Aplicacion.SeteoProceso(FrmAdmModEsp.caption)
        
CondConsulta = ArmarCondicion

sqlGral$ = ""
sqlGral$ = sqlGral$ & " SELECT "
sqlGral$ = sqlGral$ & " descrip,"
sqlGral$ = sqlGral$ & " anio,"
sqlGral$ = sqlGral$ & " mes,"
sqlGral$ = sqlGral$ & " importe, "
sqlGral$ = sqlGral$ & " ticket, "
sqlGral$ = sqlGral$ & " pax, "
sqlGral$ = sqlGral$ & " usuario "
sqlGral$ = sqlGral$ & " from estadis.modelo_estim "
sqlGral$ = sqlGral$ & CondConsulta

If Aplicacion.ObtenerRsDAO(sqlGral$, rs) Then
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

FrmAdmModEsp.caption = Aplicacion.SeteoFin

End Sub
Private Sub MeEliminar()

If Aplicacion.username = rs!usuario Then

If MsgBox("Serán eliminados todos los Modelos para el Año-Mes ", vbInformation + vbOKCancel, "ATENCION") = vbOK Then
    FrmAdmModEsp.caption = Aplicacion.SeteoProceso(FrmAdmModEsp.caption)
    
    Aplicacion.ComienzoTrans
    
    If cl_Estim.Delete_Estim(mskAnio.Text, mskMes.Text) Then
        Aplicacion.TerminarConExitoTrans
        If MeReconsultar > 0 Then
            'Tollbar.Buttons(2).Enabled = False
            'MeSetearBotonesToolBar
        Else
            NuevaSeleccion
        End If
    Else
        Aplicacion.TerminarConErrorTrans
    End If
    
    FrmAdmModEsp.caption = Aplicacion.SeteoFin
    End If
Else
    MsgBox "No puede Eliminar el modelo de " & rs!usuario, vbCritical + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub MeLlenarObjeto()
Dim cl_esp As CLGeneric
Dim i, j
Dim valor As Variant

cl_Estim.anio = mskAnio.Text
cl_Estim.Mes = mskMes.Text
cl_Estim.Descrip = txtNom.Text
cl_Estim.Importe = mskImp.Text
cl_Estim.ticket = mskTick.Text
cl_Estim.Pax = mskPax.Text

Set cl_Estim.col_PorcEsp = New Collection

For i = 1 To 9
    
    For j = 2 To 4
        Set cl_esp = New CLGeneric
        Select Case i
            
            Case 1
                cl_esp.depn = "EZE"
                cl_esp.Identif = "INTAL"
            Case 2
                cl_esp.depn = "EZE"
                cl_esp.Identif = "INTAS"
            Case 3
                cl_esp.depn = "EZE"
                cl_esp.Identif = "INTB"
            Case 4
                cl_esp.depn = "AEP"
                cl_esp.Identif = "AEP"
            Case 5
                cl_esp.depn = "INT"
                cl_esp.Identif = "BARI"
            Case 6
                cl_esp.depn = "INT"
                cl_esp.Identif = "CORD"
            Case 7
                cl_esp.depn = "INT"
                cl_esp.Identif = "IGUA"
            Case 8
                cl_esp.depn = "INT"
                cl_esp.Identif = "MDPL"
            Case 9
                cl_esp.depn = "INT"
                cl_esp.Identif = "MEND"
                
        End Select
        Select Case j
            Case 2
                cl_esp.tipo = "I"
            Case 3
                cl_esp.tipo = "T"
            Case 4
                cl_esp.tipo = "P"
            Case 5
                cl_esp.tipo = "V"
        End Select
    
        sprPorc.GetText j, i, valor
        cl_esp.Porc = valor
    
        cl_Estim.col_PorcEsp.Add cl_esp
    Next
Next

End Sub


Private Function L_TodoCargado() As Boolean
Dim salida As Boolean
Dim i
Dim valor As Variant


salida = True

If mskImp.Text > 0 And mskTick.Text > 0 And mskPax.Text > 0 Then

        For i = 2 To 4
        sprPorc.GetText i, 10, valor
        If Val(valor) <> 0 And Val(valor) <> 100 Then
            salida = False
            MsgBox "Algún % no suma 100", vbExclamation + vbOKOnly, "ATENCION"
            Exit For
        End If
    Next
Else
   MsgBox "El código y/o el Importe no son correcto", vbExclamation + vbOKOnly, "ATENCION"
   salida = False
End If

L_TodoCargado = salida

End Function

Private Sub MellenarPantalla()
Dim rsPorc As Recordset
Dim sql As String
Dim i As Integer
Dim j As Integer

'Borro los campos anteriores de la grilla por si
'en el nuevo registro que pongo falta el interior

For i = 1 To 9
  For j = 2 To 4
      sprPorc.SetText j, i, ""
  Next
Next


mskAnio.Text = rs!anio
mskMes.Text = rs!Mes
mskImp.Text = rs!Importe
mskTick.Text = rs!ticket
mskPax.Text = rs!Pax
txtNom.Text = rs!Descrip

sql = sql & " SELECT "
sql = sql & " anio,"
sql = sql & " mes,"
sql = sql & " cod_depn,"
sql = sql & " cod_sdep,"
sql = sql & " tipo_porc,"
sql = sql & " porcentaje"
sql = sql & " FROM estadis.porciento_espigon "
sql = sql & " WHERE anio = " & rs!anio
sql = sql & " and mes = " & rs!Mes
sql = sql & " ORDER BY cod_depn,cod_sdep"

If Aplicacion.ObtenerRsDAO(sql, rsPorc) Then
    If Aplicacion.CantReg(rsPorc) > 0 Then
        Do While Not rsPorc.EOF
            Select Case rsPorc!cod_sdep
                Case "INTAL"
                    Select Case rsPorc!tipo_porc
                        Case "I"
                            sprPorc.SetText 2, 1, str(rsPorc!porcentaje)
                        Case "T"
                            sprPorc.SetText 3, 1, str(rsPorc!porcentaje)
                        Case "P"
                            sprPorc.SetText 4, 1, str(rsPorc!porcentaje)
                    End Select
                Case "INTAS"
                    Select Case rsPorc!tipo_porc
                        Case "I"
                            sprPorc.SetText 2, 2, str(rsPorc!porcentaje)
                        Case "T"
                            sprPorc.SetText 3, 2, str(rsPorc!porcentaje)
                        Case "P"
                            sprPorc.SetText 4, 2, str(rsPorc!porcentaje)
                    End Select
                
                Case "INTB"
                    Select Case rsPorc!tipo_porc
                        Case "I"
                            sprPorc.SetText 2, 3, str(rsPorc!porcentaje)
                        Case "T"
                            sprPorc.SetText 3, 3, str(rsPorc!porcentaje)
                        Case "P"
                            sprPorc.SetText 4, 3, str(rsPorc!porcentaje)
                    End Select
                Case "AEP"
                    Select Case rsPorc!tipo_porc
                        Case "I"
                            sprPorc.SetText 2, 4, str(rsPorc!porcentaje)
                        Case "T"
                            sprPorc.SetText 3, 4, str(rsPorc!porcentaje)
                        Case "P"
                            sprPorc.SetText 4, 4, str(rsPorc!porcentaje)
                    End Select
                Case "BARI"
                     Select Case rsPorc!tipo_porc
                        Case "I"
                            sprPorc.SetText 2, 5, str(rsPorc!porcentaje)
                        Case "T"
                            sprPorc.SetText 3, 5, str(rsPorc!porcentaje)
                        Case "P"
                            sprPorc.SetText 4, 5, str(rsPorc!porcentaje)
                    End Select
                Case "CORD"
                     Select Case rsPorc!tipo_porc
                        Case "I"
                            sprPorc.SetText 2, 6, str(rsPorc!porcentaje)
                        Case "T"
                            sprPorc.SetText 3, 6, str(rsPorc!porcentaje)
                        Case "P"
                            sprPorc.SetText 4, 6, str(rsPorc!porcentaje)
                    End Select
                Case "IGUA"
                     Select Case rsPorc!tipo_porc
                        Case "I"
                            sprPorc.SetText 2, 7, str(rsPorc!porcentaje)
                        Case "T"
                            sprPorc.SetText 3, 7, str(rsPorc!porcentaje)
                        Case "P"
                            sprPorc.SetText 4, 7, str(rsPorc!porcentaje)
                    End Select
                Case "MDPL"
                     Select Case rsPorc!tipo_porc
                        Case "I"
                            sprPorc.SetText 2, 8, str(rsPorc!porcentaje)
                        Case "T"
                            sprPorc.SetText 3, 8, str(rsPorc!porcentaje)
                        Case "P"
                            sprPorc.SetText 4, 8, str(rsPorc!porcentaje)
                    End Select
                Case "MEND"
                     Select Case rsPorc!tipo_porc
                        Case "I"
                            sprPorc.SetText 2, 9, str(rsPorc!porcentaje)
                        Case "T"
                            sprPorc.SetText 3, 9, str(rsPorc!porcentaje)
                        Case "P"
                            sprPorc.SetText 4, 9, str(rsPorc!porcentaje)
                    End Select
                
            End Select
            rsPorc.MoveNext
        Loop
    End If
    Aplicacion.CerrarDAO rsPorc
End If
End Sub
Private Function L_NombreSujerido() As String

L_NombreSujerido = mskMes.Text & "-" & mskAnio.Text

End Function


Private Sub SetBotonesGeneral(valor As Boolean)
    
    Tollbar.Buttons(1).Enabled = valor
    Tollbar.Buttons(2).Enabled = Not valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = valor
    Tollbar.Buttons(6).Enabled = valor
    Tollbar.Buttons(7).Enabled = valor

    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor
'
'    TollBar.Buttons(12).Enabled = Not valor
'    TollBar.Buttons(13).Enabled = Not valor
'

'habilitar frames
    frCab.Enabled = Not valor

    
    If Not valor Then
        txtReg.Text = 0
        txtCantReg.Text = 0
    End If

End Sub

Private Function ArmarCondicion()
Dim Con$

Con$ = ""

If mskAnio.Text <> "" Then
    Con$ = Con$ & " and Anio = " & mskAnio.Text
End If

If mskMes.Text <> "" Then
    Con$ = Con$ & " and Mes = " & mskMes.Text
End If

If txtNom.Text <> "" Then
    Con$ = Con$ & " and descrip like '" & txtNom.Text & "%'"
End If

If Con$ <> "" Then
    Con$ = " WHERE " & Mid(Con$, 5, Len(Con$))
End If

ArmarCondicion = Con$

End Function



Private Sub MePrepararMod()
    
If Aplicacion.username = rs!usuario Then
    SeteoBotonesMod False
Else
    MsgBox "No puede Modificar el modelo de " & rs!usuario, vbCritical + vbOKOnly, "ATENCION"
End If
End Sub

Private Function MeReconsultar() As Integer
Dim sql$
Dim i%
    
'frm_.caption = Aplicacion.SeteoProceso (frm_.caption)
    

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
        'MeSetearBotonesToolBar
        MeReconsultar = txtCantReg.Text
End If

'frm_.caption = Aplicacion.SeteoProceso (frm_.caption)

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
        
    txtCantReg.Visible = False
    txtReg.Visible = False
    de.Visible = False
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

'habilitar o des frames y/o campos
    frEsp.Enabled = Not valor
    frCab.Enabled = Not valor
    frMod.Enabled = valor
End Sub





Private Sub Form_Load()
Top = 1000
Left = 800
Width = 5700
Height = 7000

Set cl_Estim = New CLEstimado

If Modo = "MOD" Then
    frEsp.Enabled = False
Else
    mskAnio.Text = Year(Date)
    mskMes.Text = Month(Date)
    
    mskImp.Text = 0
    mskPax.Text = 0
    mskTick.Text = 0
    
    txtNom.Text = L_NombreSujerido
    chk.Value = 0
    
End If

End Sub







Private Sub mskAnio_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub mskAnio_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 45 Then
    If Modo = "ALTA" Then
        Call TollBar_ButtonClick(Tollbar.Buttons(12))
        mskAnio.SetFocus
    Else
        Call TollBar_ButtonClick(Tollbar.Buttons(2))
    End If

 End If
End Sub

Private Sub mskAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub mskAnio_LostFocus()
If Modo = "ALTA" Then
    If Val(mskAnio.Text) < 1996 Or Val(mskAnio) > 2050 Then
        mskAnio.Text = Year(Date)
    End If
End If
txtNom.Text = L_NombreSujerido
End Sub


Private Sub mskImp_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub mskImp_GotFocus()
'mskImp.SelText
End Sub


Private Sub mskImp_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 45 Then
    If Modo = "ALTA" Then
        Call TollBar_ButtonClick(Tollbar.Buttons(12))
        mskAnio.SetFocus
    Else
        Call TollBar_ButtonClick(Tollbar.Buttons(2))
    End If

 End If
End Sub

Private Sub mskImp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub mskImp_LostFocus()

If Not IsNumeric(mskImp.Text) Then
    mskImp.Text = 0
End If

End Sub


Private Sub mskMes_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub mskMes_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 45 Then
    If Modo = "ALTA" Then
        Call TollBar_ButtonClick(Tollbar.Buttons(12))
        mskAnio.SetFocus
    Else
        Call TollBar_ButtonClick(Tollbar.Buttons(2))
    End If

 End If
End Sub


Private Sub mskMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub mskMes_LostFocus()
If Modo = "ALTA" Then
    If Val(mskMes.Text) < 1 Or Val(mskMes.Text) > 12 Then
        mskMes.Text = Month(Date)
    End If
End If
txtNom.Text = L_NombreSujerido
End Sub



Private Sub mskPax_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub mskPax_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 45 Then
    If Modo = "ALTA" Then
        Call TollBar_ButtonClick(Tollbar.Buttons(12))
        mskAnio.SetFocus
    Else
        Call TollBar_ButtonClick(Tollbar.Buttons(2))
    End If

 End If
End Sub


Private Sub mskPax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub mskPax_LostFocus()
If Not IsNumeric(mskPax.Text) Then
    mskPax.Text = 0
End If
End Sub

Private Sub mskTick_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub

Private Sub mskTick_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 45 Then
    If Modo = "ALTA" Then
        Call TollBar_ButtonClick(Tollbar.Buttons(12))
        mskAnio.SetFocus
    Else
        Call TollBar_ButtonClick(Tollbar.Buttons(2))
    End If

 End If
End Sub


Private Sub mskTick_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub mskTick_LostFocus()

If Not IsNumeric(mskTick.Text) Then
    mskTick.Text = 0
End If

End Sub




Private Sub sprPorc_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 45 Then
    If Modo = "ALTA" Then
        Call TollBar_ButtonClick(Tollbar.Buttons(12))
        mskAnio.SetFocus
    Else
        Call TollBar_ButtonClick(Tollbar.Buttons(2))
    End If

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
    Case "d"
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
    
End Select

If Not saltear Then
    txtReg.Text = pos
    MellenarPantalla
    MeSetearBotonesToolBar
End If

End Sub





Private Sub txtNom_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub


Private Sub txtNom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


