VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form FrmAdmCajero 
   Caption         =   "Administración de Cajeros"
   ClientHeight    =   2415
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   7470
   Icon            =   "FrmAdmCajero.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2415
   ScaleWidth      =   7470
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   15
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   20
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "o"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "p"
            Object.ToolTipText     =   "Estado Mod / Consulta"
            Object.Tag             =   ""
            ImageIndex      =   14
            Value           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "k"
            Object.ToolTipText     =   "Nueva Seleción"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "l"
            Object.ToolTipText     =   "Buscar"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "a"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "b"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "c"
            Object.ToolTipText     =   "Registro Siguiente"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "e"
            Object.ToolTipText     =   "Ultimo Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "f"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "g"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "h"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "i"
            Object.ToolTipText     =   "Abortar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "m"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "Grilla"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "j"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CheckBox chk 
      Caption         =   "chk"
      Height          =   195
      Left            =   1380
      TabIndex        =   5
      Top             =   450
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4185
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   465
   End
   Begin VB.TextBox txtCantReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5100
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   420
      Width           =   480
   End
   Begin VB.Frame frCab 
      Height          =   1590
      Left            =   195
      TabIndex        =   2
      Top             =   480
      Width           =   6945
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   405
         Width           =   3360
      End
      Begin VB.CommandButton botHelpCaj 
         Height          =   285
         Left            =   6255
         Picture         =   "FrmAdmCajero.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   420
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskLegajo 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   405
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskCajero 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   855
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Legajo"
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
         Left            =   300
         TabIndex        =   11
         Top             =   405
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cod. Cajero"
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
         Left            =   300
         TabIndex        =   10
         Top             =   855
         Width           =   1185
      End
      Begin VB.Label de 
         Caption         =   "de"
         Height          =   255
         Left            =   4575
         TabIndex        =   3
         Top             =   120
         Width           =   405
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":085E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":0B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":0E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":11AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":14C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":17E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":1AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":1E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":212E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":2240
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":2352
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":28F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":3126
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":38E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":41A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajero.frx":44C0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAdmCajero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RS As Recordset

Dim cl_Cajero As CLEquipo

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$

Dim ModoEdit As Integer
Dim DatoValido As Boolean

Dim HelpP As Boolean

Dim ModGrilla As Boolean

Private Function L_TodoCargado() As Boolean
    L_TodoCargado = False
    
    If mskLegajo.Text <> "" And txtDesc.Text <> "" _
    And mskCajero.Text <> "" Then
        L_TodoCargado = True
    Else
        L_TodoCargado = False
    End If
    
End Function

Private Function ArmarCondicion()
Dim Con$

Con$ = ""
If mskLegajo.Text <> "" Then
    Con$ = Con$ & " And V.legajo = " & mskLegajo.Text
End If

If mskCajero.Text <> "" Then
    Con$ = Con$ & " And V.Cod_cajero = " & mskCajero.Text
End If

ArmarCondicion = Con$

End Function
Private Sub MeActualizar()
Dim ViejoOrgan$
Dim Viejocargo%

If L_TodoCargado Then

FrmAdmCajero.caption = Aplicacion.SeteoProceso(FrmAdmCajero.caption)

Aplicacion.ComienzoTrans

MeLlenarObjeto


If cl_Cajero.Update_Cajero Then '
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


FrmAdmCajero.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub L_AltasDatos()

If L_TodoCargado Then
FrmAdmCajero.caption = Aplicacion.SeteoProceso(FrmAdmCajero.caption)

Aplicacion.ComienzoTrans

MeLlenarObjeto

If cl_Cajero.Insert_Cajero() Then
    Aplicacion.TerminarConExitoTrans
    chk.Value = 0
   
    NuevaSeleccion
    
Else
    Aplicacion.TerminarConErrorTrans
End If


FrmAdmCajero.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Public Sub Altas()
    SetearBotonesAltas True
    Modo = "ALTA"
    FrmAdmCajero.caption = "Administración Cajeros - ALTAS -"
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
Dim sql As String

SetearBotonesAltas False
Modo = "MOD"
FrmAdmCajero.caption = "Administración de Cajeros - Modificacion y Bajas -"

'Me.Show 1
End Sub
Private Sub NuevaSeleccion()
Dim i%

If Modo = "MOD" Then
    SetBotonesGeneral False
Else
    If chk.Value = 1 Then
        If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
            L_AltasDatos
        End If
        
    End If
End If
'Limpiar campos de pantallas
Set cl_Cajero = New CLEquipo

    'spr.MaxRows = 0
    txtDesc.Text = ""
    'txtCli.Text = ""
    mskLegajo.Text = ""
    mskCajero.Text = ""
    'chkActivo.Value = 0
    'mskMes.Text = ""
    
    chk.Value = 0

End Sub

Private Sub MeAbortarMod()
    
    SeteoBotonesMod True
    
    Tollbar.Buttons(2).Enabled = False
    
    MeSetearBotonesToolBar
    
    MellenarPantalla
    
End Sub


Private Sub MeCargarDatos()
Dim sql$

FrmAdmCajero.caption = Aplicacion.SeteoProceso(FrmAdmCajero.caption)
        
    CondConsulta = ArmarCondicion

    sqlGral$ = ""
    sqlGral$ = sqlGral$ & " SELECT V.cod_cajero ,"
    sqlGral$ = sqlGral$ & " V.legajo, "
    sqlGral$ = sqlGral$ & " P.apellido, "
    sqlGral$ = sqlGral$ & " P.nombre "
    sqlGral$ = sqlGral$ & " FROM PERSONAL.Empleado P , VENTAS.Cajeros V "
    sqlGral$ = sqlGral$ & " WHERE v.Legajo = p.Legajo "
    sqlGral$ = sqlGral$ & CondConsulta
    
If Aplicacion.ObtenerRsDAO(sqlGral$, RS) Then
    txtCantReg.Text = Aplicacion.CantReg(RS)
    If txtCantReg.Text > 0 Then
        txtReg.Text = 1
        SetBotonesGeneral True
        MellenarPantalla
        MeSetearBotonesToolBar
    Else
        txtReg.Text = 0
    End If
End If

FrmAdmCajero.caption = Aplicacion.SeteoFin

End Sub
Private Sub MeEliminar()

If MsgBox("Esta seguro de eliminar el registro", vbYesNo + vbExclamation, "ATENCION") = vbYes Then

'MeLlenarObjeto
cl_Cajero.Legajo = mskLegajo.Text
cl_Cajero.cajero = mskCajero.Text
FrmAdmCajero.caption = Aplicacion.SeteoProceso(FrmAdmCajero.caption)

Aplicacion.ComienzoTrans

If cl_Cajero.Delete_Cajero Then
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

FrmAdmCajero.caption = Aplicacion.SeteoFin

End If

End Sub

Private Sub MeLlenarObjeto()

Dim concprov As CLProdPrec
Dim i As Long, valor As Variant

cl_Cajero.Legajo = mskLegajo.Text
cl_Cajero.desc = txtDesc.Text
cl_Cajero.cajero = mskCajero.Text

End Sub




Private Sub MellenarPantalla()
Dim desc As String


mskLegajo.Text = RS!Legajo
txtDesc.Text = RS!Apellido & ", " & RS!NOMBRE
mskCajero.Text = RS!Cod_Cajero
'If rs!Activo = "S" Then
'    chkActivo.Value = 1
'Else
'    chkActivo.Value = 0
'End If

'mskFDesde.Text = Format$(rs!fch_vdesde, FTOFECHA)
'mskFHasta.Text = Format$(rs!fch_vhasta, FTOFECHA)

'mskMes.Text = rs!mes

'L_llenarProveedor

End Sub

'Public Sub PonerValores(cod As Variant, desc As String)
    
'    spr.SetText 1, spr.MaxRows, Trim(cod)
'    spr.SetText 2, spr.MaxRows, desc
'    Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
'    If spr.MaxRows > 6 Then
'        spr.Row = spr.MaxRows - 4
'    Else
'        spr.Row = spr.MaxRows
'    End If
'    spr.col = 1
'    spr.Position = SS_POSITION_UPPER_LEFT
'    spr.Action = SS_ACTION_GOTO_CELL

'DatoValido = True
'ModGrilla = True
'End Sub

Private Sub SetBotonesGeneral(valor As Boolean)
    
    Tollbar.Buttons(1).Enabled = Not valor
    Tollbar.Buttons(2).Enabled = Not valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = Not valor
    
    Tollbar.Buttons(7).Enabled = valor
    Tollbar.Buttons(8).Enabled = valor
    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor

    Tollbar.Buttons(12).Enabled = valor
    Tollbar.Buttons(13).Enabled = valor
'
'    TollBar.Buttons(12).Enabled = Not valor
'    TollBar.Buttons(13).Enabled = Not valor
'
'    Spread.spread_LockGrilla spr, valor
    frCab.Enabled = Not valor
'habilitar frames

    If Not valor Then
        txtReg.Text = 0
        txtCantReg.Text = 0
    End If

End Sub





Private Sub MePrepararMod()
    
    SeteoBotonesMod False
    DatoValido = True
    ModGrilla = False
    
End Sub

Private Function MeReconsultar() As Integer
Dim sql$
Dim i%
    

If Aplicacion.ObtenerRsDAO(sqlGral$, RS) Then
        txtCantReg.Text = Aplicacion.CantReg(RS)
        If Val(txtReg.Text) > Val(txtCantReg.Text) Then
            txtReg.Text = txtCantReg.Text
        End If
        
        For i% = 1 To txtReg.Text - 1
            RS.MoveNext
        Next
        If txtCantReg.Text > 0 Then
            MellenarPantalla
        End If
        'MeSetearBotonesToolBar
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
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(8).Enabled = False
    Tollbar.Buttons(9).Enabled = False
    Tollbar.Buttons(10).Enabled = False
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
ElseIf txtReg.Text = txtCantReg.Text Then
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(8).Enabled = True
    Tollbar.Buttons(9).Enabled = False
    Tollbar.Buttons(10).Enabled = False
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
ElseIf txtReg.Text = 1 Then
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(8).Enabled = False
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
Else
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(8).Enabled = True
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
    
End If
    


End Sub



Private Sub SetearBotonesAltas(valor As Boolean)
'valor = true -> altas
'valor = false -> modif
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(15).Enabled = valor
    Tollbar.Buttons(16).Visible = Not valor
    
    Tollbar.Buttons(17).Visible = Not valor 'False
    Tollbar.Buttons(18).Visible = Not valor 'False
    
    Tollbar.Buttons(5).Visible = Not valor 'False
    
    Tollbar.Buttons(7).Visible = Not valor 'False
    Tollbar.Buttons(8).Visible = Not valor 'False
    Tollbar.Buttons(9).Visible = Not valor 'False
    Tollbar.Buttons(10).Visible = Not valor 'False
    
    Tollbar.Buttons(12).Visible = Not valor 'False
    Tollbar.Buttons(13).Visible = Not valor 'False

    Tollbar.Buttons(18).Visible = Not valor 'False
    Tollbar.Buttons(19).Visible = Not valor 'False
    
    txtCantReg.Visible = Not valor 'False
    txtReg.Visible = Not valor 'False
    de.Visible = Not valor 'False
    
    'Spread.spread_LockGrilla spr, Not valor
    'txtCli.Locked = Not valor
    frCab.Enabled = valor
    
End Sub


Private Sub SeteoBotonesMod(valor As Boolean)
    
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = valor
    
    Tollbar.Buttons(7).Enabled = valor
    Tollbar.Buttons(8).Enabled = valor
    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor

    Tollbar.Buttons(12).Enabled = valor
    Tollbar.Buttons(13).Enabled = valor

    Tollbar.Buttons(15).Enabled = Not valor
    Tollbar.Buttons(16).Enabled = Not valor

    Tollbar.Buttons(18).Enabled = valor
    Tollbar.Buttons(19).Enabled = valor

'habilitar o des frames y/o campos
   
    frCab.Enabled = Not valor
    
End Sub



Private Sub botHelpCaj_Click()

Dim cod As String
Dim desc As String
Dim sql As String

HelpP = True

sql = "Select legajo,apellido || ', ' || nombre ape from PERSONAL.EMPLEADO "
 
If frmHelpPv.MuestraHlp(cod, desc, "CAJEROS", sql) = vbOK Then
    mskLegajo.Text = cod
    txtDesc.Text = desc
  Else
    mskLegajo.Text = ""
    txtDesc.Text = ""
End If

End Sub












Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If Shift = 4 Then
    Select Case KeyCode
        Case 78 'Nueva sel
            Call TollBar_ButtonClick(Tollbar.Buttons(4))
        Case 71 'Guardar
            If Tollbar.Buttons(12).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(15))
            End If
        Case 66 'Buscar
            If Modo = "MOD" Then
            Call TollBar_ButtonClick(Tollbar.Buttons(5))
            End If
        Case 83 'Salir
            Call TollBar_ButtonClick(Tollbar.Buttons(20))
       ' Case 107
       '     Call Toolbar1_ButtonClick(Toolbar1.Buttons(4))
       ' Case 109
       '     Call Toolbar1_ButtonClick(Toolbar1.Buttons(6))
    
    End Select
    If Modo = "MOD" And Val(txtCantReg.Text) > 0 Then
    Select Case KeyCode
        Case 37 'Izq
            If Tollbar.Buttons(8).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(5))
            End If
        Case 38 'Arriba
            Call TollBar_ButtonClick(Tollbar.Buttons(10))
        Case 40 'Abajo
            Call TollBar_ButtonClick(Tollbar.Buttons(7))
        Case 39 'Der
            If Tollbar.Buttons(6).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(9))
            End If
    End Select
    End If
Else
'    Select Case KeyCode
'        Case 107
'            Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
'        Case 109
'            Call Toolbar1_ButtonClick(Toolbar1.Buttons(3))
 '   End Select

End If
'Debug.Print KeyCode
End Sub

Private Sub MePrepararAlterar()

    Tollbar.Buttons(2).Value = tbrPressed
    Tollbar.Buttons(1).Value = tbrUnpressed
    
    Modificacion
    
    NuevaSeleccion
    
End Sub

Private Sub MePrepararAgregar()

    Tollbar.Buttons(1).Value = tbrPressed
    Tollbar.Buttons(2).Value = tbrUnpressed
    
    Altas
    
    NuevaSeleccion
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Tag = "" Then
        SendKeys "{TAB}"
    End If
End If
End Sub


Private Sub Form_Load()
Dim sql As String

Top = 800
Left = 1000
Width = 7500
Height = 2800

Set cl_Cajero = New CLEquipo

HelpP = False
Modo = "MOD"
    
'If Aplicacion.Nivel <> 0 Then
'    chkActivo.Enabled = False
'End If



End Sub





Private Sub spr_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If Mode = 1 Then
    ModoEdit = Mode
End If
End Sub


Private Sub spr_GotFocus()
FrmAdmPxC.Tag = "NOSALTAR"
End Sub




Private Sub mskLegajo_Change()
HelpP = False
End Sub

Private Sub mskLegajo_LostFocus()
Dim sql As String
Dim desc As String

If mskLegajo.Text <> "" Then
    If Not HelpP Then
       sql = "SELECT Apellido || ', ' || Nombre as descrip FROM personal.empleado " _
       & " WHERE legajo = " & mskLegajo.Text
    
      If Func_ObtenerDesc(sql, desc) Then
        txtDesc.Text = desc
       Else
        mskLegajo.SetFocus
        mskLegajo.Text = ""
        txtDesc.Text = ""
       End If
    Else
    
    End If
  
  Else
   txtDesc.Text = ""
End If

HelpP = False

End Sub


Private Sub TollBar_ButtonClick(ByVal Button As ComctlLib.Button)
Dim a%
Dim pos As String
Dim saltear As Boolean

saltear = True

pos = txtReg.Text

Select Case Button.Key
    Case "a"
         saltear = False
         Func_MoverPrimero RS, pos
    Case "b"
         saltear = False
        Func_MoverAnterior RS, pos
    Case "c"
         saltear = False
        Func_MoverSiguiente RS, pos
    Case "e"
         saltear = False
        Func_MoverUltimo RS, pos
    Case "f"
         MePrepararMod
    Case "g"
         MeEliminar
    Case "h"
        If Modo = "MOD" Then
            MeActualizar
        Else
            L_AltasDatos
        End If
    Case "i"
        MeAbortarMod
    Case "j"
        If chk.Value = 1 Then
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
        NuevaSeleccion
    Case "l"
        MeCargarDatos
    Case "n"
'        L_DatosGrilla
    
    Case "m"
        MeImpDatos
    Case "o"
        MePrepararAgregar
    Case "p"
        MePrepararAlterar
End Select

If Not saltear Then
    txtReg.Text = pos
    MellenarPantalla
    MeSetearBotonesToolBar
End If

End Sub

Private Sub txtDesc_Change()
If Modo = "ALTA" Then
    chk.Value = 1
End If
End Sub


