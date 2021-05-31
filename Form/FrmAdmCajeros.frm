VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmAdmCajeros 
   Caption         =   "Administración de "
   ClientHeight    =   2145
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   7170
   Icon            =   "FrmAdmCajeros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   7170
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   135
      TabIndex        =   6
      Top             =   15
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
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
            Object.Visible         =   0   'False
            Key             =   "m"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "Vista General"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
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
      Left            =   6030
      TabIndex        =   7
      Top             =   90
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5595
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   450
      Width           =   465
   End
   Begin VB.TextBox txtCantReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6390
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   450
      Width           =   480
   End
   Begin VB.Frame frCab 
      Height          =   1395
      Left            =   90
      TabIndex        =   8
      Top             =   465
      Width           =   6900
      Begin VB.CommandButton botHelpEq 
         Height          =   285
         Left            =   6330
         Picture         =   "FrmAdmCajeros.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   375
         Width           =   375
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   3360
      End
      Begin MSMask.MaskEdBox mskLegajo 
         Height          =   285
         Left            =   1620
         TabIndex        =   0
         Top             =   360
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskCajero 
         Height          =   285
         Left            =   1635
         TabIndex        =   1
         Top             =   810
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label de 
         Caption         =   "de"
         Height          =   255
         Left            =   6030
         TabIndex        =   11
         Top             =   120
         Width           =   405
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
         Left            =   360
         TabIndex        =   10
         Top             =   810
         Width           =   1185
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
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1185
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   -60
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":085E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":0B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":0E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":11AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":14C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":17E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":1AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":1E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":212E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":2240
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":2352
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmCajeros.frx":28F4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAdmCajeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As Recordset

Dim cl_Cajeros As CLEquipo

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$
Dim HelpEq As Boolean



Public Sub Altas()
    SetearBotonesAltas
    Modo = "ALTAS"
    FrmAdmCajeros.caption = FrmAdmCajeros.caption & " -Altas- "
    Me.Show 1
End Sub
Private Sub L_AltasDatos()

If L_TodoCargado Then
FrmAdmCajeros.caption = Aplicacion.SeteoProceso(FrmAdmCajeros.caption)

Aplicacion.ComienzoTrans

MeLlenarObjeto

If cl_Cajeros.Insert_Cajero() Then
    Aplicacion.TerminarConExitoTrans
    chk.Value = 0
   
    NuevaSeleccion
    
Else
    Aplicacion.TerminarConErrorTrans
End If


FrmAdmCajeros.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub L_PonerEsp(Esp As String)

Select Case Esp
    Case "INTA"
        optEsp(0).Value = True
    Case "INTB"
        optEsp(1).Value = True
    Case "AEP"
        optEsp(2).Value = True
End Select
End Sub

Private Sub L_PonerGrupo(GR As String)

optGr(Asc(GR) - 65).Value = True

End Sub

Private Sub MeImpDatos()
Dim Nom As String, NombreArchivo As String


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
FrmAdmPersonas.caption = FrmAdmPersonas.caption & " -Modificacion y Bajas- "
Me.Show 1
End Sub
Private Sub NuevaSeleccion()
Dim i

If Modo = "MOD" Then
    SetBotonesGeneral False
    mskCod.Text = ""
    'For i = 0 To 5
    '    chkRub(i).Value = 0
    'Next
    'For i = 0 To 2
    '    optGr(i).Value = False
    '    optEsp(i).Value = False
    'Next
Else
    If chk.Value = 1 Then
        If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
            L_AltasDatos
        End If
    End If
End If
'Limpiar campos de pantallas
Set cl_P = New CLEquipo

mskLegajo.Text = ""
txtDesc.Text = ""
'txtDescEq.Text = ""
'if msklegajo.Enabled then
mskLegajo.SetFocus
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

'frm_.caption = Aplicacion.SeteoProceso (frm_.caption)
        
CondConsulta = ArmarCondicion

sqlGral$ = "SELECT V.legajo,V.cod_cajero " _
& " FROM ventas.cajero V, baires.cajero C " _
& " where V.cod_cajero = V.cod_cajero " _
& CondConsulta & " ORDER BY cod_cajero "


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

'frm_.caption = Aplicacion.Seteofin

End Sub
Private Sub MeEliminar()

If MsgBox("Esta seguro de eliminar el registro", vbYesNo + vbExclamation, "ATENCION") = vbYes Then

MeLlenarObjeto

FrmAdmPersonas.caption = Aplicacion.SeteoProceso(FrmAdmPersonas.caption)

Aplicacion.ComienzoTrans

If cl_P.Delete_Persona Then
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

FrmAdmPersonas.caption = Aplicacion.SeteoFin
End If

End Sub

Private Sub MeLlenarObjeto()
 
 cl_P.cod = mskCod.Text
 cl_P.Legajo = mskLegajo.Text

End Sub


Private Sub MellenarPantalla()
mskLegajo.Text = rs!Legajo
mskCod.Text = rs!cod_cajero

'Call mskCod_LostFocus
'Call mskLegajo_LostFocus

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
'
    Tollbar.Buttons(16).Enabled = valor
'    TollBar.Buttons(13).Enabled = Not valor
'

'habilitar frames
frCab.Enabled = Not valor
    
    If Not valor Then
        txtReg.Text = 0
        txtCantReg.Text = 0
    End If

End Sub



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
    Tollbar.Buttons(15).Visible = False
    
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
frCab.Enabled = Not valor
mskLegajo.Enabled = valor

End Sub





Private Sub botHelpEq_Click()
Dim cl As CLEquipo

Set cl = New CLEquipo

FrmAdmEquip.EquipoAyuda cl

HelpEq = True

If cl.cod <> 0 Then
    mskCod.Text = cl.cod
    txtDescEq.Text = cl.desc
    
    L_PonerRubros cl.Rubros
    L_PonerGrupo cl.Grupo
    L_PonerEsp cl.CodSdep
End If
If mskCod.Enabled Then
    mskCod.SetFocus
End If

End Sub

Private Sub L_PonerRubros(RUBR As String)
Dim rubT As String, rub As String
Dim pos, i

    For i = 0 To 6
        chkRub(i).Value = 0
    Next

rubT = RUBR

pos = InStr(1, rubT, "/")

Do While pos <> 0
    rub = Left$(rubT, pos - 1)
    For i = 0 To 6
        If chkRub(i).caption = rub Then
            chkRub(i).Value = 1
            Exit For
        End If
    Next
    rubT = Right(rubT, Len(rubT) - pos)
    pos = InStr(1, rubT, "/")
Loop
    For i = 0 To 6
        If chkRub(i).caption = rubT Then
            chkRub(i).Value = 1
            Exit For
        End If
    Next

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 4 Then
    Select Case KeyCode
        Case 78 'Nueva sel
            Call Tollbar_ButtonClick(Tollbar.Buttons(1))
        Case 71 'Guardar
            If Tollbar.Buttons(12).Enabled Then
            Call Tollbar_ButtonClick(Tollbar.Buttons(12))
            End If
        Case 66 'Buscar
            If Modo = "MOD" Then
            Call Tollbar_ButtonClick(Tollbar.Buttons(2))
            End If
        Case 83 'Salir
            Call Tollbar_ButtonClick(Tollbar.Buttons(16))
    End Select
    If Modo = "MOD" And Val(txtCantReg.Text) > 0 Then
    Select Case KeyCode
        Case 37 'Izq
            If Tollbar.Buttons(5).Enabled Then
            Call Tollbar_ButtonClick(Tollbar.Buttons(5))
            End If
        Case 38 'Arriba
            Call Tollbar_ButtonClick(Tollbar.Buttons(7))
        Case 40 'Abajo
            Call Tollbar_ButtonClick(Tollbar.Buttons(4))
        Case 39 'Der
            If Tollbar.Buttons(6).Enabled Then
            Call Tollbar_ButtonClick(Tollbar.Buttons(6))
            End If
    End Select
    End If

End If
'Debug.Print KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub Form_Load()
Width = Screen.Width * 0.65
Height = Screen.Height * 0.6
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height - 500) / 2

Set cl_P = New CLEquipo

End Sub
Private Sub mskCod_LostFocus()
Dim sql As String
Dim rs As Recordset
Dim cod, i

    'If Not HelpEq Then
    '    cod = IIf(mskCod.Text = "", -1, mskCod.Text)
    '    sql = " SELECT cod_equipo, " _
    '          & " descrip, " _
    '          & " cod_depn, " _
    '          & " cod_sdep, " _
    '          & " grupo, " _
    '          & " rubros " _
    '          & " FROM estadis.Equipos " _
    '          & " WHERE cod_equipo = " & cod
    '
    '        If Aplicacion.ObtenerRsDAO(sql, rs) Then
    '            If Aplicacion.CantReg(rs) > 0 Then
    '                txtDescEq.Text = IIf(IsNull(rs!Descrip), "", rs!Descrip)
    '                L_PonerRubros rs!Rubros
    '                L_PonerGrupo rs!Grupo
    '                L_PonerEsp rs!cod_sdep
    '            Else
    '                txtDescEq.Text = ""
    '                L_PonerRubros ""
    '                For i = 0 To 2
    '                    optGr(i).Value = False
    '                    optEsp(i).Value = False
    '                Next
    '            End If
    '            Aplicacion.CerrarDAO rs
    '        End If
        
    'End If

    'HelpEq = False
    If mskLegajo.Enabled Then
        mskLegajo.SetFocus
    End If
End Sub


Private Sub Tollbar_ButtonClick(ByVal Button As ComctlLib.Button)
Dim A%
Dim pos As String
Dim saltear As Boolean

saltear = True

pos = txtReg.Text

Select Case Button.Key
    Case "a"
         saltear = False
         Func_MoverPrimero rs, pos
    Case "b"
         saltear = False
        Func_MoverAnterior rs, pos
    Case "c"
         saltear = False
        Func_MoverSiguiente rs, pos
    Case "e"
         saltear = False
        Func_MoverUltimo rs, pos
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
        L_DatosGrilla
    
    Case "m"
        MeImpDatos
    
End Select

If Not saltear Then
    txtReg.Text = pos
    MellenarPantalla
    MeSetearBotonesToolBar
End If

End Sub
Private Sub L_DatosGrilla()
Dim i, sql As String
Dim nro As Integer

On Error GoTo DG:

sql = "SELECT PE.legajo,apellido || ', ' || nombre Ape" _
& " FROM estadis.persona_equipos PE, estadis.equipos E, personal.Empleado EM" _
& " where E.cod_equipo = PE.cod_equipo " _
& " And PE.legajo = EM.legajo " _
& CondConsulta & " ORDER BY PE.COD_EQUIPO "

nro = frmGridPerEquip.DatosGrilla(sql)

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


