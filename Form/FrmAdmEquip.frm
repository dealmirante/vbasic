VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmAdmEquip 
   Caption         =   "Administración de Equipos"
   ClientHeight    =   3405
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   5940
   Icon            =   "FrmAdmEquip.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3405
   ScaleWidth      =   5940
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   105
      TabIndex        =   14
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
         NumButtons      =   18
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
            Object.Visible         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "Grilla"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "j"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "o"
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
      Left            =   5055
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   480
      Width           =   465
   End
   Begin VB.TextBox txtCantReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   480
      Width           =   480
   End
   Begin VB.Frame frCab 
      Height          =   2535
      Left            =   105
      TabIndex        =   16
      Top             =   465
      Width           =   5700
      Begin VB.CheckBox chkRub 
         Caption         =   "CAJ"
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
         Height          =   225
         Index           =   6
         Left            =   4845
         TabIndex        =   11
         Top             =   1200
         Width           =   690
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1980
         Width           =   1020
      End
      Begin VB.TextBox txtDescrip 
         Height          =   315
         Left            =   1455
         TabIndex        =   1
         Top             =   720
         Width           =   2205
      End
      Begin MSMask.MaskEdBox mskCod 
         Height          =   285
         Left            =   1450
         TabIndex        =   0
         Top             =   330
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.ListBox LstEspigon 
         Height          =   255
         Left            =   3210
         TabIndex        =   20
         Top             =   1920
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox CboCodAeropuerto 
         Height          =   315
         ItemData        =   "FrmAdmEquip.frx":0442
         Left            =   1450
         List            =   "FrmAdmEquip.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1140
         Width           =   2205
      End
      Begin VB.ComboBox CboEspigon 
         Height          =   315
         Left            =   1450
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Width           =   2205
      End
      Begin VB.CheckBox chkRub 
         Caption         =   "ACC"
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
         Height          =   225
         Index           =   0
         Left            =   3990
         TabIndex        =   5
         Top             =   465
         Width           =   690
      End
      Begin VB.CheckBox chkRub 
         Caption         =   "COS"
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
         Height          =   225
         Index           =   4
         Left            =   3990
         TabIndex        =   9
         Top             =   1800
         Width           =   690
      End
      Begin VB.CheckBox chkRub 
         Caption         =   "COM"
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
         Height          =   240
         Index           =   3
         Left            =   3990
         TabIndex        =   8
         Top             =   1455
         Width           =   765
      End
      Begin VB.CheckBox chkRub 
         Caption         =   "CIG"
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
         Height          =   225
         Index           =   2
         Left            =   4005
         TabIndex        =   7
         Top             =   1110
         Width           =   690
      End
      Begin VB.CheckBox chkRub 
         Caption         =   "PER"
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
         Height          =   225
         Index           =   5
         Left            =   3990
         TabIndex        =   10
         Top             =   2160
         Width           =   690
      End
      Begin VB.CheckBox chkRub 
         Caption         =   "BEB"
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
         Height          =   240
         Index           =   1
         Left            =   3990
         TabIndex        =   6
         Top             =   780
         Width           =   690
      End
      Begin VB.Label LblEspigon 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupo"
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
         Left            =   225
         TabIndex        =   23
         Top             =   1980
         Width           =   1170
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
         Left            =   225
         TabIndex        =   22
         Top             =   1140
         Width           =   1170
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
         Left            =   225
         TabIndex        =   21
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción"
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
         Left            =   225
         TabIndex        =   19
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código"
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
         Index           =   8
         Left            =   225
         TabIndex        =   18
         Top             =   330
         Width           =   1170
      End
      Begin VB.Label de 
         Caption         =   "de"
         Height          =   255
         Left            =   4665
         TabIndex        =   17
         Top             =   120
         Width           =   405
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
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":0446
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":0760
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":0A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":0D94
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":10AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":13C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":16E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":19FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":1D16
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":2030
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":2142
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":2254
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmEquip.frx":27F6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAdmEquip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As Recordset

Dim cl_Eq As CLEquipo

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$
Public Sub Altas()
    Modo = "ALTA"
    SetearBotonesAltas
    FrmAdmEquip.caption = FrmAdmEquip.caption & " -Altas- "
    Me.Show 1
End Sub
Public Sub EquipoAyuda(ByRef cl As CLEquipo)

SetearBotonesAyuda
Modo = "AYUDA"
Me.Show 1

Set cl = cl_Eq

End Sub

Private Sub L_AltasDatos()

If L_TodoCargado Then
FrmAdmEquip.caption = Aplicacion.SeteoProceso(FrmAdmEquip.caption)

Aplicacion.ComienzoTrans

MeLlenarObjeto

If cl_Eq.Insert_Equipo() Then
    Aplicacion.TerminarConExitoTrans
    chk.Value = 0
   
    NuevaSeleccion
    
Else
    Aplicacion.TerminarConErrorTrans
End If

FrmAdmEquip.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub L_PonerRubros()
Dim rubT As String, rub As String
Dim pos, i

    For i = 0 To 6
        chkRub(i).Value = 0
    Next

rubT = rs!Rubros

pos = InStr(1, rubT, "/")

Do While pos <> 0
    rub = Left$(rubT, pos - 1)
    For i = 0 To 5
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

Private Function L_Rubros() As String
Dim i
Dim rub As String

rub = ""

For i = 0 To 6
    If chkRub(i).Value = 1 Then
        rub = rub & chkRub(i).caption & "/"
    End If
Next

If rub = "" Then
    L_Rubros = ""
Else
    L_Rubros = Left(rub, Len(rub) - 1)
End If

End Function

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
FrmAdmEquip.caption = FrmAdmEquip.caption & " -Modificación y Bajas- "
Me.Show 1
End Sub
Private Sub NuevaSeleccion()
Dim i

If Modo = "MOD" Or Modo = "AYUDA" Then
    SetBotonesGeneral False
    cboGrupo.ListIndex = -1
Else
    If chk.Value = 1 Then
        If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
            L_AltasDatos
        End If
    End If
End If
'Limpiar campos de pantallas
Set cl_Eq = New CLEquipo

mskCod.Text = ""
txtDescrip.Text = ""
CboCodAeropuerto.ListIndex = -1

For i = 0 To 6
    chkRub(i).Value = 0
Next

mskCod.SetFocus
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

FrmAdmEquip.caption = Aplicacion.SeteoProceso(FrmAdmEquip.caption)

Aplicacion.ComienzoTrans

MeLlenarObjeto


If cl_Eq.Update_Equipos Then '
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


FrmAdmEquip.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub MeCargarDatos()
Dim sql$

FrmAdmEquip.caption = Aplicacion.SeteoProceso(FrmAdmEquip.caption)
        
    CondConsulta = ArmarCondicion
    
    sqlGral$ = " SELECT cod_equipo, " _
                & " descrip, " _
                & " cod_depn, " _
                & " cod_sdep, " _
                & " grupo, " _
                & " rubros " _
                & " FROM estadis.Equipos " & CondConsulta
    
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

FrmAdmEquip.caption = Aplicacion.SeteoFin

End Sub
Private Sub MeEliminar()

If MsgBox("Esta seguro de eliminar el registro", vbYesNo + vbExclamation, "ATENCION") = vbYes Then

MeLlenarObjeto

FrmAdmEquip.caption = Aplicacion.SeteoProceso(FrmAdmEquip.caption)

Aplicacion.ComienzoTrans

If cl_Eq.Delete_Equipos Then
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

FrmAdmEquip.caption = Aplicacion.SeteoFin

End If

End Sub

Private Sub MeLlenarObjeto()

cl_Eq.cod = IIf(mskCod.Text = "", 0, mskCod.Text)
cl_Eq.desc = txtDescrip.Text
cl_Eq.CodDep = CboCodAeropuerto.Text
cl_Eq.CodSdep = LstEspigon.List(CboEspigon.ListIndex)
cl_Eq.Grupo = cboGrupo.Text
cl_Eq.Rubros = L_Rubros

End Sub


Private Function L_TodoCargado() As Boolean
    
If CboEspigon.Text <> "" And cboGrupo.Text <> "" _
   And mskCod.Text <> "" And L_Rubros <> "" Then
    L_TodoCargado = True
Else
    L_TodoCargado = False
End If


End Function

Private Sub MellenarPantalla()

mskCod.Text = rs!cod_equipo
txtDescrip.Text = IIf(IsNull(rs!Descrip), "", rs!Descrip)

Func.Func_SetearCboSTR CboCodAeropuerto, rs!cod_depn

Func.Func_SetearCboConLst CboEspigon, LstEspigon, rs!cod_sdep

Func.Func_SetearCboSTR cboGrupo, rs!Grupo

L_PonerRubros

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

'    TollBar.Buttons(12).Enabled = Not valor
'    TollBar.Buttons(13).Enabled = Not valor


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
If L_Rubros <> "" Then
    Con$ = Con$ & " And rubros = '" & L_Rubros & "'"
End If
If CboCodAeropuerto.Text <> "" Then
    Con$ = Con$ & " And Cod_depn = '" & CboCodAeropuerto.Text & "'"
End If
If CboEspigon.Text <> "" Then
    Con$ = Con$ & " And Cod_sdep = '" & LstEspigon.List(CboEspigon.ListIndex) & "'"
End If
If cboGrupo.Text <> "" Then
    Con$ = Con$ & " And grupo = '" & cboGrupo.Text & "'"
End If

If Con$ <> "" Then
    Con$ = " WHERE " & Mid(Con$, 5, Len(Con$))
End If

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

Private Sub SetearBotonesAyuda()
    
    Tollbar.Buttons(9).Visible = False
    Tollbar.Buttons(10).Visible = False
    Tollbar.Buttons(12).Visible = False
    Tollbar.Buttons(13).Visible = False
    Tollbar.Buttons(17).Visible = False

    Tollbar.Buttons(18).Visible = True
    
'    txtCantReg.Visible = False
'    txtReg.Visible = False
'    de.Visible = False

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
mskCod.Enabled = valor

End Sub





Private Sub CboCodAeropuerto_Click()
Dim sql As String

sql = " SELECT cod_sdep,descrip FROM baires.subdependencia "
sql = sql & " WHERE cod_depn = '" & CboCodAeropuerto.Text & "'"
sql = sql & " ORDER BY cod_sdep"
 
FuncCbos_LlenarCboLst CboEspigon, LstEspigon, sql

End Sub


Private Sub CboCodAeropuerto_KeyPress(KeyAscii As Integer)
If Modo = "MOD" Or Modo = "AYUDA" Then
    If KeyAscii = 32 Then
        CboCodAeropuerto.ListIndex = -1
    End If
End If
End Sub


Private Sub CboEspigon_KeyPress(KeyAscii As Integer)
If Modo = "MOD" Or Modo = "AYUDA" Then
    If KeyAscii = 32 Then
        CboEspigon.ListIndex = -1
    End If
End If
End Sub


Private Sub cboGrupo_KeyPress(KeyAscii As Integer)
If Modo = "MOD" Or Modo = "AYUDA" Then
    If KeyAscii = 32 Then
        cboGrupo.ListIndex = -1
    End If
End If

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
        Case 66 'Buscar
            If Modo = "MOD" Then
            Call TollBar_ButtonClick(Tollbar.Buttons(2))
            End If
        Case 83 'Salir
            Call TollBar_ButtonClick(Tollbar.Buttons(16))
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
Dim sql As String

Width = Screen.Width * 0.6
Height = Screen.Height * 0.5
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height - 500) / 2


sql = " SELECT cod_depn,descrip FROM baires.dependencia "
sql = sql & " ORDER BY cod_depn"

FuncCbos_LlenarCbo CboCodAeropuerto, sql

cboGrupo.AddItem "A"
cboGrupo.AddItem "B"
cboGrupo.AddItem "C"

If Modo = "ALTA" Then
    cboGrupo.ListIndex = 0
End If

Set cl_Eq = New CLEquipo

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
         Func_MoverPrimero rs, pos
    Case "b"
         saltear = False
        Func_MoverAnterior rs, pos
    Case "c"
         saltear = False
        Func_MoverSiguiente rs, pos
    Case "d"
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
    Case "o"
        MeLlenarObjeto
        Unload Me
End Select

If Not saltear Then
    txtReg.Text = pos
    MellenarPantalla
    MeSetearBotonesToolBar
End If

End Sub
Private Sub L_DatosGrilla()
Dim i
Dim nro As Integer

On Error GoTo DG:

'nro = frmGrid.DatosGrilla(sqlGral$)

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


