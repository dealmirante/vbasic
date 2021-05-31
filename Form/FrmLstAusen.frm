VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmLstAusen 
   Caption         =   "Administración de "
   ClientHeight    =   3165
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   6720
   Icon            =   "FrmLstAusen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3165
   ScaleWidth      =   6720
   Begin VB.Frame frCab 
      Height          =   2595
      Left            =   135
      TabIndex        =   3
      Top             =   45
      Width           =   6390
      Begin VB.CommandButton botFHasta 
         Height          =   345
         Left            =   5835
         Picture         =   "FrmLstAusen.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   435
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ListBox lstAus 
         Height          =   255
         Left            =   4830
         TabIndex        =   9
         Top             =   1530
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.ComboBox cboAus 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1485
         Width           =   3000
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2805
         Picture         =   "FrmLstAusen.frx":05B4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   435
         Width           =   375
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   3420
      End
      Begin MSMask.MaskEdBox mskLegajo 
         Height          =   285
         Left            =   1575
         TabIndex        =   1
         Top             =   975
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1605
         TabIndex        =   0
         Top             =   450
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFHasta 
         Height          =   285
         Left            =   4665
         TabIndex        =   11
         Top             =   465
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
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
         Index           =   4
         Left            =   3375
         TabIndex        =   10
         Top             =   465
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Causa"
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
         Left            =   315
         TabIndex        =   8
         Top             =   1485
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
         Left            =   315
         TabIndex        =   7
         Top             =   465
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
         Left            =   315
         TabIndex        =   5
         Top             =   975
         Width           =   1185
      End
   End
End
Attribute VB_Name = "FrmLstAusen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RS As Recordset

Dim cl_Eq As CLEquipo

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$
Public Sub Altas()
    SetearBotonesAltas True
    Modo = "ALTA"
    FrmAdmAusen.caption = "Administracion de Ausencias" & " -Altas- "
'    Me.Show 1
End Sub
Private Sub L_AltasDatos()
Dim fch As Date

If L_TodoCargado Then
       
FrmAdmAusen.caption = Aplicacion.SeteoProceso(FrmAdmAusen.caption)

For fch = CDate(mskFDesde.FormattedText) To CDate(mskFHasta.FormattedText)
    If L_TesteoDia(fch) Then
        Aplicacion.ComienzoTrans
        
        MeLlenarObjeto
        
        If cl_Eq.Insert_Ausent(fch) Then
            Aplicacion.TerminarConExitoTrans
        Else
            Aplicacion.TerminarConErrorTrans
        End If
    Else
        MsgBox mskLegajo.Text & " de Franco el día " & Format$(fch, FTOFECHA), vbOKOnly, "Atención"
    End If
Next
    chk.Value = 0
        
    NuevaSeleccion

    FrmAdmAusen.caption = Aplicacion.SeteoFin
Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Function L_TesteoDia(d As Date) As Boolean
Dim sql As String
Dim desc As String

If txtGr.Text = "A" Or txtGr.Text = "B" Or txtGr.Text = "C" Then
sql = " SELECT grupo descrip From personal.ROTACION Where " _
& " FECHA = to_date('02/01/01','dd/mm/yy') + MOD(" & func_ToDate(Format$(d, FTOFECHA)) & "-to_date('02/01/01','dd/mm/yy'),6)  " _
& " And grupo = '" & txtGr.Text & "'"
Else
sql = " SELECT grupo descrip From personal.ROTACION Where " _
& " FECHA = to_date('02/01/01','dd/mm/yy') + MOD(" & func_ToDate(Format$(d, FTOFECHA)) & "-to_date('02/01/01','dd/mm/yy'),7)  " _
& " And grupo = '" & txtGr.Text & "'"
End If

    If Func_ObtenerDesc(sql, desc) Then
        L_TesteoDia = True
    Else
        L_TesteoDia = False
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

Private Sub MePrepararAgregar()

    Tollbar.Buttons(1).Value = tbrPressed
    Tollbar.Buttons(2).Value = tbrUnpressed
    
    mskFHasta.Visible = True
    botFHasta.Visible = True
    Label1(4).Visible = True
    Altas
    
End Sub
Private Sub MePrepararAlterar()

    Tollbar.Buttons(2).Value = tbrPressed
    Tollbar.Buttons(1).Value = tbrUnpressed
    
    mskFHasta.Visible = False
    botFHasta.Visible = False
    Label1(4).Visible = False
    
    Modificacion
    
End Sub

Public Sub Modificacion()

SetearBotonesAltas False
Modo = "MODIF"
FrmAdmAusen.caption = "Administracion de Ausencias" & " -Modificacion y Bajas- "
' Me.Show 1
End Sub
Private Sub NuevaSeleccion()
Dim i%

If Modo = "MODIF" Then
    SetBotonesGeneral False
    mskFDesde.Text = ""
    mskFHasta.Text = ""
Else
    If chk.Value = 1 Then
        If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
            L_AltasDatos
        End If
    End If
End If
'Limpiar campos de pantallas

Set cl_Eq = New CLEquipo


mskLegajo.Text = ""
txtDesc.Text = ""
txtGr.Text = ""

cboAus.ListIndex = -1

mskFDesde.SetFocus

chk.Value = 0

End Sub

Private Sub MeAbortarMod()
    
If Modo = "MODIF" Then
    SeteoBotonesMod True
    
    Tollbar.Buttons(2).Enabled = False
    
    MeSetearBotonesToolBar
    
    MellenarPantalla
Else
    
End If

End Sub

Private Sub MeActualizar()
Dim ViejoOrgan$
Dim Viejocargo%

If L_TodoCargado Then

FrmAdmAusen.caption = Aplicacion.SeteoProceso(FrmAdmAusen.caption)

Aplicacion.ComienzoTrans

MeLlenarObjeto


If cl_Eq.Update_Ausen Then '
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


FrmAdmAusen.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub MeCargarDatos()
Dim sql$

FrmAdmAusen.caption = Aplicacion.SeteoProceso(FrmAdmAusen.caption)
        
    CondConsulta = ArmarCondicion

    sqlGral$ = ""
    sqlGral$ = " SELECT fecha,A.legajo,Apellido||', '||nombre ape ,causa" _
    & " From personal.ausencia A, personal.empleado E " _
    & " WHERE A.legajo = E.legajo (+) " _
    & CondConsulta
    

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

FrmAdmAusen.caption = Aplicacion.SeteoFin

End Sub
Private Sub MeEliminar()

If MsgBox("Esta seguro de eliminar el registro", vbYesNo + vbExclamation, "ATENCION") = vbYes Then

MeLlenarObjeto

FrmAdmAusen.caption = Aplicacion.SeteoProceso(FrmAdmAusen.caption)

Aplicacion.ComienzoTrans

If cl_Eq.Delete_Ausen Then
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
'
FrmAdmAusen.caption = Aplicacion.SeteoFin
End If

End Sub

Private Sub MeLlenarObjeto()

cl_Eq.Legajo = mskLegajo.Text
cl_Eq.desc = lstAus.List(cboAus.ListIndex)
cl_Eq.TipoTime = mskFDesde.FormattedText

End Sub


Private Function L_TodoCargado() As Boolean
    
If mskLegajo.Text <> "" And mskFDesde.Text <> "" And mskFHasta.Text <> "" And cboAus.Text <> "" Then
    L_TodoCargado = True
Else
    L_TodoCargado = False
End If


End Function

Private Sub MellenarPantalla()
Dim sql As String
Dim desc As String

mskFDesde.Text = Format$(RS!Fecha, FTOFECHA)
mskLegajo.Text = RS!Legajo
txtDesc.Text = RS!ape

Func_SetearCboConLst cboAus, lstAus, RS!causa

    sql = "SELECT grupo descrip FROM ESTADIS.PERSONA_EQUIPOS " _
    & " WHERE legajo = " & mskLegajo.Text
    If Func_ObtenerDesc(sql, desc) Then
        txtGr.Text = desc
    Else
        txtGr.Text = "NN"
    End If

End Sub

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
    Tollbar.Buttons(18).Enabled = valor
    Tollbar.Buttons(19).Enabled = valor

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
If mskLegajo.Text <> "" Then
    Con$ = Con$ & " And A.legajo = " & mskLegajo.Text
End If
If mskFDesde.Text <> "" Then
    Con$ = Con$ & " And fecha = " & func_ToDate(mskFDesde.FormattedText)
End If
If cboAus.Text <> "" Then
    Con$ = Con$ & " And A.causa = '" & lstAus.List(cboAus.ListIndex) & "' "
End If
'If Con$ <> "" Then
'    Con$ = " WHERE " & Mid(Con$, 5, Len(Con$))
'End If

ArmarCondicion = Con$

End Function



Private Sub MePrepararMod()
    
    SeteoBotonesMod False

End Sub

Private Function MeReconsultar() As Integer
Dim sql$
Dim i%
    
FrmAdmAusen.caption = Aplicacion.SeteoProceso(FrmAdmAusen.caption)
    

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

FrmAdmAusen.caption = Aplicacion.SeteoProceso(FrmAdmAusen.caption)

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
    Tollbar.Buttons(16).Enabled = valor
    
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

  '  frCab.Enabled = Not valor

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
    mskFDesde.Enabled = valor
    mskLegajo.Enabled = valor
End Sub





Private Sub botCuadro_Click()
frmControlAus.Show
End Sub

Private Sub botFHasta_Click()
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



Private Sub cboAus_Change()
FrmAdmAusen.Tag = "T"
End Sub

Private Sub cboAus_LostFocus()
            If Tollbar.Buttons(15).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(15))
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
            If Modo = "MODIF" Then
            Call TollBar_ButtonClick(Tollbar.Buttons(5))
            End If
        Case 83 'Salir
            Call TollBar_ButtonClick(Tollbar.Buttons(19))
    End Select
    If Modo = "MODIF" And Val(txtCantReg.Text) > 0 Then
    Select Case KeyCode
        Case 37 'Izq
            If Tollbar.Buttons(5).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(8))
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

Top = 1000
Left = 800
Width = 6700
Height = 3700

Set cl_Eq = New CLEquipo

'mskFDesde.Text = Format$(Date - 1, FTOFECHA)

sql = " SELECT causa,descrip FROM personal.causa "
sql = sql & " ORDER BY descrip "

FuncCbos_LlenarCboLst cboAus, lstAus, sql

Modo = "MODIF"

End Sub






Private Sub mskFDesde_Change()
FrmAdmAusen.Tag = ""
End Sub

Private Sub mskFDesde_LostFocus()

If mskFDesde.Text <> "" Then
    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    End If
    
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
    mskFHasta.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
    
End If


End Sub


Private Sub mskFHasta_LostFocus()

If mskFHasta.Text <> "" Then
    If Not IsDate(mskFHasta.FormattedText) Then
        mskFHasta.Text = Format$(Date - 1, FTOFECHA)
    End If
    
    mskFHasta.Text = Format$(mskFHasta.FormattedText, FTOFECHA)
    
End If

End Sub


Private Sub mskLegajo_Change()
FrmAdmAusen.Tag = ""
End Sub

Private Sub mskLegajo_LostFocus()
Dim sql As String
Dim desc As String

If mskLegajo.Text <> "" Then
    sql = "SELECT Apellido || ', ' || Nombre as descrip FROM personal.empleado " _
    & " WHERE legajo = " & mskLegajo.Text
    
    If Func_ObtenerDesc(sql, desc) Then
        txtDesc.Text = desc
    Else
        txtDesc.Text = "Legajo no resgistrado"
    End If
    
    sql = "SELECT grupo descrip FROM ESTADIS.PERSONA_EQUIPOS " _
    & " WHERE legajo = " & mskLegajo.Text
    If Func_ObtenerDesc(sql, desc) Then
        txtGr.Text = desc
    Else
        txtGr.Text = "NN"
    End If

End If

End Sub

Private Sub TollBar_ButtonClick(ByVal Button As Button)
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
        If Modo = "MODIF" Then
            MeActualizar
        Else
            L_AltasDatos
        End If
    Case "i"
        MeAbortarMod
    Case "j"
        If chk.Value = 1 Then
            If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
                If Modo = "MODIF" Then
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
Private Sub L_DatosGrilla()
Dim i
Dim nro As Integer

On Error GoTo DG:

nro = frmGridAus.DatosGrilla(sqlGral$)

If nro > 0 Then
    RS.MoveFirst
    For i = 1 To nro - 1
        RS.MoveNext
    Next
    MellenarPantalla
    txtReg.Text = nro
    MeSetearBotonesToolBar
End If

DG:
    Exit Sub


End Sub

