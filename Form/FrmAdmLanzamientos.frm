VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form FrmAdmLanzamientos 
   Caption         =   "Administración de "
   ClientHeight    =   5400
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   8550
   Icon            =   "FrmAdmLanzamientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   8550
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   120
      TabIndex        =   2
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
            Key             =   ""
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
            Key             =   ""
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
            Key             =   ""
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
            Key             =   ""
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
            Key             =   ""
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
      MouseIcon       =   "FrmAdmLanzamientos.frx":0442
   End
   Begin VB.Frame Frame1 
      Height          =   3840
      Left            =   165
      TabIndex        =   6
      Top             =   1425
      Width           =   8280
      Begin FPSpread.vaSpread spr 
         Height          =   2865
         Left            =   105
         OleObjectBlob   =   "FrmAdmLanzamientos.frx":045E
         TabIndex        =   15
         Top             =   690
         Width           =   8070
      End
      Begin VB.Frame frBot 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   450
         Left            =   105
         TabIndex        =   7
         Top             =   180
         Width           =   7305
         Begin ComctlLib.Toolbar Toolbar1 
            Height          =   420
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   741
            ButtonWidth     =   635
            ButtonHeight    =   582
            Appearance      =   1
            ImageList       =   "ImageList1"
            _Version        =   327680
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   7
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "A"
                  Object.ToolTipText     =   "Agreagar Fila"
                  Object.Tag             =   ""
                  ImageIndex      =   16
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   ""
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "B"
                  Object.ToolTipText     =   "Sacar Fila"
                  Object.Tag             =   ""
                  ImageIndex      =   17
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   ""
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "C"
                  Object.ToolTipText     =   "Limpiar Todo"
                  Object.Tag             =   ""
                  ImageIndex      =   9
               EndProperty
               BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Visible         =   0   'False
                  Key             =   ""
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Visible         =   0   'False
                  Key             =   "D"
                  Object.ToolTipText     =   "Salir"
                  Object.Tag             =   ""
               EndProperty
            EndProperty
            BorderStyle     =   1
            MouseIcon       =   "FrmAdmLanzamientos.frx":07A6
         End
         Begin VB.CommandButton botHelpProd 
            Height          =   315
            Left            =   6765
            Picture         =   "FrmAdmLanzamientos.frx":07C2
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Códigos de Productos"
            Top             =   90
            Width           =   405
         End
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "chk"
      Height          =   195
      Left            =   3045
      TabIndex        =   3
      Top             =   570
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4725
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   465
   End
   Begin VB.TextBox txtCantReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   480
   End
   Begin VB.Frame frCabecera 
      Height          =   945
      Left            =   165
      TabIndex        =   4
      Top             =   450
      Width           =   8265
      Begin VB.Frame frCab 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   135
         TabIndex        =   10
         Top             =   195
         Width           =   5625
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   4110
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   180
            Width           =   930
         End
         Begin VB.ComboBox cboAnio 
            Height          =   315
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   165
            Width           =   930
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
            Index           =   1
            Left            =   2895
            TabIndex        =   12
            Top             =   180
            Width           =   1080
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
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   180
            Width           =   1050
         End
      End
      Begin VB.Label de 
         Caption         =   "de"
         Height          =   255
         Left            =   5130
         TabIndex        =   5
         Top             =   135
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
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":08C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":0BDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":0EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":1212
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":152C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":1846
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":1B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":1E7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":2194
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":24AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":25C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":26D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":2C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":34A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":3C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":4526
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmLanzamientos.frx":4840
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAdmLanzamientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As Recordset

Dim ModoEdit As Integer
Dim DatoValido As Boolean

Dim cl_Productividad As CLConc

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$
Public Sub altas()
    SetearBotonesAltas True
    Modo = "ALTA"
    FrmAdmLanzamientos.caption = "Lanzamientos  -Altas- "
End Sub
Private Sub L_AltasDatos()

If L_TodoCargado Then
    FrmAdmLanzamientos.caption = Aplicacion.SeteoProceso(FrmAdmLanzamientos.caption)

    Aplicacion.ComienzoTrans

    MeLlenarObjeto

    If cl_Productividad.Insertar_Lanzamientos() Then
        Aplicacion.TerminarConExitoTrans
        chk.Value = 0
    
        NuevaSeleccion
    
    Else
        Aplicacion.TerminarConErrorTrans
    End If

    FrmAdmLanzamientos.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub
Private Sub MeImpDatos()
End Sub
Private Sub MePrepararAgregar()

    Tollbar.Buttons(1).Value = tbrPressed
    Tollbar.Buttons(2).Value = tbrUnpressed
    
    altas
    NuevaSeleccion
    
End Sub
Private Sub MePrepararAlterar()

    Tollbar.Buttons(2).Value = tbrPressed
    Tollbar.Buttons(1).Value = tbrUnpressed
    
    modificacion
    NuevaSeleccion
    
End Sub
Public Sub modificacion()

SetearBotonesAltas False
Modo = "MOD"
FrmAdmLanzamientos.caption = "Lanzamientos  -Modificacion y Bajas- "

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
Set cl_Productividad = New CLConc

spr.MaxRows = 0

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

FrmAdmLanzamientos.caption = Aplicacion.SeteoProceso(FrmAdmLanzamientos.caption)

Aplicacion.ComienzoTrans

MeLlenarObjeto

If cl_Productividad.Update_Lanzamientos Then
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

FrmAdmLanzamientos.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub
Private Sub MeCargarDatos()
Dim sql$
    
CondConsulta = ArmarCondicion

sqlGral$ = ""
sqlGral$ = sql$ & " SELECT aniomes, fch_desde,fch_hasta,c.cod_prod,descrip,OBJETIVO " _
& " FROM estadis.codigos_Lanzamientos c , baires.producto p" _
& " Where c.cod_prod = p.cod_prod " & CondConsulta
    
If Aplicacion.ObtenerRsDAO(sqlGral$, rs) Then
    'txtCantReg.Text = Aplicacion.CantReg(RS)
        If Aplicacion.CantReg(rs) = 0 Then
            txtCantReg.Text = 0
        Else
            txtCantReg.Text = 1
        End If

    If txtCantReg.Text > 0 Then
        txtReg.Text = 1
        SetBotonesGeneral True
        MellenarPantalla
        MeSetearBotonesToolBar
    Else
        txtReg.Text = 0
    End If
End If

End Sub
Private Sub MeEliminar()

If MsgBox("Esta seguro de eliminar el registro", vbYesNo + vbExclamation, "ATENCION") = vbYes Then

MeLlenarObjeto

FrmAdmLanzamientos.caption = Aplicacion.SeteoProceso(FrmAdmLanzamientos.caption)

Aplicacion.ComienzoTrans

If cl_Productividad.Delete_Lanzamientos Then
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

FrmAdmLanzamientos.caption = Aplicacion.SeteoFin
End If

End Sub
Private Sub MeLlenarObjeto()
Dim clProd As CLProdPrec
Dim i As Long, valor As Variant

cl_Productividad.Mes = cboAnio.Text & cboMes.Text
'cl_Productividad.fchDesde = mskFDesde.FormattedText
'cl_Productividad.fchHasta = mskFHasta.FormattedText

Set cl_Productividad.col_producto = New Collection

For i = 1 To spr.MaxRows

    spr.GetText 1, i, valor
    Set clProd = New CLProdPrec
    If valor <> "" And Spread_FilaOcupada(spr, i) Then
        clProd.codProd = valor
        spr.GetText 3, i, valor
        clProd.fdesde = valor
        spr.GetText 4, i, valor
        clProd.fhasta = valor
        spr.GetText 5, i, valor
        clProd.Objetivo = valor
        AdicionarAColeccion clProd
        
    End If

Next

End Sub
Private Sub AdicionarAColeccion(cod As CLProdPrec)
Dim clProd As CLProdPrec
Dim Resp As Boolean

Resp = True

For Each clProd In cl_Productividad.col_producto
    If cod.codProd = clProd.codProd Then
        Resp = False
        Exit For
    End If
Next

If Resp Then
    cl_Productividad.col_producto.Add cod
End If

End Sub
Private Function L_TodoCargado() As Boolean
    
'If mskFDesde.Text <> "" And mskFHasta.Text <> "" Then
    L_TodoCargado = True
'Else
'    L_TodoCargado = False
'End If

End Function
Private Sub MellenarPantalla()

spr.MaxRows = 0
If Not rs.EOF Then
    'mskFDesde.Text = Format(rs!fch_desde, FTOFECHA)
    'mskFHasta.Text = Format(rs!fch_hasta, FTOFECHA)

    Do While Not rs.EOF
        spr.MaxRows = spr.MaxRows + 1
        
        spr.SetText 1, spr.MaxRows, Trim(rs!Cod_prod)
        spr.SetText 2, spr.MaxRows, Trim(rs!Descrip)
        spr.SetText 3, spr.MaxRows, Format(rs!fch_desde, "dd/mm/yyyy")
        spr.SetText 4, spr.MaxRows, Format(rs!fch_hasta, "dd/mm/yyyy")
        spr.SetText 5, spr.MaxRows, str(rs!Objetivo)
        rs.MoveNext
    Loop
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

'habilitar frames
frCab.Enabled = Not valor
Spread.spread_LockGrilla spr, valor, 1, spr.MaxCols

    If Not valor Then
        txtReg.Text = 0
        txtCantReg.Text = 0
    End If

End Sub
Private Function ArmarCondicion()
Dim Con$

Con$ = ""

If cboAnio.Text <> "" Then
    Con$ = " And aniomes = " & cboAnio.Text & cboMes.Text
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
    
If Aplicacion.ObtenerRsDAO(sqlGral$, rs) Then
        
        'txtCantReg.Text = Aplicacion.CantReg(rs)
        If Aplicacion.CantReg(rs) = 0 Then
            txtCantReg.Text = 0
        Else
            txtCantReg.Text = 1
        End If
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

    frBot.Enabled = valor
    Spread.spread_LockGrilla spr, Not valor, 1, spr.MaxCols

    
End Sub
Private Sub SetearBotonesConsBases(valor As Boolean)
'valor = true -> altas
'valor = false -> modif
    
    Tollbar.Buttons(1).Visible = Not valor
    Tollbar.Buttons(2).Visible = Not valor
    
    Tollbar.Buttons(4).Visible = Not valor
    Tollbar.Buttons(15).Visible = Not valor
    Tollbar.Buttons(16).Visible = Not valor
    
    Tollbar.Buttons(17).Visible = Not valor 'False
    Tollbar.Buttons(18).Visible = Not valor 'False
    
    Tollbar.Buttons(5).Visible = Not valor 'False
    
    Tollbar.Buttons(7).Visible = valor  'False
    Tollbar.Buttons(8).Visible = valor  'False
    Tollbar.Buttons(9).Visible = valor  'False
    Tollbar.Buttons(10).Visible = valor  'False
    
    Tollbar.Buttons(12).Visible = Not valor 'False
    Tollbar.Buttons(13).Visible = Not valor 'False

    Tollbar.Buttons(18).Visible = Not valor 'False
    Tollbar.Buttons(19).Visible = Not valor 'False
    
    txtCantReg.Visible = Not valor 'False
    txtReg.Visible = Not valor 'False
    de.Visible = Not valor 'False

    frBot.Enabled = valor
    Spread.spread_LockGrilla spr, Not valor, 1, spr.MaxCols
    
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
Spread.spread_LockGrilla spr, valor, 1, spr.MaxCols
frBot.Enabled = Not valor
'botPlus.Enabled = Not valor
End Sub

Private Sub botHelpProd_Click()
Dim cod As String
Dim desc As String, CodProv As String
Dim sql As String

 
    If spr.MaxRows > 0 Then
        sql = "Select cod_prod,descrip from baires.producto "
        CodProv = "NA"
    If frmHelpProd.MuestraHlp(cod, desc, "Producto", sql, CodProv) = vbOK Then
       spr.SetText 1, spr.ActiveRow, cod
       spr.SetText 2, spr.ActiveRow, desc
       If spr.ActiveRow = spr.MaxRows Then
          Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
       End If
       DatoValido = True
    End If
    End If

spr.SetFocus

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
            Call TollBar_ButtonClick(Tollbar.Buttons(19))
    End Select
    If Modo = "MOD" And Val(txtCantReg.Text) > 0 Then
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

Top = 800
Left = 800
Width = 8800
Height = 5900

FrmAdmLanzamientos.caption = "Lanzamientos -Modificacion y Bajas- "

cboAnio.AddItem "2002"
cboAnio.AddItem "2003"
cboAnio.AddItem "2004"
cboAnio.AddItem "2005"

cboAnio.ListIndex = Year(Now) - 2002

cboMes.AddItem "01"
cboMes.AddItem "02"
cboMes.AddItem "03"
cboMes.AddItem "04"
cboMes.AddItem "05"
cboMes.AddItem "06"
cboMes.AddItem "07"
cboMes.AddItem "08"
cboMes.AddItem "09"
cboMes.AddItem "10"
cboMes.AddItem "11"
cboMes.AddItem "12"

cboMes.ListIndex = Month(Now) - 1

Modo = "MOD"

NuevaSeleccion

DatoValido = True

End Sub
Private Sub spr_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If Mode = 1 Then
    ModoEdit = Mode
End If
End Sub
Private Sub spr_KeyPress(KeyAscii As Integer)
Dim sql As String
Dim desc As String
Dim cod As Variant
Dim st As Variant

If KeyAscii = 13 And ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    spr.GetText 1, spr.ActiveRow, cod
    
    sql = "SELECT descrip "
    sql = sql & " FROM  baires.producto "
    sql = sql & " where cod_prod = " & cod

    Select Case spr.ActiveCol
        Case 1
            DatoValido = Func_ObtenerDesc(sql, desc)
            spr.SetText 2, spr.ActiveRow, desc
            If DatoValido Then
                spr.Col = 3
                spr.Row = spr.ActiveRow
                spr.Action = 0
                spr.Action = 1
            End If
        Case 3
                spr.Col = 4
                spr.Row = spr.ActiveRow
                spr.Action = 0
                spr.Action = 1
        Case 4
                spr.Col = 5
                spr.Row = spr.ActiveRow
                spr.Action = 0
                spr.Action = 1
        Case 5
            spr.GetText 3, spr.ActiveRow, st
            DatoValido = Func_ObtenerDesc(sql, desc)
            If DatoValido And spr.ActiveRow = spr.MaxRows Then
                Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                spr.Col = 1
                spr.Row = spr.MaxRows
                spr.Action = 0
                spr.Action = 1
            End If

    End Select
            spr.TopRow = spr.TopRow - 6
End If

End Sub
Private Sub spr_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim sql As String
Dim desc As String
Dim cod As Variant
Dim st As Variant

If ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    spr.GetText 1, Row, cod
    
    sql = "SELECT descrip "
    sql = sql & " FROM  baires.producto "
    sql = sql & " where cod_prod = " & cod
    
    Select Case spr.ActiveCol
        Case 1
            DatoValido = Func_ObtenerDesc(sql, desc)
            spr.SetText 2, spr.ActiveRow, desc
        Case 2
        Case 3
            spr.GetText 3, spr.ActiveRow, st
            DatoValido = Func_ObtenerDesc(sql, desc)
    End Select
    
End If

Cancel = Not DatoValido

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
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim valor As Variant

Select Case Button.Key
    Case "A"
        If Spread_FilaOcupada(spr, spr.MaxRows) Then
           Spread_AddRow spr
        End If
    Case "B"
        Spread_DelOneRow spr, spr.ActiveRow
        DatoValido = True
    Case "C"
        spr.MaxRows = 0
    Case "D"
        
End Select
End Sub

Private Sub vaSpread1_Advance(ByVal AdvanceNext As Boolean)

End Sub


Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If Mode = 1 Then
    ModoEdit = Mode
End If
End Sub


Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
Dim sql As String
Dim desc As String
Dim cod As Variant
Dim st As Variant

If KeyAscii = 13 And ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    spr.GetText 1, spr.ActiveRow, cod
    
    sql = "SELECT descrip "
    sql = sql & " FROM  baires.producto "
    sql = sql & " where cod_prod = " & cod

    Select Case spr.ActiveCol
        Case 1
            DatoValido = Func_ObtenerDesc(sql, desc)
            spr.SetText 2, spr.ActiveRow, desc
            If DatoValido Then
                spr.Col = 3
                spr.Row = spr.ActiveRow
                spr.Action = 0
                spr.Action = 1
            End If
        Case 3
                spr.Col = 4
                spr.Row = spr.ActiveRow
                spr.Action = 0
                spr.Action = 1
        Case 4
                spr.Col = 5
                spr.Row = spr.ActiveRow
                spr.Action = 0
                spr.Action = 1
        Case 5
            spr.GetText 3, spr.ActiveRow, st
            DatoValido = Func_ObtenerDesc(sql, desc)
            If DatoValido And spr.ActiveRow = spr.MaxRows Then
                Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                spr.Col = 1
                spr.Row = spr.MaxRows
                spr.Action = 0
                spr.Action = 1
            End If

    End Select
            spr.TopRow = spr.TopRow - 6
End If

End Sub


Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim sql As String
Dim desc As String
Dim cod As Variant
Dim st As Variant

If ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    spr.GetText 1, Row, cod
    
    sql = "SELECT descrip "
    sql = sql & " FROM  baires.producto "
    sql = sql & " where cod_prod = " & cod
    
    Select Case spr.ActiveCol
        Case 1
            DatoValido = Func_ObtenerDesc(sql, desc)
            spr.SetText 2, spr.ActiveRow, desc
        Case 2
        Case 3
            spr.GetText 3, spr.ActiveRow, st
            DatoValido = Func_ObtenerDesc(sql, desc)
    End Select
    
End If

Cancel = Not DatoValido


End Sub


