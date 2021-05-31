VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form FrmAdmDocumento 
   Caption         =   "Administración de"
   ClientHeight    =   5520
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   6675
   Icon            =   "FrmAdmDocumento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   6675
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   120
      TabIndex        =   5
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
            Object.Visible         =   0   'False
            Key             =   "f"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   135
      TabIndex        =   9
      Top             =   1710
      Width           =   6390
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   420
         Left            =   300
         TabIndex        =   10
         Top             =   105
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   741
         ButtonWidth     =   635
         ButtonHeight    =   582
         Appearance      =   1
         ImageList       =   "ImageList2"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   7
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "A"
               Object.ToolTipText     =   "Agreagar Fila"
               Object.Tag             =   ""
               ImageIndex      =   1
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
               ImageIndex      =   2
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
               ImageIndex      =   4
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
               ImageIndex      =   3
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin FPSpread.vaSpread spr 
         Height          =   2745
         Left            =   165
         OleObjectBlob   =   "FrmAdmDocumento.frx":0442
         TabIndex        =   14
         Top             =   585
         Width           =   6105
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "chk"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   390
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4185
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   420
      Width           =   465
   End
   Begin VB.TextBox txtCantReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   420
      Width           =   480
   End
   Begin VB.Frame frCab 
      Height          =   1215
      Left            =   135
      TabIndex        =   7
      Top             =   480
      Width           =   6375
      Begin MSMask.MaskEdBox mskPedido 
         Height          =   330
         Left            =   1710
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   765
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   4230
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   315
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtTicket 
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Doc."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3090
         TabIndex        =   15
         Top             =   750
         Width           =   2805
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pedido Asociado"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   13
         Top             =   765
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3090
         TabIndex        =   12
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   11
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label de 
         Caption         =   "de"
         Height          =   255
         Left            =   4620
         TabIndex        =   8
         Top             =   120
         Width           =   405
      End
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":07A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":0AC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":0DDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":160F
            Key             =   ""
         EndProperty
      EndProperty
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
            Picture         =   "FrmAdmDocumento.frx":1929
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":1C43
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":1F5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":2277
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":2591
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":28AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":2BC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":2EDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":31F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":3513
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":3625
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":3737
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmDocumento.frx":3CD9
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAdmDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As Recordset

Dim cl_Ped As CLTicket

Dim CondConsulta As String

Dim Modo As String
Dim Tipo As Integer

Dim sqlGral$

Dim ModoEdit As Integer
Dim DatoValido As Boolean

Public Sub Altas(pTipo As Integer)
    Tipo = pTipo
    Modo = ALTA
    FrmAdmDocumento.caption = FrmAdmDocumento.caption & " - ALTAS -"
    SetearBotonesAltas
    Me.Show 1
    
End Sub
Private Sub L_AltasDatos()
If L_TodoCargado Then
    FrmAdmDocumento.caption = Aplicacion.SeteoProceso(FrmAdmDocumento.caption)

    Aplicacion.ComienzoTrans

    MeLlenarObjeto

    If cl_Ped.Insert_Doc() Then
        Aplicacion.TerminarConExitoTrans
        chk.Value = 0

        NuevaSeleccion

    Else
        Aplicacion.TerminarConErrorTrans
    End If

    FrmAdmDocumento.caption = Aplicacion.SeteoFin
Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub L_PedidoAutomatico()
Dim sql As String
Dim rs As Recordset
Dim desc As String
Dim nro As Integer

desc = ""

nro = 0

If mskPedido.Text <> 0 Then
    sql = " Select nro_doc From documentos_C "
    sql = sql & " Where pedido_asociado = " & mskPedido.Text
    sql = sql & " And tipo_doc = 'R' "
    If Aplicacion.ObtenerRs(sql, rs, g_Tipo) Then
        If Not rs.EOF Then
            nro = rs!nro_doc
        End If
    End If
    rs.Close
End If

If nro = 0 Then
    sql = " Select P.producto_id,descrip, cantidad "
    sql = sql & " From PedidoS_D d, producto P "
    sql = sql & " Where p.producto_id = d.producto_id "
    sql = sql & " And d.nro_pedido = " & mskPedido.Text
    sql = sql & " And d.sucursal = " & sucursal
    
    If Aplicacion.ObtenerRs(sql, rs, g_Tipo) Then
        spr.MaxRows = 0
        Do While Not rs.EOF
            If rs!Cantidad > 0 Then
                spr.MaxRows = spr.MaxRows + 1
                                
                spr.SetText 1, spr.MaxRows, str(rs!Cantidad)
                spr.SetText 2, spr.MaxRows, str(rs!Producto_id)
                spr.SetText 3, spr.MaxRows, Trim(rs!Descrip)
                
            End If
            rs.MoveNext
        Loop
    
    End If
Else
    spr.MaxRows = 0
    MsgBox "Pedido asociado al remito " & str(nro), vbCritical + vbOKOnly, "Atención"
End If

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
Modo = MODIF
FrmAdmDocumento.caption = FrmAdmDocumento.caption & " - Modificacion y Bajas -"
Frame1.Enabled = False

Me.Show 1


End Sub
Private Sub NuevaSeleccion()
Dim i%

If Modo = MODIF Then
    SetBotonesGeneral False
    txtTicket.Text = ""
    mskFecha.Text = ""
Else
    txtTicket.Text = Func_ObtenerNumero("ControlDoc")
    If chk.Value = 1 Then
        If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
            L_AltasDatos
        End If
    End If
End If
'Limpiar campos de pantallas
'''Set cl_CxM = New CCliente


spr.MaxRows = 0
mskPedido.Text = 0

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

'If L_TodoCargado Then
'
'FrmAdmDocumento.caption = Aplicacion.SeteoProceso(FrmAdmDocumento.caption)
'
'Aplicacion.ComienzoTrans
'
'MeLlenarObjeto
'
'If cl_CxM.Update_CxM Then '
'    Aplicacion.TerminarConExitoTrans
'    SeteoBotonesMod True
'
'    If MeReconsultar > 0 Then
'
'    Tollbar.Buttons(2).Enabled = False
'
'    MeSetearBotonesToolBar
'    Else
'            NuevaSeleccion
'    End If
'
'Else
'    Aplicacion.TerminarConErrorTrans
'End If

'FrmAdmDocumento.caption = Aplicacion.SeteoFin
'
'Else
'    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
'End If

End Sub

Private Sub MeCargarDatos()
Dim sql$

FrmAdmDocumento.caption = Aplicacion.SeteoProceso(FrmAdmDocumento.caption)
        
    CondConsulta = ArmarCondicion

    sqlGral$ = ""
    sqlGral$ = sqlGral$ & " SELECT C.Nro_doc,fecha,pedido_asociado "
    sqlGral$ = sqlGral$ & " FROM  documentos_C C "
    sqlGral$ = sqlGral$ & "  "
    sqlGral$ = sqlGral$ & " "
    sqlGral$ = sqlGral$ & CondConsulta
    
If Aplicacion.ObtenerRs(sqlGral$, rs, g_Tipo) Then
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

FrmAdmDocumento.caption = Aplicacion.SeteoFin

End Sub
Private Sub MeEliminar()

'If MsgBox("Esta seguro de eliminar el registro", vbYesNo + vbExclamation, "ATENCION") = vbYes Then
'
'MeLlenarObjeto
'
'FrmAdmDocumento.caption = Aplicacion.SeteoProceso(FrmAdmDocumento.caption)
'
'Aplicacion.ComienzoTrans
'
'If cl_CxM.Delete_CxM Then
'    Aplicacion.TerminarConExitoTrans
'    SeteoBotonesMod True
'
'    If MeReconsultar > 0 Then
'
'    Tollbar.Buttons(2).Enabled = False
'
'    MeSetearBotonesToolBar
'    Else
'            NuevaSeleccion
'    End If
'
'Else
'    Aplicacion.TerminarConErrorTrans
'End If

'FrmAdmDocumento.caption = Aplicacion.SeteoFin

'End If

End Sub

Private Sub MeLlenarObjeto()
Dim cl_Prod As CLTicketD
Dim i As Long, valor As Variant

cl_Ped.NroTicket = txtTicket.Text
cl_Ped.Fecha = mskFecha.FormattedText
cl_Ped.NroAsociado = mskPedido.Text
cl_Ped.Tipo_doc = "R"

Set cl_Ped.col_Item = New Collection

For i = 1 To spr.MaxRows
    Set cl_Prod = New CLTicketD '

    spr.GetText 2, i, valor
    If valor <> 0 And Spread_FilaOcupada(spr, i) Then
        cl_Prod.CodProd = valor
        
        spr.GetText 1, i, valor
        cl_Prod.Cant = valor
        
        AdicionarAColeccion cl_Prod
        
    End If

Next


End Sub


Private Sub AdicionarAColeccion(cl As CLTicketD)
Dim Item As CLTicketD
Dim Resp As Boolean

Resp = True

For Each Item In cl_Ped.col_Item
    If Item.CodProd = cl.CodProd Then
        Item.Cant = Item.Cant + cl.Cant
        Resp = False
        Exit For
    End If
Next

If Resp Then
    cl_Ped.col_Item.Add cl
End If

End Sub


Private Function L_TodoCargado() As Boolean
    L_TodoCargado = False
    
    If mskFecha.Text <> "" Then
        L_TodoCargado = True
    End If
    
End Function

Private Sub MellenarPantalla()
Dim desc As String
Dim sql As String
Dim rsDet As Recordset

txtTicket.Text = rs!nro_doc
mskFecha.Text = Format(rs!Fecha, FTOFECHA)
mskPedido.Text = rs!pedido_asociado

sql = "Select p.producto_id, descrip, cantidad " _
& " From Documentos_D D , Producto P " _
& " Where d.producto_id = p.producto_id And " _
& " sucursal = " & sucursal & " " _
& " And nro_doc = " & rs!nro_doc & " " _
& "  "

If Aplicacion.ObtenerRs(sql, rsDet, g_Tipo) Then
    spr.MaxRows = 0
        
    Do While Not rsDet.EOF
    
        spr.MaxRows = spr.MaxRows + 1
    
        spr.SetText 1, spr.MaxRows, str(rsDet!Cantidad)
        spr.SetText 2, spr.MaxRows, str(rsDet!Producto_id)
        spr.SetText 3, spr.MaxRows, Trim(rsDet!Descrip)
    
        rsDet.MoveNext
    Loop
End If

End Sub

Public Sub PonerValores(cod As Variant, desc As String)
Dim sql As String, descM As String

    spr.SetText 1, spr.MaxRows, Trim(cod)
    spr.SetText 2, spr.MaxRows, desc
        
        sql = "SELECT codigo_id FROM merch_x_mercado where mercado_id = " & cod
        Func_ObtenerDesc sql, descM
        spr.SetText 3, spr.MaxRows, descM

'    If spr.ActiveRow = spr.MaxRows Then
       Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
'    End If
    spr.Row = spr.MaxRows
    spr.Col = 1
    spr.Position = SS_POSITION_UPPER_LEFT
    spr.Action = SS_ACTION_GOTO_CELL

DatoValido = True
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
'    TollBar.Buttons(12).Enabled = Not valor
'    TollBar.Buttons(13).Enabled = Not valor
'

'habilitar frames
frCab.Enabled = Not valor
'Frame1.Enabled = valor

    If Not valor Then
        txtReg.Text = 0
        txtCantReg.Text = 0
    End If

End Sub

Private Function ArmarCondicion()
Dim Con$

Con$ = ""

If txtTicket.Text <> "" Then
    Con$ = Con$ & " and nro_doc = " & txtTicket.Text & " "
End If

If mskFecha.Text <> "" Then
    Con$ = Con$ & " and fecha_para = datevalue('" & mskFecha.FormattedText & "') "
End If

If mskPedido.Text <> "" Then
    Con$ = Con$ & " and pedido_asociado = " & mskPedido.Text & " "
End If

If Con$ <> "" Then
    Con$ = " WHERE " & Mid(Con$, 5, Len(Con$))
End If

ArmarCondicion = Con$

End Function



Private Sub MePrepararMod()
    
    SeteoBotonesMod False
    DatoValido = True
End Sub

Private Function MeReconsultar() As Integer
Dim sql$
Dim i%
    

If Aplicacion.ObtenerRs(sqlGral$, rs, g_Tipo) Then
        txtCantReg.Text = Aplicacion.CantReg(rs)
        If val(txtReg.Text) > val(txtCantReg.Text) Then
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
Frame1.Enabled = Not valor

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
            If Modo = MODIF Then
            Call Tollbar_ButtonClick(Tollbar.Buttons(2))
            End If
        Case 83 'Salir
            Call Tollbar_ButtonClick(Tollbar.Buttons(17))
        Case 107
            Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
        Case 109
            Call Toolbar1_ButtonClick(Toolbar1.Buttons(3))
    
    End Select
    If Modo = MODIF And val(txtCantReg.Text) > 0 Then
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
Else
'    Select Case KeyCode
'        Case 107
'            Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
'        Case 109
'            Call Toolbar1_ButtonClick(Toolbar1.Buttons(3))
 '   End Select

End If

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
Width = 6400
Height = 5800

Set cl_Ped = New CLTicket

mskPedido.Text = 0
If Modo = ALTA Then
    mskFecha.Text = Format$(Date + 1, FTOFECHA)
    
    txtTicket.Text = Func_ObtenerNumero("ControlDoc")
    'L_PedidoAutomatico
Else
    txtTicket.Locked = False
End If

DatoValido = True

If Tipo = 1 Then
    Label4.caption = "INGRESOS"
Else
    Label4.caption = "BAJAS/SALIDAS"
    spr.Col = 4
    spr.Row = 0
    spr.Col2 = 4
    spr.Row2 = 0
    spr.BlockMode = True
    spr.ColHidden = False
End If
'sql = " SELECT documento_id, descrip FROM tipo_documento "
'FuncCbos_LlenarCboiTEM cboDoc, sql
'cboDoc.ListIndex = 0



End Sub
Private Sub mskFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And mskPedido.Enabled Then
    mskPedido.SetFocus
End If
End Sub

Private Sub mskFecha_LostFocus()

If mskFecha.Text <> "" Then
If Not IsDate(mskFecha.FormattedText) Then
    mskFecha.Text = Format$(Date, FTOFECHA)
End If
    
mskFecha.Text = Format$(mskFecha.FormattedText, FTOFECHA)

'If Modo = ALTA Then
'    L_PedidoAutomatico
'End If
End If
End Sub

Private Sub mskPedido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And spr.Enabled Then
    spr.SetFocus
End If
End Sub

Private Sub mskPedido_LostFocus()


If Modo = ALTA Then
'Buscar el pedido asociado
    L_PedidoAutomatico
End If

End Sub


Private Sub spr_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If Mode = 1 Then
    ModoEdit = Mode
End If
End Sub


Private Sub spr_KeyPress(KeyAscii As Integer)
Dim sql As String
Dim desc As String, Precio As Single
Dim cod As Variant
Dim Cant As Variant, prec As Variant
Dim Tot As Single


If KeyAscii = 13 And ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    spr.GetText 2, spr.ActiveRow, cod
    
    sql = "SELECT descrip,precio "
    sql = sql & " FROM producto P "
    sql = sql & " where P.producto_id = " & cod

    Select Case spr.ActiveCol
        Case 2
            DatoValido = L_ObtenerDesc(sql, desc, Precio)
            spr.SetText 3, spr.ActiveRow, desc
            spr.SetText 4, spr.ActiveRow, str(Precio)
            spr.SetText 5, spr.ActiveRow, ""
            
            If DatoValido Then
                If Tipo = 1 Then
                    spr.GetText 1, spr.ActiveRow, Cant
                    spr.GetText 4, spr.ActiveRow, prec
                    If spr.ActiveRow = spr.MaxRows Then
                        Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                    End If
                    spr.Col = 1
                    spr.Row = spr.ActiveRow + 1
                    spr.Action = 0
                    'spr.Action = 1
                
                Else
                    spr.Col = 4
                    spr.Row = spr.ActiveRow
                    spr.Action = 0
                
                End If
    
                
            End If

        Case 1
            'If Spread_FilaOcupada(spr, spr.ActiveRow) Then
            '    spr.GetText 1, spr.ActiveRow, Cant
            '    spr.GetText 4, spr.ActiveRow, prec
            '
            '    spr.SetText 5, spr.ActiveRow, str(Cant * prec)
            '
            '    'mskTotal.Text = Format$(Tot + (Cant * prec), "#,##0.00")
            'End If
           ' SendKeys "{Right}"
                spr.Col = 2
                spr.Row = spr.ActiveRow
                spr.Action = 0

        Case 4
            If Spread_FilaOcupada(spr, spr.ActiveRow) Then
                If spr.ActiveRow = spr.MaxRows Then
                    Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                End If
                    spr.Col = 1
                    spr.Row = spr.ActiveRow + 1
                    spr.Action = 0
            
            End If
        
    End Select
    
    
ElseIf KeyAscii = 104 And spr.ActiveCol = 2 Then
    sql = "SELECT producto_id,descrip "
    sql = sql & " FROM producto P "

    frmHelp.MuestraHlp cod, desc, "producto", sql
    spr.SetText 2, spr.ActiveRow, str(cod)
    desc = ""

End If


End Sub


Private Function L_ObtenerDesc(sql As String, ByRef desc As String, ByRef Precio As Single) As Boolean
Dim rs As Recordset

If Aplicacion.ObtenerRs(sql$, rs, g_Tipo) Then
    If Aplicacion.CantReg(rs) > 0 Then
        L_ObtenerDesc = True
        desc = rs.Fields(0).Value
        Precio = rs.Fields(1).Value
    Else
        L_ObtenerDesc = False
        desc = ""
        Precio = 0
    End If
    Aplicacion.CerrarDAO rs
End If

End Function


Private Sub spr_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim sql As String
Dim desc As String, Precio As Single
Dim cod As Variant
Dim Cant As Variant, prec As Variant


If ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    spr.GetText 1, Row, cod
    
    sql = "SELECT descrip, precio "
    sql = sql & " FROM producto P "
    sql = sql & " where P.producto_id = " & cod
    
    Select Case spr.ActiveCol
        Case 2
            DatoValido = L_ObtenerDesc(sql, desc, Precio)
            spr.SetText 3, spr.ActiveRow, desc
            spr.SetText 4, spr.ActiveRow, str(Precio)
            spr.SetText 5, spr.ActiveRow, ""
            If DatoValido Then
                spr.GetText 1, spr.ActiveRow, Cant
                spr.GetText 4, spr.ActiveRow, prec
                
                spr.SetText 5, spr.ActiveRow, str(Cant * prec)
                'mskTotal.Text = Format$(mskTotal.Text + (Cant * prec), "#,##0.00")
            End If
        Case 1
            If Spread_FilaOcupada(spr, spr.ActiveRow) Then
                spr.GetText 1, spr.ActiveRow, Cant
                spr.GetText 4, spr.ActiveRow, prec
                
                spr.SetText 5, spr.ActiveRow, str(Cant * prec)
                'mskTotal.Text = Format$(mskTotal.Text + (Cant * prec), "#,##0.00")
            End If
            'SendKeys "{Right}"
    End Select
    
End If

Cancel = Not DatoValido


End Sub

Private Sub Tollbar_ButtonClick(ByVal Button As Button)
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
        If Modo = MODIF Then
            MeActualizar
        Else
            L_AltasDatos
        End If
    Case "i"
        MeAbortarMod
    Case "j"
        If chk.Value = 1 Then
            If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
                If Modo = MODIF Then
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
    
End Select

If Not saltear Then
    txtReg.Text = pos
    MellenarPantalla
    MeSetearBotonesToolBar
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
'        LlenarColeccion
'        Unload Me
        
End Select

End Sub





Private Sub txtTicket_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And mskFecha.Enabled Then
    mskFecha.SetFocus
End If

End Sub


