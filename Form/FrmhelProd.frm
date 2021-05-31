VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmHelpProd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar "
   ClientHeight    =   4800
   ClientLeft      =   1500
   ClientTop       =   1665
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4800
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4665
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   4725
      Begin FPSpread.vaSpread Spread1 
         Bindings        =   "FrmhelProd.frx":0000
         Height          =   2655
         Left            =   180
         OleObjectBlob   =   "FrmhelProd.frx":0010
         TabIndex        =   1
         Top             =   1365
         Width           =   4005
      End
      Begin VB.ComboBox cboDos 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   180
         Width           =   2370
      End
      Begin VB.TextBox txtDos 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   3915
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   195
         Width           =   690
      End
      Begin VB.TextBox txtUno 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   585
         Width           =   690
      End
      Begin VB.ComboBox cboUno 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   570
         Width           =   2370
      End
      Begin VB.ListBox lstUno 
         Height          =   255
         Left            =   -300
         TabIndex        =   8
         Top             =   210
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.ListBox lstDos 
         Height          =   255
         Left            =   -465
         TabIndex        =   7
         Top             =   585
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Timer tmrBusca 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   3960
         Top             =   -15
      End
      Begin VB.CommandButton botAceptar 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   2310
         TabIndex        =   3
         Top             =   4200
         Width           =   810
      End
      Begin VB.CommandButton botCerrar 
         Caption         =   "&Cerrar"
         Height          =   330
         Left            =   3210
         TabIndex        =   2
         Top             =   4200
         Width           =   810
      End
      Begin MSMask.MaskEdBox mskCodigo 
         Height          =   300
         Left            =   1485
         TabIndex        =   5
         Top             =   960
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   529
         _Version        =   327680
         Format          =   ">"
         PromptChar      =   " "
      End
      Begin VB.Frame Frame2 
         Height          =   600
         Left            =   135
         TabIndex        =   4
         Top             =   4005
         Width           =   4110
         Begin VB.CommandButton botTodos 
            Caption         =   "&Todos"
            Height          =   330
            Left            =   1260
            TabIndex        =   12
            Top             =   195
            Width           =   810
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   180
         TabIndex        =   11
         Top             =   570
         Width           =   1245
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
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmHelpProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim codigo$
Dim Descripcion$
Dim Resp%
Dim lProv As String

Dim Tabla As String
Dim Seleccion As String

Private Function ArmarSQL() As String
Dim s_Criterio As String
Dim pos As Integer

Select Case Tabla
    Case "Producto"
        If mskCodigo.Text <> "" Then
            s_Criterio = " and descrip LIKE  '" + Trim$(mskCodigo.Text) + "%'"
        Else
            s_Criterio = ""
        End If
        If txtUno.Text <> "" Then
            s_Criterio = s_Criterio & " and cod_marca = '" & txtUno.Text & "' "
        End If
        If txtDos.Text <> "" Then
            s_Criterio = s_Criterio & " and cod_rubr = '" & txtDos.Text & "' "
        End If
        If lProv = "NA" Then
            If s_Criterio <> "" Then
                s_Criterio = " Where " & Right(s_Criterio, Len(s_Criterio) - 4)
            End If
        End If
        ArmarSQL = Seleccion & s_Criterio & " ORDER BY cod_prod "

    Case "Proveedor"
        If mskCodigo.Text <> "" Then
            s_Criterio = " Where descrip LIKE  '" + Trim$(mskCodigo.Text) + "%'"
        Else
            s_Criterio = ""
        End If
        
        ArmarSQL = Seleccion & s_Criterio & " " _
        & " UNION SELECT 'N' as cod_prov, 'DISCONTINUADOS' as Descrip FROM Dual " _
        & " "
        

End Select

End Function


Private Function CodigoResultado(cod As String) As String

Select Case Tabla
    Case "ESPACIOS"
    Case Else
        CodigoResultado = cod
End Select

End Function

Public Function NombreProducto() As String
Spread1.col = 2
NombreProducto = Spread1.Text
End Function

Private Sub botAceptar_Click()
  
 Spread1.col = 1
 Spread1.Row = Spread1.ActiveRow
  
 codigo$ = CodigoResultado(Spread1.Text)
 
 Spread1.col = 2
 Descripcion$ = Spread1.Text
 
 
 If codigo$ <> "" Then
  Resp% = vbOK
 Else
   Resp% = vbCancel
 End If
 
 Unload Me
 
 DoEvents

End Sub

Private Sub botCerrar_Click()
   
  Resp% = vbCancel
   
  Unload Me
  
  DoEvents

End Sub

Public Sub CargarHelp()
Dim sql$, rs As Recordset

'Screen.MousePointer = vbDefault

'sql$ = "select cod_rubr, descrip from Rubro ORDER BY cod_rubr"

'Call Aplicacion.ObtenerRsDAO(sql$, rs)

'If rs.RecordCount > 0 Then
'  DoEvents
'  Spread_CargarGrillaGauge rs, Spread1, FrmPross.prossbar
'Else
'    Spread1.MaxRows = 0
'End If

'DoEvents

'  mskCodigo.Text = ""

End Sub

Private Sub botTodos_Click()
Dim cod As String
Dim Descrip As String
Dim i

For i = 1 To Spread1.MaxRows
 Spread1.col = 1
 Spread1.Row = i
  
 cod = CodigoResultado(Spread1.Text)
 
 Spread1.col = 2
 Descrip = Spread1.Text

 Select Case Tabla
  Case "Producto"
    FrmAdmConcurso.PonerValores cod, Descrip
  Case "Proveedor"
    FrmAdmPxC.PonerValores cod, Descrip
 End Select
  
 
Next

End Sub


Private Sub cboDos_Click()
Dim sql As String

txtDos.Text = lstDos.List(cboDos.ListIndex)
'Select Case Tabla
'    Case "Producto"
'
'End Select
  
tmrBusca.Enabled = False
  
tmrBusca.Enabled = True

End Sub

Private Sub cboDos_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
 cboDos.ListIndex = -1
End If
End Sub

Private Sub cboUno_Click()
Dim sql As String

txtUno.Text = lstUno.List(cboUno.ListIndex)
'Select Case Tabla
'    Case "Producto"
'
'End Select
  
tmrBusca.Enabled = False
  
tmrBusca.Enabled = True

End Sub


Private Sub cboUno_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
cboUno.ListIndex = -1
End If
End Sub

Private Sub Form_Activate()
    tmrBusca.Enabled = False
    mskCodigo.SetFocus

    frmHelpPv.caption = frmHelpPv.caption & " " & Tabla
    
End Sub

Private Sub Form_Load()
Dim sql As String

Select Case Tabla
    Case "Producto"
        Label1(3).caption = "Marcas"
        Label1(1).caption = "Rubros"
        If lProv = "NA" Then
            sql = " SELECT cod_marca, descrip FROM baires.marca WHERE cod_prov = '" & lProv & "' "
        Else
            sql = " SELECT cod_marca, descrip FROM baires.marca  "
        End If
        FuncCbos_LlenarCboLst cboUno, lstUno, sql
        sql = " SELECT cod_rubr, descrip FROM baires.rubro "
        FuncCbos_LlenarCboLst cboDos, lstDos, sql
        If lProv = "NA" Then
            cboUno.Enabled = False
        End If
    
    Case "Proveedor"
        Label1(3).Visible = False
        Label1(1).Visible = False
        cboUno.Visible = False
        txtUno.Visible = False
        cboDos.Visible = False
        txtDos.Visible = False
        Label1(0).Top = Label1(0).Top - 500
        mskCodigo.Top = mskCodigo.Top - 500
        Spread1.Top = Spread1.Top - 450
        Spread1.Height = Spread1.Height + 400
End Select


End Sub


Private Sub mskCodigo_Change()
  tmrBusca.Enabled = False
  
  tmrBusca.Enabled = True

End Sub




Private Sub Spread1_DblClick(ByVal col As Long, ByVal Row As Long)
Dim cod As String
Dim Descrip As String

 Spread1.col = 1
 Spread1.Row = Spread1.ActiveRow
  
 cod = CodigoResultado(Spread1.Text)
 
 Spread1.col = 2
 Descrip = Spread1.Text

 Select Case Tabla
  Case "Producto"
    If lProv <> "NA" Then
        FrmAdmConcurso.PonerValores cod, Descrip
    Else
        FrmAdmProductividad.PonerValores cod, Descrip
    End If
  Case "Proveedor"
    FrmAdmPxC.PonerValores cod, Descrip
 End Select

End Sub


Private Sub tmrBusca_Timer()
Dim s_Criterio As String
Dim sql$, rs As Recordset

frmHelpProd.caption = Aplicacion.SeteoProceso(frmHelpProd.caption)

sql$ = ArmarSQL

Call Aplicacion.ObtenerRsDAO(sql$, rs)

If rs.RecordCount > 0 Then
  Spread_CargarGrilla rs, Spread1
Else
  Spread1.MaxRows = 0
End If

Aplicacion.CerrarDAO rs
 
tmrBusca.Enabled = False

frmHelpProd.caption = Aplicacion.SeteoFin()

End Sub

Public Function MuestraHlp(ByRef cod As Variant, ByRef desc As String, ByVal NT As String, ByVal CA As String, Pv As String) As Integer
    
    Tabla = NT
    Seleccion = CA
    lProv = Pv
    Me.Show 1
    
    cod = codigo$
    desc = Descripcion$
    
    MuestraHlp = Resp%
    
End Function

