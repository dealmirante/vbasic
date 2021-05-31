VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmHelpPv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar "
   ClientHeight    =   4800
   ClientLeft      =   1500
   ClientTop       =   1665
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4800
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox VideoSoftElastic1 
      Height          =   4575
      Left            =   45
      ScaleHeight     =   4515
      ScaleWidth      =   4395
      TabIndex        =   4
      Top             =   105
      Width           =   4455
      Begin FPSpread.vaSpread Spread1 
         Bindings        =   "Frmhelpr.frx":0000
         Height          =   2865
         Left            =   195
         OleObjectBlob   =   "Frmhelpr.frx":0010
         TabIndex        =   1
         Top             =   645
         Width           =   4005
      End
      Begin VB.CommandButton botCerrar 
         Caption         =   "&Cerrar"
         Height          =   330
         Left            =   3225
         TabIndex        =   3
         Top             =   3825
         Width           =   810
      End
      Begin VB.CommandButton botAceptar 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   2325
         TabIndex        =   2
         Top             =   3825
         Width           =   810
      End
      Begin VB.Timer tmrBusca 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   3975
         Top             =   0
      End
      Begin VB.Frame Frame1 
         Height          =   690
         Left            =   135
         TabIndex        =   6
         Top             =   3600
         Width           =   4110
         Begin VB.OptionButton optTipo 
            Caption         =   "Esp"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   855
            TabIndex        =   8
            Top             =   255
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Std"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   7
            Top             =   255
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin MSMask.MaskEdBox mskCodigo 
         Height          =   300
         Left            =   1500
         TabIndex        =   0
         Top             =   240
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   529
         _Version        =   327680
         Format          =   ">"
         PromptChar      =   " "
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
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmHelpPv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim codigo$
Dim Descripcion$
Dim Resp%

Dim Tabla As String
Dim Seleccion As String
Private Function ArmarSQL() As String
Dim s_Criterio As String
Dim pos As Integer

Select Case Tabla
    Case "CAJEROS"
        If mskCodigo.Text <> "" Then
            s_Criterio = " WHERE apellido LIKE  '" + Trim$(mskCodigo.Text) + "%'"
        Else
            s_Criterio = ""
        End If
        ArmarSQL = Seleccion & s_Criterio & " ORDER BY legajo"
    Case "PROVEEDOR"
        If mskCodigo.Text <> "" Then
            s_Criterio = " WHERE descrip LIKE  '" + Trim$(mskCodigo.Text) + "%'"
        Else
            s_Criterio = ""
        End If
        ArmarSQL = Seleccion & s_Criterio & " ORDER BY cod_prov"
    Case "MARCA"
        If mskCodigo.Text <> "" Then
            s_Criterio = " and descrip LIKE  '" + Trim$(mskCodigo.Text) + "%' "
        Else
            s_Criterio = ""
        End If
        ArmarSQL = Seleccion & s_Criterio & " ORDER BY cod_marca"
    Case "PRODUCTO"
        If mskCodigo.Text <> "" Then
            s_Criterio = " AND descrip LIKE  '" + Trim$(mskCodigo.Text) + "%'"
        Else
            s_Criterio = ""
        End If
        ArmarSQL = Seleccion & s_Criterio & " ORDER BY cod_prod"
        'ArmarSQL = "Select cod_prod, descrip from proveedor " & s_Criterio & " ORDER BY cod_prod "
    Case "ESPACIOS"
        If mskCodigo.Text <> "" Then
            s_Criterio = " and nombre LIKE  '" + Trim$(mskCodigo.Text) + "%' "
        Else
            s_Criterio = ""
        End If
        pos = 32
        If optTipo(0).Value Then
            Mid(Seleccion, pos, 16) = " esp_standar   "
        Else
            Mid(Seleccion, pos, 16) = " esp_especial  "
        End If

        ArmarSQL = Seleccion & s_Criterio & " ORDER BY id_espacio "
End Select

End Function


Private Function CodigoResultado(cod As String) As String

Select Case Tabla
    Case "ESPACIOS"
        If optTipo(0).Value Then
            CodigoResultado = "STD" & cod
        Else
            CodigoResultado = "ESP" & cod
        End If
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






Private Sub Form_Activate()
    tmrBusca.Enabled = False
    mskCodigo.SetFocus
    Select Case Tabla
        Case "CAJEROS"
            mskCodigo.Text = ""
        Case "PROVEEDOR"
            mskCodigo.Text = "A"
        Case "ESPACIOS"
            mskCodigo.Text = ""
            optTipo(0).Visible = True
            optTipo(1).Visible = True
        Case Else
            mskCodigo.Text = ""
    End Select
    
    frmHelpPv.caption = frmHelpPv.caption & " " & Tabla
    
End Sub

Private Sub mskCodigo_Change()
  tmrBusca.Enabled = False
  
  tmrBusca.Enabled = True

End Sub



Private Sub optTipo_Click(Index As Integer)
  tmrBusca.Enabled = False
  
  tmrBusca.Enabled = True
End Sub

Private Sub tmrBusca_Timer()
Dim s_Criterio As String
Dim sql$, rs As Recordset

frmHelpPv.caption = Aplicacion.SeteoProceso(frmHelpPv.caption)

sql$ = ArmarSQL

Call Aplicacion.ObtenerRsDAO(sql$, rs)

If rs.RecordCount > 0 Then
  Spread_CargarGrilla rs, Spread1
Else
  Spread1.MaxRows = 0
End If

Aplicacion.CerrarDAO rs
 
tmrBusca.Enabled = False

frmHelpPv.caption = Aplicacion.SeteoFin()

End Sub




Public Function MuestraHlp(ByRef cod As Variant, ByRef desc As String, ByVal NT As String, ByVal CA As String) As Integer
    
    Tabla = NT
    Seleccion = CA
        
    Me.Show 1
    
    cod = codigo$
    desc = Descripcion$
    
    MuestraHlp = Resp%
    
End Function

